param(
    [Parameter(Mandatory)][string]$InputPath,
    [Parameter(Mandatory)][string]$OutputPath,
    [int]$MaxEntries = 8192,
    [long]$MaxStreamBytes = 67108864,
    [long]$MaxTotalBytes = 134217728
)
# Windows container oracle only: no Office application, macro or object activation.
# https://learn.microsoft.com/en-us/windows/win32/api/coml2api/nf-coml2api-stgopenstorage
# https://learn.microsoft.com/en-us/windows/win32/api/objidl/nn-objidl-istorage
$ErrorActionPreference = 'Stop'
$inputFile = (Resolve-Path -LiteralPath $InputPath).Path
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Cryptography;

namespace NativeCfb {
  [ComImport, Guid("0000000B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  interface IStorage {
    void CreateStream([MarshalAs(UnmanagedType.LPWStr)] string name, int mode, int reserved1, int reserved2, out IStream stream);
    void OpenStream([MarshalAs(UnmanagedType.LPWStr)] string name, IntPtr reserved1, int mode, int reserved2, out IStream stream);
    void CreateStorage([MarshalAs(UnmanagedType.LPWStr)] string name, int mode, int reserved1, int reserved2, out IStorage storage);
    void OpenStorage([MarshalAs(UnmanagedType.LPWStr)] string name, IntPtr priority, int mode, IntPtr exclude, int reserved, out IStorage storage);
    void CopyTo(int count, IntPtr iids, IntPtr exclude, IStorage destination);
    void MoveElementTo([MarshalAs(UnmanagedType.LPWStr)] string name, IStorage destination, [MarshalAs(UnmanagedType.LPWStr)] string newName, int flags);
    void Commit(int flags);
    void Revert();
    void EnumElements(int reserved1, IntPtr reserved2, int reserved3, out IEnumStat enumerator);
    void DestroyElement([MarshalAs(UnmanagedType.LPWStr)] string name);
    void RenameElement([MarshalAs(UnmanagedType.LPWStr)] string oldName, [MarshalAs(UnmanagedType.LPWStr)] string newName);
    void SetElementTimes([MarshalAs(UnmanagedType.LPWStr)] string name, IntPtr creation, IntPtr access, IntPtr modification);
    void SetClass(ref Guid clsid);
    void SetStateBits(int bits, int mask);
    void Stat(out STATSTG stat, int flags);
  }
  [ComImport, Guid("0000000D-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  interface IEnumStat {
    [PreserveSig] int Next(int count, out STATSTG stat, out int fetched);
    void Skip(int count);
    void Reset();
    void Clone(out IEnumStat copy);
  }
  public sealed class Entry {
    public string[] path;
    public int type;
    public long size;
    public string classId;
    public int stateBits;
    public long creationTime;
    public long modificationTime;
    public string sha256;
  }
  public static class Snapshot {
    [DllImport("ole32.dll", CharSet=CharSet.Unicode, PreserveSig=false)]
    static extern void StgOpenStorage(string path, IntPtr priority, int mode, IntPtr exclude, int reserved, out IStorage storage);
    static long Time(FILETIME time) { return ((long)time.dwHighDateTime << 32) | (uint)time.dwLowDateTime; }
    static Entry Record(string[] path, STATSTG stat) {
      return new Entry { path=path, type=stat.type, size=stat.cbSize, classId=stat.clsid.ToString(), stateBits=stat.grfStateBits,
        creationTime=path.Length==0 ? 0 : Time(stat.ctime), modificationTime=path.Length==0 ? 0 : Time(stat.mtime) };
    }
    public static Entry[] Read(string path, int maxEntries, long maxStream, long maxTotal) {
      if (maxEntries<1 || maxStream<0 || maxTotal<0) throw new ArgumentException("Invalid native CFB limits");
      IStorage storage=null;
      try {
        StgOpenStorage(path, IntPtr.Zero, 0x20, IntPtr.Zero, 0, out storage); // READ | SHARE_DENY_WRITE
        storage.Stat(out var stat, 1); // no root filesystem name
        var result=new List<Entry>{Record(Array.Empty<string>(),stat)};
        long total=0;
        Walk(storage, Array.Empty<string>(), result, maxEntries, maxStream, maxTotal, ref total);
        return result.ToArray();
      } finally { if(storage!=null) Marshal.FinalReleaseComObject(storage); }
    }
    static void Walk(IStorage storage, string[] parent, List<Entry> result, int maxEntries, long maxStream, long maxTotal, ref long total) {
      if(parent.Length>32) throw new InvalidOperationException("Native CFB depth exceeds32");
      IEnumStat enumerator=null;
      try {
        storage.EnumElements(0,IntPtr.Zero,0,out enumerator);
        while(true) {
          int hr=enumerator.Next(1,out var stat,out int fetched);
          if(hr<0) Marshal.ThrowExceptionForHR(hr);
          if(fetched==0) break;
          if(fetched!=1 || string.IsNullOrEmpty(stat.pwcsName)) throw new InvalidOperationException("Malformed native enumeration");
          if(result.Count>=maxEntries) throw new InvalidOperationException("Native CFB entry limit");
          var path=new string[parent.Length+1];Array.Copy(parent,path,parent.Length);path[parent.Length]=stat.pwcsName;
          var entry=Record(path,stat);result.Add(entry);
          if(stat.type==1) {
            IStorage child=null;
            try { storage.OpenStorage(stat.pwcsName,IntPtr.Zero,0x10,IntPtr.Zero,0,out child);Walk(child,path,result,maxEntries,maxStream,maxTotal,ref total); }
            finally { if(child!=null) Marshal.FinalReleaseComObject(child); }
          } else if(stat.type==2) {
            if(stat.cbSize<0 || stat.cbSize>maxStream || stat.cbSize>maxTotal-total) throw new InvalidOperationException("Native CFB byte limit");
            total+=stat.cbSize;
            IStream stream=null;IntPtr count=IntPtr.Zero;
            try {
              storage.OpenStream(stat.pwcsName,IntPtr.Zero,0x10,0,out stream);
              stream.Stat(out var actual,1);
              if(actual.cbSize!=stat.cbSize) throw new InvalidOperationException("Native stream size changed");
              count=Marshal.AllocHGlobal(4);byte[] buffer=new byte[65536];long remaining=stat.cbSize;
              using(var hash=IncrementalHash.CreateHash(HashAlgorithmName.SHA256)) {
                while(remaining>0) {
                  int requested=(int)Math.Min(remaining,buffer.Length);Marshal.WriteInt32(count,0);
                  stream.Read(buffer,requested,count);int read=Marshal.ReadInt32(count);
                  if(read<=0 || read>requested) throw new InvalidOperationException("Native stream truncated");
                  hash.AppendData(buffer,0,read);remaining-=read;
                }
                entry.sha256=Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant();
              }
            } finally { if(count!=IntPtr.Zero) Marshal.FreeHGlobal(count);if(stream!=null) Marshal.FinalReleaseComObject(stream); }
          } else throw new InvalidOperationException("Unexpected native storage element type");
        }
      } finally { if(enumerator!=null) Marshal.FinalReleaseComObject(enumerator); }
    }
  }
}
'@
$entries = [NativeCfb.Snapshot]::Read($inputFile, $MaxEntries, $MaxStreamBytes, $MaxTotalBytes)
[ordered]@{
    consumer = 'Windows ole32 IStorage'
    version = [Environment]::OSVersion.Version.ToString()
    evidence = 'native-container-streams-and-storage-metadata'
    officeFidelity = $false
    rootFilesystemTimestampsExcluded = $true
    entries = @($entries | Sort-Object { $_.path -join '/' })
} | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $OutputPath -Encoding utf8
