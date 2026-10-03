# Authored binary Visio v11 fixture

`owned-v11.vsd` is generated entirely by `makeVsdFixture()` in `generated.ts`.
It contains an 8 × 11 inch page, shape 7, text `Hello` followed by LF, a
literal stored transform and a MoveTo/LineTo path. An opaque chunk and a
separate unknown CFB stream exercise preservation. No third-party document
bytes, personal documents, macros or embedded objects are included.

SHA-256: `a3781bcdaadfb48f3e9669808930013197b541af164dd9e44932da6d744ef36e`.
The fixture generator and binary are original project material under the
repository Apache-2.0 license.

Binary layout facts were researched using Apache POI HDGF and libvisio's
public readers. No implementation code from those readers was incorporated.
Microsoft's MS-VSDX documentation describes the modern XML package and is
not a specification for this binary format.

Independent libvisio 0.1.7 callback checks observe the finite page dimensions,
one path and one text object. Replacing `Hello` with `World` changes only the
text callback. Changing stored pinX from 2 to 6 and width from 4 to 5 shifts
the path and text frame by 4 inches and changes the declared text-frame width.
The literal path segment retains its 4-inch spacing; no formula recalculation
or automatic geometry scaling is claimed. Microsoft Visio 16.0.20430.20140
subsequently rejected this authored fixture and its edited variants with
HRESULT -2032466854 (newer or unrecognized version). These are parser-mechanics
fixtures, not native-valid drawing evidence.

`native-visio16-v11.vsd` was generated from a new blank document through the
owned, invisible native Visio COM instance using `generate-native-fixture.ps1`.
It contains an 8 x 11 inch page, three original rectangle/line/label shapes,
the text `Hello` and a Unicode U+03A9 label, six native styles, and no masters
or layers. No source documents, macros, embedded objects or third-party bytes
were used. Author/company metadata is deliberately neutral fixture metadata.
The generator, fixture and harness are original Apache-2.0 project material.

SHA-256: `c6c97822e7bb2cc3e96da9d7d35fe74ac16a7f2c90d19ad4c80967d2896e0974`.

The native baseline and its native save/reopen matched all captured fields.
The model agrees with page dimensions, IDs and stored transform values;
stored text includes a terminal paragraph LF that COM `Text` omits. Version
0.9.0's equal-length text edit relocated its page block and Visio rejected the
output as corrupt (HRESULT -2032466840). The writer now refuses block
relocation; this regression is distinct from successful native write evidence.
No positive master/layer, visual, formula or full-format fidelity is claimed.
Parser self-roundtrips alone are not native fidelity evidence.
