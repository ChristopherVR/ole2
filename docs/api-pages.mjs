export const apiPages = [
  {
    "slug": "models",
    "title": "Document models",
    "kicker": "CHECKED PARSING · TRANSACTIONAL EDITS",
    "description": "Typed editable legacy Office classes, checked format expectations and explicit serialization limits.",
    "intro": "The primary API is parseOle2: a discriminated union of CfbDocument, DocDocument, XlsDocument and PptDocument. Concrete format helpers return checked editable classes.",
    "sections": [
      {
        "title": "Runtime-checked types",
        "body": "The generic is a format key, never an erased document-class cast. parseOle2<K> requires expect: K and validates the bytes.",
        "bullets": [
          "Default kind narrows the union; parseDoc/parseXls/parsePpt return concrete classes.",
          "Wrong formats throw; ambiguous or unsupported automatic decoding returns a CFB view with diagnostics.",
          "expect: cfb explicitly requests container inspection for any valid compound file."
        ]
      },
      {
        "title": "Mutation and serialization",
        "body": "Supported setters validate candidate edits before committing. UnsupportedOle2EditError leaves bytes, dirty, revision and model values unchanged.",
        "bullets": [
          "All classes extend Ole2DocumentBase and own copied input bytes.",
          "serialize returns a detached preserved-byte copy and does not reset dirty state.",
          "Capabilities describe supported operations; a specific target can still be refused.",
          "Opaque data is retained; complete legacy format fidelity and arbitrary edits are not claimed."
        ]
      }
    ],
    "example": "import { parseOle2, parseDoc, parseXls, parsePpt } from '@christophervr/ole2';\n\nconst document = parseOle2(bytes);\nif (document.kind === 'xls') document.sheets[0].cell(1, 0).value = 'Updated';\nconst ppt = parseOle2<'ppt'>(pptBytes, { expect: 'ppt' });\nppt.slides[0].texts[0].text = 'Native title updated';\nconst saved = ppt.serialize();"
  },
  {
    "slug": "word",
    "title": "Legacy Word",
    "kicker": "DOC · WORD 97–2003",
    "description": "Read main-body paragraphs from legacy Word DOC files and apply narrowly guarded paragraph text edits.",
    "intro": "parseDoc returns a DocDocument with stable main-body paragraph handles. Supported text setters preserve opaque binary records and serialize the edited compound file without reconstructing unsupported content.",
    "sections": [
      {
        "title": "Document model",
        "body": "Read paragraph.text and assign supported replacement text, then call serialize(). Successful edits mark dirty and increment revision.",
        "bullets": [
          "Equal-length plain targets retain runs and character-position tables when the encoding fits.",
          "Guarded growth/shrink outside balanced main-story fields shifts field positions while preserving codes, results and flags.",
          "Rejected setters throw UnsupportedOle2EditError before changing bytes or model state."
        ]
      },
      {
        "title": "Explicit limits",
        "body": "The paragraph collection has fixed structure. Complete run, table, field and drawing models are not yet editable.",
        "bullets": [
          "No paragraph insertion/removal or embedded paragraph breaks.",
          "Edits inside field code/result ranges and unsupported CP-dependent features are refused.",
          "Processing budgets and supported container layout limits apply; no fresh-document writer is implied."
        ]
      },
      {
        "title": "Compatibility APIs",
        "body": "readOleDocParagraphs and tryWriteOleDocParagraphEdit remain available. New code should navigate the document model.",
        "bullets": []
      }
    ],
    "example": "import { parseDoc } from '@christophervr/ole2';\n\nconst document = parseDoc(docBytes);\nconsole.log(document.paragraphs[0].text);\ndocument.paragraphs[0].text = 'A longer plain paragraph outside fields.';\nconst saved = document.serialize();"
  },
  {
    "slug": "excel",
    "title": "Legacy Excel",
    "kicker": "XLS · BIFF8",
    "description": "Read whole BIFF8 Excel workbooks, preview a bounded first worksheet and perform guarded numeric and string cell edits.",
    "intro": "parseXls returns an XlsDocument with tab-order sheets and stable cell handles. Workbook styles, cached formulas and merges are read-only snapshots; supported cell value setters preserve existing record structure.",
    "sections": [
      {
        "title": "Document model",
        "body": "Use sheets[index].cell(row, col).value with zero-based coordinates, then serialize(). Numeric and plain-string writes invoke bounded record editors.",
        "bullets": [
          "Existing NUMBER/RK/MULRK numeric values require exact original encoding.",
          "Existing LABELSST/RK/NUMBER/MULRK cells can become plain Unicode strings, including continued SST entries.",
          "Packed siblings, selected XF and unrelated records are preserved; supported BOUNDSHEET/INDEX/DBCELL/ExtSST pointers are updated.",
          "Formula caches remain saved values; recalculationRequired signals that a consuming application must recalculate."
        ]
      },
      {
        "title": "Explicit limits",
        "body": "Capabilities describe an editing surface, not guaranteed eligibility of each target.",
        "bullets": [
          "No missing-cell creation, formula writing or evaluation.",
          "Unsafe pointer-bearing records and allocation layouts are refused atomically.",
          "A selected rich string becomes plain while other shared aliases retain their formatting.",
          "Charts/drawings and general workbook construction are not modeled by this adapter."
        ]
      },
      {
        "title": "Compatibility APIs",
        "body": "readXlsWorkbook, previews and the numeric/string operation functions remain available. New code should navigate the document model.",
        "bullets": []
      }
    ],
    "example": "import { parseXls } from '@christophervr/ole2';\n\nconst document = parseXls(xlsBytes);\nconst cell = document.sheets[0].cell(1, 0);\ncell.value = 'Unicode Ω 日本';\nconst saved = document.serialize();\nconsole.log(document.recalculationRequired);"
  },
  {
    "slug": "powerpoint",
    "title": "Legacy PowerPoint",
    "kicker": "PPT · POWERPOINT 97–2003",
    "description": "Read active legacy PPT slide text, make fixed-length preservation edits, and export a neutral WDeck model.",
    "intro": "parsePpt returns a PptDocument with active slide identities and stable outline/inline text handles. Text atoms are not a complete shape model; unsupported records and streams remain opaque and preserved.",
    "sections": [
      {
        "title": "Document model",
        "body": "Assign slides[index].texts[index].text and serialize() after a supported edit.",
        "bullets": [
          "Replacement text retains the UTF-16 length, original encoding and control/field marker positions.",
          "Active save history is resolved; stale saves, unknown records and surrounding bytes are retained.",
          "Mirrored text and physical atoms shared by multiple active slide positions are refused.",
          "Unsupported edits throw before changing dirty state, revision, model values or bytes."
        ]
      },
      {
        "title": "Explicit limits",
        "body": "Shapes, shape-to-text references, rich runs, notes, masters and animations are not decoded into editable model nodes in this release.",
        "bullets": [
          "No arbitrary text growth or shape/layout mutation is implied.",
          "Encrypted PPT cannot produce a successful checked model.",
          "The separate buildPptFile exporter consumes a neutral WDeck model; source-to-WDeck conversion and rendering remain in the viewer."
        ]
      },
      {
        "title": "Compatibility APIs",
        "body": "readPptSlideTexts, editPptSlideText and record utilities remain available. New existing-file edits should use the document model.",
        "bullets": []
      }
    ],
    "example": "import { parsePpt } from '@christophervr/ole2';\n\nconst document = parsePpt(pptBytes);\ndocument.slides[0].texts[0].text = 'Native title updated';\nconst saved = document.serialize();"
  },
  {
    "slug": "visio",
    "title": "Legacy Visio",
    "kicker": "VSD · VISIO",
    "description": "Validate structural facts in legacy Visio containers without claiming drawing or page parsing.",
    "intro": "`inspectLegacyVisio` checks a legacy VSD compound file’s VisioDocument signature, version, and TrailerStream pointer bounds. It reports structural facts only.",
    "sections": [
      {
        "title": "What works",
        "body": "The inspector validates the header signature, version-dependent pointer layout, trailer type, and trailer range against the declared document size. It returns stream names and trailer metadata when those checks pass.",
        "bullets": [
          "Supports the validated V5 and V6+ pointer layouts.",
          "The generic metadata API can read or update an existing SummaryInformation text property when its value fits the allocated slot.",
          "Invalid or unsupported structures return `undefined` from the inspector."
        ]
      },
      {
        "title": "Known limits",
        "body": "Visio shapes, pages, geometry, and drawing relationships are not decoded or editable. A valid TrailerStream pointer is not a parsed drawing.",
        "bullets": [
          "No page preview or rendering API is included.",
          "No general VSD save or structural edit API is included.",
          "Metadata edits do not change drawing content and do not add missing properties."
        ]
      }
    ],
    "example": "import { inspectLegacyVisio, readLegacyOfficeMetadata } from '@christophervr/ole2';\n\nconst inspection = inspectLegacyVisio(vsdBytes);\nif (inspection) {\n  console.log(inspection.version, inspection.trailerLength);\n  console.log(readLegacyOfficeMetadata(vsdBytes));\n}"
  },
  {
    "slug": "publisher",
    "title": "Legacy Publisher",
    "kicker": "PUB · PUBLISHER",
    "description": "Inspect known legacy Publisher container signatures and versions; page layout is not decoded.",
    "intro": "`inspectLegacyPublisher` validates known Publisher Contents signatures and the additional stream set required for Publisher 2002. This is format identification and container inspection, not publication-page support.",
    "sections": [
      {
        "title": "What works",
        "body": "The inspector recognizes validated Publisher 97/2000 and 2002 container signatures and returns stream names plus the streams used to validate the result.",
        "bullets": [
          "Publisher 2002 requires the Contents, Escher/EscherStm, and Quill/QuillSub/CONTENTS streams.",
          "The generic metadata API can edit an existing SummaryInformation text property when the replacement fits its allocated space.",
          "Unknown or incomplete signatures do not produce a successful inspection."
        ]
      },
      {
        "title": "Known limits",
        "body": "The package does not decode publication pages, text frames, images, or layout and does not provide a Publisher document writer.",
        "bullets": [
          "Inspection results do not imply page content was parsed.",
          "Metadata editing is independent of publication content.",
          "No rendering, page preview, or structural publication edits are available."
        ]
      }
    ],
    "example": "import { inspectLegacyPublisher, writeLegacyOfficeMetadata } from '@christophervr/ole2';\n\nconst inspection = inspectLegacyPublisher(pubBytes);\nif (inspection) {\n  console.log(inspection.version, inspection.validatedStreams);\n  const updated = writeLegacyOfficeMetadata(pubBytes, 'title', 'New title');\n  if (updated === pubBytes) console.log('Metadata was unchanged or unsupported.');\n}"
  },
  {
    "slug": "containers",
    "title": "OLE2 containers",
    "kicker": "MS-CFB · NAMED STREAMS",
    "description": "Read and build OLE2 compound-file containers and safely replace existing nested streams.",
    "intro": "parseOle2 returns a checked document union by default. Explicit expect: cfb selects a CfbDocument container view for any valid compound file; this does not imply Office record semantics are decoded.",
    "sections": [
      {
        "title": "Container model",
        "body": "Use stream(path).bytes to inspect or replace an existing stream by full storage path, then serialize().",
        "bullets": [
          "V3 regular/mini growth, shrink, emptying and transitions preserve hierarchy, unknown streams and directory metadata.",
          "V4 regular-stream resizing validates 4096-byte sectors and directory counts.",
          "New storage entries, external DIFAT expansion and variable-length v4 mini edits remain unsupported."
        ]
      },
      {
        "title": "Detection and diagnostics",
        "body": "Automatic parsing falls back to a CFB view with explicit diagnostics for ambiguous or unsupported Office models. Checked Office expectations throw instead.",
        "bullets": [
          "Container stream edits do not repair offsets inside Office records.",
          "Compatibility entries/getStream remain; returned inspection data is detached from owned bytes.",
          "parseCompoundFile names the original raw reader."
        ]
      }
    ],
    "example": "import { parseOle2 } from '@christophervr/ole2';\n\nconst container = parseOle2(bytes, { expect: 'cfb' });\nconst stream = container.stream(['Custom', 'Data']);\nstream.bytes = new Uint8Array([1, 2, 3]);\nconst saved = container.serialize();"
  }
];
