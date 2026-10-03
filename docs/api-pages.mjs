export const apiPages = [
  {
    "slug": "models",
    "title": "Document models",
    "kicker": "CHECKED PARSING · TRANSACTIONAL EDITS",
    "description": "Typed editable legacy Office classes, checked format expectations and explicit serialization limits.",
    "intro": "The primary API is parseOle2: a checked discriminated union of CfbDocument, DocDocument, XlsDocument, PptDocument and VsdDocument. Concrete helpers return preserving editable classes.",
    "sections": [
      {
        "title": "Runtime-checked types",
        "body": "The generic is a format key, never an erased document-class cast. parseOle2<K> requires expect: K and validates the bytes.",
        "bullets": [
          "Default kind narrows the union; parseDoc/parseXls/parsePpt/parseVsd return concrete classes.",
          "Wrong formats throw; ambiguous or unsupported automatic decoding returns a CFB view with diagnostics.",
          "expect: cfb explicitly requests container inspection for any valid compound file.",
          "The new vsd union variant requires exhaustive switch consumers to handle that kind."
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
        "title": "Actual character runs",
        "body": "paragraph.runs maps CLX pieces and CHPX formatting pages to character positions. styleIndex reads the PAPX paragraph style identifier.",
        "bullets": [
          "Direct bold/italic/font-size values are exposed separately from inherited or undecoded styling. Unknown SPRMs remain opaque.",
          "Existing exclusive understood bold/italic operands can be set to absolute values. directFontSizePoints can replace an existing exclusive sprmCHps slot with 1..1638 points in exact half-point increments. Shared blobs, opaque semantics and missing slots refuse.",
          "Run handles expire after edits; reacquire paragraph.runs. Styles and piece PRMs are not resolved."
        ]
      },
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
        "body": "The paragraph collection has fixed structure. Rich runs expose a bounded direct-formatting slice; tables, fields, styles and drawings are preserved without complete editable models.",
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
          "Formula caches remain saved values; recalculationRequired signals that a consuming application must recalculate.",
          "Existing NUMBER/RK/MULRK cells may become boolean/error BOOLERR records; existing BOOLERR values can be replaced. Cell type distinguishes formula, number, string, boolean, error and blank."
        ]
      },
      {
        "title": "Explicit limits",
        "body": "Capabilities describe an editing surface, not guaranteed eligibility of each target.",
        "bullets": [
          "No missing-cell creation, formula writing or evaluation.",
          "Unsafe pointer-bearing records and allocation layouts are refused atomically.",
          "A selected rich string becomes plain while other shared aliases retain their formatting.",
          "Charts/drawings and general workbook construction are not modeled by this adapter.",
          "String-to-BOOLERR and BOOLERR-to-number/string conversions remain unsupported."
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
    "intro": "parsePpt returns a PptDocument with active slides, OfficeArt shape identities, explicit anchors and validated inline text references. Opaque binary content remains preserved.",
    "sections": [
      {
        "title": "Actual shapes and small-anchor edits",
        "body": "slides[index].shapes exposes IDs, kind, flags, names, coordinate spaces and copied geometry snapshots. Bounds use exact master units, eight per PowerPoint point.",
        "bullets": [
          "Unmirrored top-level small-anchor rectangles and text boxes support bounded position and extent setters. Candidate geometry is reparsed before commit.",
          "Shared, grouped, inherited, rotated/flipped, unknown-property and OOXML-mirrored geometry refuses editing.",
          "Group and child coordinates remain local. Large ClientAnchor values are raw with unknown geometry; outline-to-shape text references can remain unresolved."
        ]
      },
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
        "body": "The model exposes a bounded shape and text slice, with explicit resource budgets and target-specific refusal reasons.",
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
    "description": "Read and preserve supported binary Visio version 11 pages, shapes, stored text and literal transforms.",
    "intro": "parseVsd returns a VsdDocument for the supported binary version 11 drawing slice. This uses legacy binary records; Microsoft MS-VSDX describes a different XML format.",
    "sections": [
      {
        "title": "Typed drawing model",
        "body": "pages and shapes expose real IDs, stored page dimensions/scale, explicit transforms, UTF-16 text and MoveTo/LineTo geometry.",
        "bullets": [
          "Version 11 only; unsupported versions fail checked parsing or remain diagnostic CFB inspection in automatic parsing.",
          "Group/parent/master identities and unsupported geometry remain explicit; styles, masters and formulas are not evaluated.",
          "Compressed pointer/record traversal has input, decoded-byte, depth and record budgets."
        ]
      },
      {
        "title": "Preserving edits",
        "body": "Supported shape.text and shape.transform setters serialize edited leaves and pointer ancestors while retaining original unknown record bytes and other CFB streams.",
        "bullets": [
          "Text retains UTF-16 length and control positions; fields, shared or overlapping allocations refuse.",
          "Transform writes require an exclusive top-level literal transform with understood unit tags and no parent/master dependency.",
          "Literal path points are not automatically scaled or recalculated when transform dimensions change."
        ]
      },
      {
        "title": "Evidence and limitations",
        "body": "Authored fixtures passed independent libvisio 0.1.7 callback comparisons and Windows IStorage preservation checks. Native Visio fidelity remains unverified.",
        "bullets": [
          "Earlier binary versions, general ShapeSheet evaluation, styles, rich text, embedded content and general drawing reconstruction remain unsupported.",
          "inspectLegacyVisio still validates the signature/version/TrailerStream pointer for compatible inspection layouts; inspection is separate from drawing decoding.",
          "No native Visio or complete rendering/roundtrip parity claim is made."
        ]
      }
    ],
    "example": "import { parseVsd } from '@christophervr/ole2';\n\nconst drawing = parseVsd(vsdBytes);\nconst shape = drawing.pages[0].shapes[0];\nshape.text = 'World\\n'; // Same UTF-16 length as owned Hello + LF fixture.\nshape.transform = { ...shape.transform, pinX: 6, width: 5 };\nconst saved = drawing.serialize();"
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
