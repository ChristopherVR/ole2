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
or automatic geometry scaling is claimed. Native Microsoft Visio validation
was unavailable. Parser self-roundtrips alone are not native fidelity evidence.
