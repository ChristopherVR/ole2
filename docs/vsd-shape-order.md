# Binary VSD drawing order

`parseVsd(bytes).pages[index].topLevelShapeIds` exposes validated page drawing
order. `shape.childShapeIds` exposes validated child order within an explicit
group owner. Both return copied frozen arrays or `undefined` when unresolved;
`shapeOrderIssue` explains unsupported or ambiguous ordering.

```ts
const page = parseVsd(bytes).pages[0];
const shapesById = new Map(page.shapes.map(shape => [shape.id, shape]));
if (page.topLevelShapeIds) {
  for (const id of page.topLevelShapeIds) inspect(shapesById.get(id)!);
} else {
  console.log(page.shapeOrderIssue);
}
```

The existing `page.shapes` list remains flat and in physical record order for
lookup compatibility. Consumers must use the new validated order metadata for
drawing order, keeping group children in their owner context. Child coordinates
remain local; master/style inheritance and general ShapeSheet evaluation are
unresolved. Order metadata does not grant transform editing support.

ShapeList element IDs are resolved through owner-scoped ShapeId mappings rather
than assumed to equal shape IDs. Duplicate or missing mappings, truncated or
overlapping ranges, ambiguous page owners, invalid parent chains and malformed
lists produce explicit unresolved diagnostics. Aggregate order elements are
bounded to 100,000 before allocation; hierarchy validation is iterative.

The owned native Visio v11 hierarchy fixture has two pages, a group, a master,
custom style, layers, Unicode text and multiple geometry sections. Its native
page orders `[2, 1, 5, 6]` and `[1, 2, 3]`, and group-five child order `[3, 4]`,
match the new metadata. Existing IDs, flat arrays, local stored values and no-op
serialization remain unchanged. Native save/reopen preserves the captured
fixture fields. Hidden secondary geometry stays explicitly unsupported; these
checks establish captured read semantics, not rendered or full-format fidelity.
