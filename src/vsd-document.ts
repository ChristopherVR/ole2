/** Stored-value binary Visio v11 model; no master/style/formula evaluation. */
import { Ole2DocumentBase, UnsupportedOle2EditError, type Ole2DocumentCapabilities } from './ole2-document-base.js';
import { readVsdDrawing, validateVsdInput, VsdError, type VsdDrawingData, type VsdTransform, type VsdGeometry } from './vsd-reader.js';
import { replaceVsdShapeText, replaceVsdShapeTransform } from './vsd-writer.js';

export class VsdShape {
 #read: () => VsdDrawingData['pages'][number]['shapes'][number]; #editText:(text:string)=>void; #editTransform:(value:VsdTransform)=>void;
 constructor(readonly id: number, read: () => VsdDrawingData['pages'][number]['shapes'][number], editText: (text:string)=>void, editTransform: (value:VsdTransform)=>void) { this.#read=read;this.#editText=editText;this.#editTransform=editTransform;Object.freeze(this); }
 get kind(): 'shape'|'group'|'foreign' { return this.#read().kind; }
 get coordinateSpace(): 'shape-local' { return this.#read().coordinateSpace; }
 get transformCoordinateSpace(): 'parent-local' {return 'parent-local';}
 get parentId():number|undefined{return this.#read().parentId;}
 get masterPageId():number|undefined{return this.#read().masterPageId;}
 get masterShapeId():number|undefined{return this.#read().masterShapeId;}
 get unsupportedGeometry():boolean{return this.#read().unsupportedGeometry;}
 get text(): string | undefined { return this.#read().text; }
 set text(value: string) { this.#editText(value); }
 get transform(): Readonly<VsdTransform> | undefined { const value=this.#read().transform; return value && Object.freeze({...value}); }
 set transform(value: VsdTransform) { this.#editTransform(value); }
 /** Validated drawing order of this owner's explicit children; undefined when unresolved. */
 get childShapeIds():readonly number[]|undefined {const ids=this.#read().childShapeIds;return ids&&Object.freeze([...ids]);}
 get shapeOrderIssue():string|undefined{return this.#read().shapeOrderIssue;}
 get geometry(): readonly VsdGeometry[] { return Object.freeze(this.#read().geometry.map(row=>Object.freeze({...row}))); }
}
export class VsdPage {
 readonly shapes: readonly VsdShape[]; #read:()=>VsdDrawingData['pages'][number];
 constructor(readonly id: number, shapes: VsdShape[], read: ()=>VsdDrawingData['pages'][number]) {this.#read=read;this.shapes=Object.freeze(shapes); Object.freeze(this); }
 /** Validated top-level drawing order. The existing shapes array remains flat. */
 get topLevelShapeIds():readonly number[]|undefined {const ids=this.#read().topLevelShapeIds;return ids&&Object.freeze([...ids]);}
 get shapeOrderIssue():string|undefined{return this.#read().shapeOrderIssue;}
 get background(): boolean { return this.#read().background; }
 get width(): number | undefined { return this.#read().width; }
 get height(): number | undefined { return this.#read().height; }
 /** Stored internal inches multiply by scale for rendered page coordinates. */
 get scale(): number | undefined {return this.#read().scale;}
}
const CAPABILITIES: Ole2DocumentCapabilities = Object.freeze({read:Object.freeze(['v11-pages','explicit-shape-transform','utf16-shape-text','move-to-line-to']),write:Object.freeze(['equal-length-shape-text','literal-shape-transform']),limitations:Object.freeze(['Version 11 only; stored values without master/style/formula evaluation.','Edits require an exact stored-block fit at the original offset; relocation is refused. Bounded compression preserves exact decoded length and supports at most 8 MiB decoded leaf/allocation, with bounded token-fit search and no added decoded bytes.','No added/deleted shapes, text length changes or text field edits.','Shared/overlapping VSD blocks cannot be edited; native evidence covers only declared fixture edits, not full Visio fidelity.'])});
export class VsdDocument extends Ole2DocumentBase {
 #drawing: VsdDrawingData; #pages: readonly VsdPage[];
 get kind(): 'vsd' {return 'vsd';}
 get version(): 11 {return 11;}
 get pages(): readonly VsdPage[] {return this.#pages;}
 get capabilities(): Ole2DocumentCapabilities {return CAPABILITIES;}
 constructor(input: Uint8Array) {
  super(validateVsdInput(input)); this.#drawing=readVsdDrawing(this.getBytes());
  this.#pages=Object.freeze(this.#drawing.pages.map(page=>new VsdPage(page.id,page.shapes.map(shape=>new VsdShape(shape.id,()=>this.#shape(page.id,shape.id),text=>this.#edit(()=>replaceVsdShapeText(this.getBytes(),page.id,shape.id,text),'invalid-text'),value=>this.#editTransform(page.id,shape.id,value))),()=>this.#page(page.id))));
 }
 #page(id:number) {const page=this.#drawing.pages.find(p=>p.id===id);if(!page)throw new VsdError('stale-page');return page;}
 #shape(pageId:number,id:number) {const shape=this.#page(pageId).shapes.find(s=>s.id===id);if(!shape)throw new VsdError('stale-shape');return shape;}
 #editTransform(pageId:number,shapeId:number,value:VsdTransform):void {
  this.#edit(()=>replaceVsdShapeTransform(this.getBytes(),pageId,shapeId,value),'invalid-transform');
 }
 #edit(prepare:()=>Uint8Array,invalidReason:'invalid-text'|'invalid-transform'):void {
  try {const revision=this.revision,candidate=prepare();if(this.revision!==revision)throw new VsdError('reentrant-edit');this.#commit(candidate);}
  catch(error){
   if(error instanceof UnsupportedOle2EditError)throw error;
   throw new UnsupportedOle2EditError(error instanceof VsdError?error.reason:error instanceof TypeError?invalidReason:'edit-failed');
  }
 }
 #commit(candidate:Uint8Array):void {
  const drawing=readVsdDrawing(candidate);
  if(drawing.pages.length!==this.#pages.length || this.#pages.some((p,i)=>{
   const next=drawing.pages[i]!;return next.id!==p.id || next.shapes.length!==p.shapes.length || p.shapes.some((s,j)=>next.shapes[j]!.id!==s.id || next.shapes[j]!.kind!==s.kind);
  }))throw new VsdError('changed-model-structure');
  this.commitBytes(candidate);this.#drawing=drawing;
 }
}
