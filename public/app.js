import { STORAGE_KEY, createWorkspace, loadWorkspace, fieldValue, valueOf, normalizeValue, updateRecord, projectRecords, gridRows, leafGroups, groupFields, sortRules, makeView, saveField, deleteField, newTable, defaultsFor, TYPES, VIEW_TYPES, editable, History, dateKey, relativeDate, addDays, uid } from './model.js';

const S = GC.Spread.Sheets;
const $ = id => document.getElementById(id);
let workspace = loadWorkspace(localStorage);
let state = workspace.tables.find(t => t.id === workspace.activeTableId);
let history = new History(workspace);
let editingId = null;
let editingFieldId = null;
let pendingBefore = null;
let editFrame = 0;
let drag = null;
let ghost = null;
let dropColumn = -1;
let storageAvailable = true;

const palette = {
  '待开始': ['#f2eef4', '#7a6b80'], '进行中': ['#f8ebf3', '#941966'], '已完成': ['#e6f5ec', '#319461'], '已暂停': ['#edf0f4', '#8b94a4'],
  '重要紧急': ['#ffeded', '#de5e67'], '紧急不重要': ['#fff0e2', '#ca8642'], '重要不紧急': ['#fff6cf', '#b28c1c'], '不紧急不重要': ['#e5f5ee', '#389b76'],
  '已延期': ['#ffeded', '#d65b65'], '正常': ['#eef7f1', '#419867'], '延期完成': ['#fff0e2', '#ca8642']
};
const personColors = ['#c2688f', '#8f74c9', '#48a89b', '#d59a55'];
const fieldIcons = { text: 'A', number: '#', date: '▦', select: '◉', multiselect: '☷', person: '●', progress: '%', rating: '★', button: '↗', creator: '♙', computed: 'ƒ' };
const view = () => state.views.find(v => v.id === state.activeView) || state.views[0];
const findRecord = id => state.records.find(r => r.id === id);
const groupTitle = key => key || '未设置';
const colorFor = value => palette[value] || ['#f6eaf1', '#8a3c6b'];

// Workbook remains the rendering host in all five view types.
const spread = new S.Workbook($('ss'), {
  sheetCount: 1, tabStripVisible: false, newTabVisible: false,
  allowUserDragDrop: false, allowUserDragFill: false, allowUserResize: false,
  allowUndo: false, showHorizontalScrollbar: true, showVerticalScrollbar: true,
  scrollbarShowMax: false
});
const sheet = spread.getActiveSheet();
// 像素滚动：所有视图共用同一个 Workbook，每次滚动 10 像素。
spread.options.scrollByPixel = true;
spread.options.scrollPixel = 10;
sheet.name('任务');
sheet.options.rowHeaderVisible = false;
sheet.options.colHeaderVisible = false;
sheet.options.gridline = { showVerticalGridline: false, showHorizontalGridline: false };
sheet.options.isProtected = true;
sheet.options.selectionBorderColor = '#941966';
sheet.options.selectionBackColor = 'rgba(148, 25, 102, 0.06)';
sheet.options.frozenlineColor = '#e4e9f1';
sheet.defaults.rowHeight = 38;
sheet.setDefaultStyle(Object.assign(new S.Style(), { font: '12px "Segoe UI", "Microsoft YaHei", sans-serif', foreColor: '#414b60', backColor: '#ffffff', locked: true, vAlign: S.VerticalAlign.center }));

function clippedText(ctx, text, x, y, maxWidth) {
  let result = String(text ?? '');
  if (ctx.measureText(result).width > maxWidth) {
    while (result.length && ctx.measureText(result + '…').width > maxWidth) result = result.slice(0, -1);
    result += '…';
  }
  ctx.fillText(result, x, y);
}

function rounded(ctx, x, y, width, height, radius, fill, stroke) {
  ctx.beginPath();
  ctx.roundRect(x, y, Math.max(0, width), Math.max(0, height), radius);
  if (fill) { ctx.fillStyle = fill; ctx.fill(); }
  if (stroke) { ctx.strokeStyle = stroke; ctx.lineWidth = 1; ctx.stroke(); }
}

function pill(ctx, text, x, y, maxWidth = 200) {
  const [background, foreground] = colorFor(text);
  ctx.font = '11px "Segoe UI", "Microsoft YaHei", sans-serif';
  ctx.textAlign = 'left';
  const width = Math.min(maxWidth, ctx.measureText(String(text)).width + 20);
  rounded(ctx, x, y, width, 21, 10, background);
  ctx.fillStyle = foreground;
  ctx.textBaseline = 'middle';
  clippedText(ctx, text, x + 10, y + 10.5, width - 17);
  return width;
}

function person(ctx, text, x, y) {
  const field = state.fields.find(f => f.id === 'owner');
  const i = Math.max(0, (field?.options||workspace.members).indexOf(text));
  ctx.fillStyle = personColors[i % personColors.length];
  ctx.beginPath(); ctx.arc(x + 10, y, 10, 0, 2 * Math.PI); ctx.fill();
  ctx.font = '9px "Microsoft YaHei", sans-serif'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillStyle = 'white';
  ctx.fillText(String(text).slice(0, 1), x + 10, y);
  ctx.textAlign = 'left'; ctx.font = '12px "Microsoft YaHei", sans-serif'; ctx.fillStyle = '#59647a';
  ctx.fillText(text || '未分配', x + 29, y);
}

class FieldCell extends S.CellTypes.Base {
  constructor(field) { super(); this.field = field; }
  paint(ctx, value, x, y, w, h, style, context) {
    if (!ctx) return;
    ctx.save(); ctx.beginPath(); ctx.rect(x, y, w, h); ctx.clip();
    const text = Array.isArray(value) ? value.join('、') : String(value ?? '');
    if (this.field.type === 'person' && text) person(ctx, text, x + 12, y + h / 2);
    else if (['select', 'computed', 'person', 'multiselect'].includes(this.field.type) && text) pill(ctx, text, x + 10, y + (h - 21) / 2, w - 18);
    else if (this.field.type === 'progress') {
      const progress = Math.max(0, Math.min(100, Number(value) || 0));
      rounded(ctx, x + 10, y + h / 2 - 4, w - 42, 8, 4, '#edf0f5');
      rounded(ctx, x + 10, y + h / 2 - 4, (w - 42) * progress / 100, 8, 4, '#941966');
      ctx.font = '10px "Segoe UI", sans-serif'; ctx.fillStyle = '#68748b'; ctx.textAlign = 'right'; ctx.textBaseline = 'middle'; ctx.fillText(`${progress}%`, x + w - 10, y + h / 2);
    } else if (this.field.type === 'rating') {
      ctx.font = '14px "Segoe UI", sans-serif'; ctx.fillStyle = '#dba943'; ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.fillText('★'.repeat(Number(value) || 0) + '☆'.repeat(5 - (Number(value) || 0)), x + 10, y + h / 2);
    } else if (this.field.type === 'button') {
      rounded(ctx, x + 10, y + h / 2 - 11, Math.min(w - 20, 88), 22, 5, '#f8ebf3', '#e8d9e1');
      ctx.font = '11px "Microsoft YaHei", sans-serif'; ctx.fillStyle = '#7d2a5e'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillText(String(value), x + 54, y + h / 2);
    }
    else {
      ctx.font = this.field.id === 'title' ? '12px "Microsoft YaHei", sans-serif' : '11px "Microsoft YaHei", sans-serif';
      ctx.fillStyle = this.field.id === 'note' ? '#8a93a4' : '#465168'; ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
      let display=text;
      if(this.field.type==='number'&&value!==''&&value!=null){const options={integer:{maximumFractionDigits:0},decimal:{minimumFractionDigits:2,maximumFractionDigits:2},percent:{style:'percent',maximumFractionDigits:2},currency:{style:'currency',currency:'CNY'}};display=new Intl.NumberFormat('zh-CN',options[this.field.format]||{maximumFractionDigits:10}).format(Number(value));}
      clippedText(ctx, display, x + 12, y + h / 2, w - 24);
    }
    ctx.restore();
  }
  // Native input/select editors keep keyboard and IME editing on the DataRange path.
  createEditorElement() {
    const root = document.createElement('div');
    root.setAttribute('gcUIElement', 'gcEditingInput');
    const input = document.createElement(['select','person','multiselect'].includes(this.field.type) ? 'select' : 'input');
    input.setAttribute('gcUIElement', 'gcEditingInput');
    input.setAttribute('aria-label', this.field.name);
    if (['select','person'].includes(this.field.type)) {
      input.add(new Option('未设置', ''));
      (this.field.options||[]).forEach(value => input.add(new Option(value, value)));
    } else if (this.field.type === 'multiselect') { input.multiple=true; (this.field.options||[]).forEach(value => input.add(new Option(value,value))); }
    else input.type = this.field.type === 'date' ? 'date' : ['number','progress','rating'].includes(this.field.type) ? 'number' : 'text';
    Object.assign(input.style, { width: '100%', height: '100%', border: '0', outline: 'none', padding: '0 9px', font: '12px "Microsoft YaHei", sans-serif', background: 'white', color: '#34415d', boxSizing: 'border-box' });
    if (input.type === 'number') { input.step = this.field.type === 'rating' ? '1' : 'any'; if(this.field.type==='progress'){input.min='0';input.max='100';} if(this.field.type==='rating'){input.min='0';input.max='5';} }
    root.append(input);
    return root;
  }
  setEditorValue(root, value) { const input=root.firstElementChild; if(input.multiple){const values=Array.isArray(value)?value:[];[...input.options].forEach(o=>o.selected=values.includes(o.value));}else input.value = value ?? ''; }
  getEditorValue(root) { const input=root.firstElementChild; return input.multiple?[...input.selectedOptions].map(o=>o.value):input.value; }
  updateEditor(root, style, rect) { Object.assign(root.style, { width: `${rect.width}px`, height: `${rect.height}px`, background: 'white' }); }
  focus(root) { root.firstElementChild.focus(); }
  isReservedKey(e) { return this.field.type === 'select' && [38, 40].includes(e.keyCode); }
  isImeAware() { return true; }
}

class LayoutCell extends S.CellTypes.Base {
  paint(ctx, value, x, y, w, h) {
    if (!ctx || !value) return;
    ctx.save(); ctx.beginPath(); ctx.rect(x, y, w, h); ctx.clip();
    ctx.textBaseline = 'middle'; ctx.textAlign = 'left';
    if (value.kind === 'group') {
      ctx.font = '12px "Microsoft YaHei", sans-serif'; ctx.fillStyle = '#8b95a5';
      ctx.fillText(view().collapsed.includes(value.key) ? '›' : '⌄', x + 15, y + h / 2 - 1);
      const width = pill(ctx, value.label||groupTitle(value.key), x + 36, y + (h - 21) / 2);
      ctx.font = '10px "Microsoft YaHei", sans-serif'; ctx.fillStyle = '#9ba4b4';
      ctx.fillText(`${value.count} 条记录`, x + 49 + width, y + h / 2);
    } else if (value.kind === 'header') {
      ctx.font = '11px "Microsoft YaHei", sans-serif'; ctx.fillStyle = '#939dae';
      ctx.fillText(value.icon || '', x + 12, y + h / 2);
      ctx.fillStyle = '#6e798e'; clippedText(ctx, value.text, x + 33, y + h / 2, w - 40);
    } else if (value.kind === 'add') {
      ctx.font = '17px "Segoe UI", sans-serif'; ctx.fillStyle = '#aab5c7'; ctx.fillText('+', x + 15, y + h / 2);
      ctx.font = '11px "Microsoft YaHei", sans-serif'; ctx.fillStyle = '#a0aabd'; ctx.fillText('添加记录', x + 39, y + h / 2);
    } else if (value.kind === 'empty') {
      ctx.font = '12px "Microsoft YaHei", sans-serif'; ctx.fillStyle = '#98a2b4'; ctx.fillText('没有符合条件的任务', x + 20, y + h / 2);
    }
    ctx.restore();
  }
  format(value) { return value?.text || value?.key || ''; }
}

function paintSummary(ctx,record,x,y,w){
  const fields=view().columns.map(id=>state.fields.find(f=>f.id===id)).filter(f=>f&&f.id!=='title'&&f.type!=='button').slice(0,4);
  fields.forEach((field,index)=>{const cy=y+index*24;ctx.font='10px "Microsoft YaHei", sans-serif';ctx.fillStyle='#949eaf';ctx.textAlign='left';ctx.textBaseline='middle';clippedText(ctx,field.name,x,cy+12,72);new FieldCell(field).paint(ctx,valueOf(record,field),x+72,cy,w-72,24,{});});
}

class CardCell extends S.CellTypes.Base {
  paint(ctx, value, x, y, w, h, style) {
    if (!ctx) return;
    ctx.save(); ctx.beginPath(); ctx.rect(x, y, w, h); ctx.clip();
    ctx.fillStyle = style.backColor || '#f5f7fb'; ctx.fillRect(x, y, w, h);
    ctx.textAlign = 'left';
    if (!value) { ctx.restore(); return; }
    if(value.kind==='header') {ctx.font='10px "Microsoft YaHei", sans-serif';ctx.fillStyle='#69758b';ctx.textBaseline='middle';ctx.textAlign='left';clippedText(ctx,value.text,x+8,y+h/2,w-16);ctx.restore();return;}
    if (value.kind === 'boardHeader') {
      const [bg, fg] = colorFor(value.key);
      ctx.fillStyle = fg; ctx.beginPath(); ctx.arc(x + 22, y + h / 2, 4, 0, 2 * Math.PI); ctx.fill();
      ctx.font = '600 13px "Microsoft YaHei", sans-serif'; ctx.fillStyle = '#4d5a72'; ctx.textBaseline = 'middle';
      clippedText(ctx, groupTitle(value.key), x + 35, y + h / 2, w - 100);
      rounded(ctx, x + w - 75, y + 19, 23, 21, 5, '#e8edf4');
      ctx.font = '10px "Segoe UI", sans-serif'; ctx.textAlign = 'center'; ctx.fillStyle = '#909bae'; ctx.fillText(value.count, x + w - 63.5, y + 29.5);
      ctx.textAlign = 'left'; ctx.font = '20px "Segoe UI", sans-serif'; ctx.fillStyle = '#9ca8bb'; ctx.fillText('+', x + w - 35, y + h / 2);
    } else if (value.kind === 'card') {
      const record = value.record;
      rounded(ctx, x + 12, y + 3, w - 24, h - 12, 7, '#fff', '#e4e9f1');
      ctx.font = '600 12px "Microsoft YaHei", sans-serif'; ctx.fillStyle = '#37435c'; ctx.textBaseline = 'middle';
      clippedText(ctx, record.title, x + 27, y + 29, w - 57);
      paintSummary(ctx,record,x+27,y+47,w-54);
    } else if (value.kind === 'boardAdd') {
      ctx.font = '12px "Microsoft YaHei", sans-serif'; ctx.fillStyle = '#9aa6b9'; ctx.textBaseline = 'middle'; ctx.fillText('+  添加任务', x + 25, y + 27);
    } else if (value.kind === 'calendarDay') {
      ctx.fillStyle=value.outside?'#f7f8fb':'#fff';ctx.fillRect(x,y,w,h);ctx.fillStyle='#f8f9fc';ctx.fillRect(x,y,w,28);ctx.font='11px "Microsoft YaHei", sans-serif';ctx.fillStyle=value.outside?'#b3bbc9':'#778197';ctx.textBaseline='middle';ctx.fillText(value.label,x+10,y+14);ctx.fillText('+',x+w-16,y+14);
      if(value.today){ctx.fillStyle='#941966';ctx.beginPath();ctx.arc(x+w-18,y+17,4,0,Math.PI*2);ctx.fill();}
      value.records.slice(0,2).forEach((record,i)=>{const cy=y+33+i*25;rounded(ctx,x+5,cy,w-10,21,4,colorFor(record.status)[0]);ctx.fillStyle=colorFor(record.status)[1];ctx.fillRect(x+5,cy,3,21);ctx.font='10px "Microsoft YaHei", sans-serif';ctx.fillStyle='#465168';ctx.textAlign='left';clippedText(ctx,record.title,x+12,cy+11,w-22);});
      if(value.records.length>2){ctx.font='10px "Microsoft YaHei", sans-serif';ctx.fillStyle='#8a5172';ctx.fillText(`查看全部 ${value.records.length} 项`,x+8,y+h-10);}
    } else if (value.kind === 'gantt') {
      const start=value.record[view().startField]||value.record[view().endField],end=value.record[view().endField]||start;
      ctx.fillStyle=value.weekend?'#f7f9fc':'#fff';ctx.fillRect(x,y,w,h);
      if(start&&end&&value.day>=start&&value.day<=end){ctx.fillStyle='#dfb2c9';ctx.fillRect(x,y+h/2-9,w,18);if(value.day===start||value.day===end){ctx.fillStyle='#8f2f65';ctx.fillRect(value.day===start?x+2:x+w-5,y+h/2-5,3,10);}}
      else {ctx.fillStyle=value.weekend?'#fafbfd':'#fff';ctx.fillRect(x,y,w,h);}
      if(value.today){ctx.fillStyle='#dc5965';ctx.fillRect(x,y,1,h);}
    } else if (value.kind === 'gallery') {
      rounded(ctx,x+8,y+6,w-16,h-12,7,'#fff','#e4e9f1');
      ctx.fillStyle=colorFor(value.record.status)[0];ctx.fillRect(x+9,y+7,w-18,Math.min(78,h*.34));
      ctx.font='600 12px "Microsoft YaHei", sans-serif';ctx.fillStyle='#37435c';ctx.textBaseline='top';ctx.textAlign='left';clippedText(ctx,value.record.title,x+19,y+98,w-38);
      paintSummary(ctx,value.record,x+19,y+122,w-38);
    }
    ctx.restore();
  }
  format(value) { return value?.record?.title || value?.key || ''; }
}

// A single provider renders different projections of the same state.records.
class TaskDataProvider {
  constructor() {
    this.typeName = 'TaskDataProvider';
    this.layoutCell = new LayoutCell();
    this.cardCell = new CardCell();
  }
  rebuild() {
    const current = view();
    this.fields = current.columns.map(id => state.fields.find(f => f.id === id)).filter(Boolean);
    if (!this.fields.some(f => f.id === 'title')) this.fields.unshift(state.fields[0]);
    this.cells = new Map(this.fields.map(f => [f.id, new FieldCell(f)]));
    this.records = projectRecords(state,current);
    this.groups = current.type === 'board' ? leafGroups(state,current) : [];
    this.rows = current.type === 'grid' ? gridRows(state, current) : [];
    if(current.type==='calendar') {
      const [year,month]=current.month.split('-').map(Number),first=new Date(year,month-1,1),offset=(first.getDay()+6)%7;
      this.calendarStart=relativeDate(-offset,first);this.calendarDays=Array.from({length:42},(_,i)=>addDays(this.calendarStart,i));
      this.rowCount=7;this.colCount=7;
    } else if(current.type==='gantt') { this.rowCount=Math.max(1,this.records.length+1);this.colCount=44; }
    else if(current.type==='gallery') { this.galleryColumns=Math.max(1,Math.floor(($('ss').clientWidth||1000)/250));this.rowCount=Math.max(1,Math.ceil(this.records.length/this.galleryColumns));this.colCount=this.galleryColumns; }
    else {
      this.rowCount = current.type === 'grid' ? this.rows.length : Math.max(1, ...this.groups.map(g => g.records.length)) + 2;
      this.colCount = current.type === 'grid' ? this.fields.length + 2 : Math.max(1, this.groups.length);
    }
    this.styleCache = new Map();
  }
  recordAt(row, col) {
    if (view().type === 'grid') return this.rows[row]?.kind === 'record' ? findRecord(this.rows[row].id) : null;
    if(view().type==='calendar') return row>0?this.records.filter(r=>(r[view().dateField]||'')===this.calendarDays[(row-1)*7+col])[0]||null:null;
    if(view().type==='gantt') return this.records[row-1]||null;
    if(view().type==='gallery') return this.records[row*this.galleryColumns+col]||null;
    return this.groups[col]?.records[row - 1] || null;
  }
  getValue(row, col) {
    if(view().type==='calendar') {
      if(row>=7||col>=7)return null;
      if(row===0)return {kind:'header',text:['周一','周二','周三','周四','周五','周六','周日'][col]};
      const day=this.calendarDays[(row-1)*7+col],dateField=state.fields.find(f=>f.id===view().dateField);
      return {kind:'calendarDay',label:day.slice(-2),today:day===dateKey(),outside:day.slice(0,7)!==view().month,records:this.records.filter(r=>r[dateField?.id]===day)};
    }
    if(view().type==='gantt') {
      if(row>=this.rowCount||col>=this.colCount)return null;
      if(row===0)return {kind:'header',text:col<2?['任务','负责人'][col]:addDays(view().timelineStart,col-2).slice(-2)};
      const record=this.records[row-1];
      if(!record)return null;
      if(col===0)return {kind:'header',text:record.title||''};
      if(col===1)return {kind:'header',text:record.owner||''};
      return {kind:'gantt',record,day:addDays(view().timelineStart,col-2),today:addDays(view().timelineStart,col-2)===dateKey(),weekend:[0,6].includes(new Date(addDays(view().timelineStart,col-2)+'T00:00:00').getDay())};
    }
    if(view().type==='gallery') {if(col>=this.galleryColumns||row>=this.rowCount)return null;const record=this.records[row*this.galleryColumns+col];return record?{kind:'gallery',record}:null;}
    if (view().type === 'board') {
      const group = this.groups[col];
      if (!group) return null;
      if (row === 0) return { kind: 'boardHeader', key: group.labels.join(' / '), count: group.records.length };
      const record = group.records[row - 1];
      if (record) return { kind: 'card', record };
      return row === group.records.length + 1 ? { kind: 'boardAdd' } : null;
    }
    const item = this.rows[row];
    if (!item) return null;
    if (item.kind === 'header') {
      if (col === 0) return '#';
      if (col === this.colCount - 1) return '+';
      const field = this.fields[col - 1];
      if (!field) return null;
      return { kind: 'header', icon: fieldIcons[field.type], text: field.name };
    }
    if (item.kind === 'record') {
      if (col === 0) return item.index;
      const field = this.fields[col - 1];
      return field ? valueOf(findRecord(item.id), field) : null;
    }
    return col === 0 && ['group', 'add', 'empty'].includes(item.kind) ? item : null;
  }
  getStyle(row, col) {
    const board = view().type === 'board';
    const special=['calendar','gantt','gallery'].includes(view().type);
    if(special){const style=new S.Style();style.locked=true;style.backColor='#fff';style.borderBottom=new S.LineBorder('#e9edf3',S.LineStyle.thin);style.borderRight=new S.LineBorder('#e9edf3',S.LineStyle.thin);style.cellType=view().type==='calendar'||view().type==='gantt'?this.cardCell:this.cardCell;return style;}
    const item = this.rows[row];
    const field = this.fields[col - 1];
    const key = board ? `board:${col}` : `${item?.kind}:${col}`;
    if (this.styleCache.has(key)) return this.styleCache.get(key);
    const style = new S.Style();
    style.locked = true; style.vAlign = S.VerticalAlign.center;
    style.backColor = '#fff'; style.foreColor = '#465168';
    style.font = '11px "Microsoft YaHei", sans-serif';
    if (board) {
      style.backColor = col === dropColumn ? '#f7e9f1' : '#f7f4f6';
      style.cellType = this.cardCell;
      style.borderRight = new S.LineBorder('#fff', S.LineStyle.medium);
    } else {
      style.borderBottom = new S.LineBorder('#edf0f5', S.LineStyle.thin);
      if (item?.kind === 'record') {
        if (field) { style.cellType = this.cells.get(field.id); style.locked = !editable(field); }
        if (col === 0) { style.hAlign = S.HorizontalAlign.center; style.foreColor = '#a4adbd'; }
        style.borderRight = new S.LineBorder('#edf0f5', S.LineStyle.thin);
      } else if (item?.kind === 'header') {
        style.backColor = '#f8f9fc';
        style.borderRight = new S.LineBorder('#e8ecf2', S.LineStyle.thin);
        if (field) style.cellType = this.layoutCell;
        else { style.hAlign = S.HorizontalAlign.center; style.foreColor = '#9ba7b8'; }
      } else if (item?.kind === 'spacer') { style.backColor = '#fafbfd'; }
      else style.cellType = this.layoutCell;
    }
    this.styleCache.set(key, style);
    return style;
  }
  getSpans(row, col, rowCount, colCount) {
    if (view().type !== 'grid') return [];
    return this.rows.flatMap((item, i) => ['group', 'add', 'empty'].includes(item.kind) && i >= row && i < row + rowCount ? [new S.Range(i, 0, 1, this.colCount)] : []);
  }
  setValue(row, col, value, changes) {
    const item = this.rows[row];
    const field = this.fields[col - 1];
    if (view().type !== 'grid' || item?.kind !== 'record' || !field || !editable(field)) return true;
    try {
      // Keep this projection stable until a paste/edit batch finishes.
      const normalized = normalizeValue(field, value);
      pendingBefore ||= history.snapshot();
      updateRecord(state, item.id, { [field.id]: normalized });
      if (!editFrame) editFrame = requestAnimationFrame(flushEdits);
    } catch (error) { console.warn(error.message); }
    return true; // The provider owns the value; do not write a second copy to Worksheet.
  }
  onMouseDown(row, col, event) {
    if (event.button !== 0) return false;
    if (view().type === 'board') {
      const group = this.groups[col];
      if (!group) return true;
      if (row === 0) {
        const rect = sheet.getCellRect(0, col);
        const host = $('ss').getBoundingClientRect();
        if (event.clientX - host.left > rect.x + rect.width - 48) openRecord(null, group.values);
        return true;
      }
      const record = this.recordAt(row, col);
      if (record) drag = { id: record.id, x: event.clientX, y: event.clientY, sourceColumn: col, moved: false };
      else if (row === group.records.length + 1) openRecord(null, group.values);
      return true;
    }
    if(view().type==='calendar'){
      if(row<1)return true;
      const day=this.calendarDays[(row-1)*7+col],value=this.getValue(row,col),rect=sheet.getCellRect(row,col),host=$('ss').getBoundingClientRect(),dy=event.clientY-host.top-rect.y,dx=event.clientX-host.left-rect.x;
      if(!view().dateField){alert('请先添加日期字段并在日历中选择');return true;}
      if(dy<28&&dx>rect.width-28){openRecord(null,{[view().dateField]:day});return true;}
      const index=Math.floor((dy-33)/25),record=dy>=33&&index<2?value.records[index]:null;
      if(record)drag={kind:'calendar',id:record.id,field:view().dateField,x:event.clientX,y:event.clientY,moved:false};
      else if(value.records.length)showRecordList(day,value.records,{[view().dateField]:day});
      else openRecord(null,{[view().dateField]:day});
      return true;
    }
    if(view().type==='gantt'){
      const record=this.recordAt(row,col);if(!record)return true;
      const start=record[view().startField],end=record[view().endField],day=addDays(view().timelineStart,col-2),rect=sheet.getCellRect(row,col),host=$('ss').getBoundingClientRect(),dx=event.clientX-host.left-rect.x;
      if(col>=2&&start&&end&&start<=day&&day<=end&&view().startField!==view().endField)drag={kind:'gantt',mode:day===start&&dx<10?'start':day===end&&dx>rect.width-10?'end':'move',id:record.id,start,end,startField:view().startField,endField:view().endField,sourceColumn:col,x:event.clientX,y:event.clientY,moved:false};
      else openRecord(record.id);return true;
    }
    if(view().type==='gallery'){const record=this.recordAt(row,col);if(record)openRecord(record.id);return true;}
    const item = this.rows[row];
    if (item?.kind === 'group') {
      sheet.endEdit();
      if (pendingBefore) flushEdits();
      const collapsed = view().collapsed;
      view().collapsed = collapsed.includes(item.key) ? collapsed.filter(k => k !== item.key) : [...collapsed, item.key];
      render(); save(); return true;
    }
    if (item?.kind === 'add') { openRecord(null, item.values || (groupFields(view())[0] ? { [groupFields(view())[0]]: item.key } : {})); return true; }
    if (row === 0 && col === this.colCount - 1) { openNewField(); return true; }
    if (item?.kind === 'record' && col === 0) { openRecord(item.id); return true; }
    if(item?.kind==='record'&&this.fields[col-1]?.type==='button'){const field=this.fields[col-1];if(field.action==='complete'){const status=state.fields.find(f=>f.id==='status');if(status?.options?.includes('已完成'))commit('快捷完成任务',()=>updateRecord(state,item.id,{status:'已完成'}));else alert('当前表需要包含“已完成”选项的状态字段');}else openRecord(item.id);return true;}
    return false;
  }
  onDoubleClick() { return view().type === 'board'; }
  onContextMenu() { return true; }
  onClear() { return false; }
  toJSON() { return { typeName: this.typeName }; }
}

const provider = new TaskDataProvider();
provider.rebuild();
const dataRange = sheet.dataRanges.add('task-view', provider, new S.Range(0, 0, provider.rowCount, provider.colCount));

function save() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(workspace));
    storageAvailable = true;
  } catch (error) { storageAvailable = false; console.warn('本地保存失败', error); }
  $('storage-status').textContent = storageAvailable ? '保存在此浏览器' : '本次更改未保存';
  document.querySelector('.storage-dot').style.background = storageAvailable ? '#46b481' : '#d76868';
}

function flushEdits() {
  cancelAnimationFrame(editFrame); editFrame = 0;
  if (pendingBefore) { history.record(pendingBefore); pendingBefore = null; }
  render(); save();
}

function commit(label, action) {
  if(!sheet.endEdit())return;
  if (pendingBefore) flushEdits();
  const before = history.snapshot();
  action();
  history.record(before);
  render(); save();
  console.log(label, { records: state.records.length, activeView: state.activeView });
}

function render() {
  provider.rebuild();
  const current = view();
  spread.suspendPaint();
  try {
    sheet.setRowCount(Math.max(provider.rowCount,current.type==='grid'?40:1));
    sheet.setColumnCount(Math.max(provider.colCount, 1));
    dataRange.range(new S.Range(0, 0, provider.rowCount, provider.colCount));
    sheet.frozenRowCount(current.type === 'calendar' || current.type === 'gallery' ? 0 : 1);
    sheet.frozenColumnCount(current.type === 'grid'||current.type==='gantt' ? 2 : 0);
    if (current.type === 'grid') {
      sheet.setColumnWidth(0, 44);
      provider.fields.forEach((field, i) => sheet.setColumnWidth(i + 1, field.width));
      sheet.setColumnWidth(provider.colCount - 1, 45);
      for (let r = 0; r < sheet.getRowCount(); r++) sheet.setRowHeight(r, provider.rows[r]?.kind === 'spacer' ? 14 : provider.rows[r]?.kind === 'group' ? 46 : r === 0 ? 40 : 38);
    } else if(current.type==='board') {
      const width = Math.max(270, Math.min(340, Math.floor(($('ss').clientWidth - 18) / Math.max(1, provider.groups.length))));
      for (let c = 0; c < provider.colCount; c++) sheet.setColumnWidth(c, width);
      for (let r = 0; r < sheet.getRowCount(); r++) sheet.setRowHeight(r, r === 0 ? 60 : 163);
    } else if(current.type==='calendar') {
      for(let c=0;c<7;c++)sheet.setColumnWidth(c,Math.max(100,Math.floor(($('ss').clientWidth-20)/7)));
      sheet.setRowHeight(0,30);for(let r=1;r<7;r++)sheet.setRowHeight(r,Math.max(105,Math.floor(($('ss').clientHeight-50)/6)));
      sheet.frozenRowCount(1);
    } else if(current.type==='gantt') {
      sheet.setColumnWidth(0,220);sheet.setColumnWidth(1,120);
      for(let c=2;c<provider.colCount;c++)sheet.setColumnWidth(c,44);
      for(let r=0;r<sheet.getRowCount();r++)sheet.setRowHeight(r,r===0?38:48);
    } else {
      const width=Math.floor(($('ss').clientWidth-20)/provider.galleryColumns);
      for(let c=0;c<provider.colCount;c++)sheet.setColumnWidth(c,width);
      for(let r=0;r<sheet.getRowCount();r++)sheet.setRowHeight(r,245);
      sheet.frozenRowCount(0);
    }
    dataRange.repaint();
  } finally { spread.resumePaint(); }
  const viewButton=(v,role)=>{const button=document.createElement('button');button.dataset.view=v.id;button.className=role==='tab'?'view-tab':'side-view';button.setAttribute('role',role);button.setAttribute('aria-selected',String(v.id===state.activeView));button.classList.toggle('active',v.id===state.activeView);button.innerHTML=`<svg><use href="#i-${v.type==='board'?'board':v.type==='grid'?'table':v.type==='calendar'?'calendar':v.type==='gantt'?'gantt':'gallery'}"/></svg>`;button.append(document.createTextNode(v.name));button.onclick=()=>switchView(v.id);button.oncontextmenu=e=>{e.preventDefault();switchView(v.id);openManager('view');};return button;};
  $('view-tabs').replaceChildren(...state.views.map(v=>viewButton(v,'tab')));
  $('side-views').replaceChildren(...state.views.map(v=>viewButton(v,'button')));
  $('table-select').replaceChildren(...workspace.tables.map(t=>new Option(t.name,t.id)));
  $('table-select').value=state.id;
  document.querySelector('h1').textContent=state.name;
  document.querySelectorAll('[data-view]').forEach(button => {
    const active = button.dataset.view === state.activeView;
    button.classList.toggle('active', active);
    if (button.getAttribute('role') === 'tab') button.setAttribute('aria-selected', String(active));
  });
  const filterCount = current.filters?.conditions?.length || 0;
  $('filter-count').hidden = !filterCount;
  $('filter-count').textContent = filterCount;
  $('filter-button').classList.toggle('active', !!filterCount);
  $('sort-button').classList.toggle('active', sortRules(current).length>0);
  $('group-button').classList.toggle('active', groupFields(current).length>0);
  $('group-label').textContent = groupFields(current).length ? `分组 · ${groupFields(current).map(id=>state.fields.find(f=>f.id===id)?.name).filter(Boolean).join(' / ')}` : '分组';
  $('search-input').value = current.query;
  const count = projectRecords(state, current).length;
  $('record-count').textContent = count === state.records.length ? `${count} 条记录` : `${count} / ${state.records.length} 条记录`;
  const groups = leafGroups(state,current);
  $('group-count').textContent = groupFields(current).length ? `${groups.length} 个分组` : '未分组';
  $('sidebar-count').textContent = state.records.length;
  $('undo-button').disabled = !history.past.length;
  $('redo-button').disabled = !history.future.length;
  $('fields-button').disabled = false;
  $('group-button').disabled=!['grid','board'].includes(current.type);
  const groupName = groupFields(current).map(id=>state.fields.find(f=>f.id===id)?.name).filter(Boolean).join(' / ');
  $('view-description').textContent = current.type === 'grid' ? (groupName ? `按${groupName}分组` : '集中查看和编辑全部任务') : `${VIEW_TYPES[current.type]}视图${groupName?` · ${groupName}`:''}`;
  $('view-hint').textContent = current.type === 'grid' ? '双击编辑单元格 · 点击序号打开详情' : current.type === 'board'?'点击卡片查看详情 · 拖动卡片调整分组':current.type==='calendar'?'拖动改期 · 日期右上角＋新增':current.type==='gantt'?'拖动任务条改期 · 两端调整起止日期':'点击记录打开详情';
  const period=$('period-control');period.replaceChildren();
  if(current.type==='calendar'){
    const prev=document.createElement('button');prev.textContent='‹';prev.title='上个月';prev.onclick=()=>{const d=new Date(current.month+'-01T00:00:00');d.setMonth(d.getMonth()-1);current.month=dateKey(d).slice(0,7);render();save();};
    const label=document.createElement('strong');label.textContent=current.month;
    const next=document.createElement('button');next.textContent='›';next.title='下个月';next.onclick=()=>{const d=new Date(current.month+'-01T00:00:00');d.setMonth(d.getMonth()+1);current.month=dateKey(d).slice(0,7);render();save();};
    const date=document.createElement('select');setOptions(date,state.fields.filter(f=>f.type==='date').map(f=>({id:f.id,name:f.name})));date.value=current.dateField;date.setAttribute('aria-label','日历日期字段');date.onchange=()=>{current.dateField=date.value;render();save();};
    const today=document.createElement('button');today.textContent='今天';today.onclick=()=>{current.month=dateKey().slice(0,7);render();save();};
    const unscheduled=provider.records.filter(r=>!r[current.dateField]);const list=document.createElement('button');list.textContent=`未排期 ${unscheduled.length}`;list.onclick=()=>showRecordList('未排期记录',unscheduled);period.append(prev,label,next,today,date,list);
  }
  if(current.type==='gantt'){
    const prev=document.createElement('button');prev.textContent='‹';prev.title='向前一周';prev.onclick=()=>{current.timelineStart=addDays(current.timelineStart,-7);render();save();};
    const label=document.createElement('strong');label.textContent=`${current.timelineStart} 起`;
    const next=document.createElement('button');next.textContent='›';next.title='向后一周';next.onclick=()=>{current.timelineStart=addDays(current.timelineStart,7);render();save();};period.append(prev,label,next);
    const dateFields=state.fields.filter(f=>f.type==='date');
    for(const key of ['startField','endField']){const select=document.createElement('select');setOptions(select,dateFields.map(f=>({id:f.id,name:f.name})));select.value=current[key];select.setAttribute('aria-label',key==='startField'?'甘特开始字段':'甘特结束字段');select.onchange=()=>{current[key]=select.value;render();save();};period.append(select);}
  }
}

function switchView(id) {
  if (!sheet.endEdit()) return;
  if (pendingBefore) flushEdits();
  state.activeView = id;
  provider.rebuild();
  sheet.setActiveCell(0, 0);
  sheet.showCell(0, 0, S.VerticalPosition.top, S.HorizontalPosition.left);
  render(); save();
}

function openTool(id, anchor) {
  sheet.endEdit();
  const dialog = $(id);
  const rect = anchor.getBoundingClientRect();
  dialog.style.left = `${Math.max(12, Math.min(rect.left, window.innerWidth - 310))}px`;
  dialog.style.top = `${rect.bottom + 8}px`;
  dialog.showModal();
}

function setOptions(select, choices, blankLabel) {
  select.replaceChildren();
  if (blankLabel !== undefined) select.add(new Option(blankLabel, ''));
  choices.forEach(choice => select.add(new Option(typeof choice === 'string' ? choice : choice.name, typeof choice === 'string' ? choice : choice.id)));
}

function showFields() {
  const container = $('field-list'); container.replaceChildren();
  const ordered = [...view().columns.map(id => state.fields.find(f => f.id === id)).filter(Boolean), ...state.fields.filter(f => !view().columns.includes(f.id))];
  ordered.forEach((field, i) => {
    const row = document.createElement('div'); row.className = 'field-row';
    const label = document.createElement('label');
    const check = document.createElement('input'); check.type = 'checkbox'; check.checked = view().columns.includes(field.id); check.disabled = field.id === 'title';
    check.onchange = () => { commit('调整字段显示',()=>{view().columns = check.checked ? [...view().columns, field.id] : view().columns.filter(id => id !== field.id);}); showFields(); };
    const type = document.createElement('span'); type.className = 'field-type-icon'; type.textContent = fieldIcons[field.type];
    label.append(check, type, document.createTextNode(field.name)); row.append(label);
    const edit=document.createElement('button');edit.textContent='✎';edit.title='编辑字段';edit.setAttribute('aria-label',`编辑字段${field.name}`);edit.onclick=()=>openNewField(field);row.append(edit);
    const remove=document.createElement('button');remove.textContent='×';remove.title='删除字段';remove.setAttribute('aria-label',`删除字段${field.name}`);remove.disabled=field.id==='title';remove.onclick=()=>{if(confirm(`删除字段“${field.name}”？`))commit('删除字段',()=>deleteField(state,field.id));showFields();};row.append(remove);
    for (const [direction, symbol] of [[-1, '↑'], [1, '↓']]) {
      const button = document.createElement('button'); button.textContent = symbol; button.title = direction === -1 ? '上移' : '下移'; button.setAttribute('aria-label', `${button.title}${field.name}`);
      const visibleIndex = view().columns.indexOf(field.id), target = visibleIndex + direction;
      button.disabled = visibleIndex < 0 || target < 0 || target >= view().columns.length;
      button.onclick = () => { commit('调整字段顺序',()=>{const ids = view().columns; [ids[visibleIndex], ids[target]] = [ids[target], ids[visibleIndex]];}); showFields(); };
      row.append(button);
    }
    container.append(row);
  });
}

function openNewField(field=null) {
  $('fields-dialog').close();
  const form=$('new-field-form');form.reset();$('new-options-label').hidden=true;$('new-field-error').textContent='';editingFieldId=field?.id||null;
  form.elements.name.value=field?.name||'';form.elements.type.value=field?.type||'text';form.elements.options.value=(field?.options||[]).join(',');form.elements.defaultValue.value=field?.defaultValue??'';form.elements.width.value=field?.width||150;
  form.querySelector('h2').textContent=field?'编辑字段':'添加字段';form.querySelector('[type=submit]').textContent=field?'保存字段':'创建字段';
  $('new-options-label').hidden=!['select','multiselect','person'].includes(form.elements.type.value);
  form.elements.format.value=field?.format||'plain';form.elements.buttonLabel.value=field?.buttonLabel||'查看详情';form.elements.action.value=field?.action||'detail';fieldTypeUI();
  $('new-field-dialog').showModal();
}

function openRecord(id, defaults = {}) {
  sheet.endEdit();
  if (pendingBefore) flushEdits();
  editingId = id;
  const record = id ? findRecord(id) : Object.assign(defaultsFor(state,workspace.currentUser), { title: '' }, defaults);
  $('record-dialog-title').textContent = id ? '任务详情' : '添加任务';
  $('record-error').textContent = ''; $('delete-record').hidden = !id;
  const formFields = $('record-fields'); formFields.replaceChildren();
  for (const field of state.fields) {
    const label = document.createElement('label'); label.append(document.createTextNode(field.name));
    const input = document.createElement(['select','person'].includes(field.type) ? 'select' : field.type === 'multiselect' ? 'select' : field.id === 'note' ? 'textarea' : 'input');
    input.name = field.id; input.setAttribute('aria-label', field.name);
    if (['select','person'].includes(field.type)) setOptions(input, field.options||[], '未设置');
    else if(field.type==='multiselect'){input.multiple=true;setOptions(input,field.options||[]);}
    else if (field.id !== 'note') input.type = field.type === 'date' ? 'date' : ['number','progress','rating'].includes(field.type) ? 'number' : 'text';
    if (['number','progress','rating'].includes(field.type)) input.step = field.type==='rating'?'1':'any';
    const raw=valueOf(record,field); if(input.multiple)[...input.options].forEach(o=>o.selected=(raw||[]).includes(o.value));else input.value=raw;
    input.disabled = !editable(field); input.required = field.id === 'title';
    label.append(input); formFields.append(label);
  }
  $('record-dialog').showModal();
}

function showRecordList(title,records,defaults={}){
  $('records-list-title').textContent=title;$('records-list').replaceChildren(...records.map(record=>{const button=document.createElement('button');button.textContent=record.title;button.onclick=()=>{$('records-list-dialog').close();openRecord(record.id);};return button;}));
  $('list-add-record').onclick=()=>{$('records-list-dialog').close();openRecord(null,defaults);};$('records-list-dialog').showModal();
}

$('add-button').onclick = () => openRecord(null);
$('fields-button').onclick = () => { sheet.endEdit(); showFields(); $('fields-dialog').showModal(); };
$('new-field-button').onclick = () => openNewField();
$('search-input').oninput = event => { if (pendingBefore) flushEdits(); view().query = event.target.value; render(); save(); };
let filterDraft;
function filterEditor(node, container) {
  const group=document.createElement('section');group.className='filter-group';
  const head=document.createElement('div');head.className='filter-group-head';
  const logic=document.createElement('select');logic.setAttribute('aria-label','条件组关系');setOptions(logic,[{id:'and',name:'全部满足 AND'},{id:'or',name:'任一满足 OR'}]);logic.value=node.logic;logic.onchange=()=>node.logic=logic.value;
  const add=document.createElement('button');add.type='button';add.textContent='＋ 条件';add.onclick=()=>{node.conditions.push({field:state.fields[0].id,op:'eq',value:''});drawFilters();};
  const sub=document.createElement('button');sub.type='button';sub.textContent='＋ 条件组';sub.onclick=()=>{node.conditions.push({logic:'or',conditions:[]});drawFilters();};head.append(logic,add,sub);group.append(head);
  for(const condition of node.conditions){
    const wrapper=document.createElement('div');wrapper.className=condition.conditions?'filter-subgroup':'filter-condition';
    if(condition.conditions)filterEditor(condition,wrapper);
    else {
      const field=document.createElement('select');field.setAttribute('aria-label','筛选字段');setOptions(field,state.fields.filter(f=>f.type!=='button'));field.value=condition.field;
      const op=document.createElement('select');op.setAttribute('aria-label','筛选操作');setOptions(op,Object.entries({eq:'等于',ne:'不等于',contains:'包含',notContains:'不包含',gt:'大于 / 晚于',gte:'大于等于',lt:'小于 / 早于',lte:'小于等于',empty:'为空',notEmpty:'不为空'}).map(([id,name])=>({id,name})));op.value=condition.op;
      let value;
      const updateValue=()=>{
        const f=state.fields.find(f=>f.id===field.value),select=['select','person','multiselect'].includes(f?.type)&&['eq','ne'].includes(op.value);
        const input=document.createElement(select?'select':'input');input.setAttribute('aria-label','筛选值');
        if(select)setOptions(input,f.options||[],'未设置');else input.type=f?.type==='date'?'date':['number','progress','rating'].includes(f?.type)?'number':'text';
        input.value=condition.value??'';input.disabled=['empty','notEmpty'].includes(op.value);input.oninput=()=>condition.value=input.value;input.onchange=()=>condition.value=input.value;
        if(value)value.replaceWith(input);else wrapper.append(input);value=input;
      };
      field.onchange=()=>{condition.field=field.value;condition.value='';updateValue();};op.onchange=()=>{condition.op=op.value;updateValue();};wrapper.append(field,op);updateValue();
    }
    const remove=document.createElement('button');remove.type='button';remove.textContent='×';remove.setAttribute('aria-label','删除条件');remove.onclick=()=>{node.conditions.splice(node.conditions.indexOf(condition),1);drawFilters();};wrapper.append(remove);group.append(wrapper);
  }
  container.append(group);
}
function drawFilters(){const container=$('filter-conditions');container.replaceChildren();filterEditor(filterDraft,container);}
$('filter-button').onclick=()=>{filterDraft=structuredClone(view().filters);drawFilters();openTool('filter-dialog',$('filter-button'));};
$('add-condition').onclick=()=>{filterDraft.conditions.push({field:state.fields[0].id,op:'eq',value:''});drawFilters();};
$('filter-form').onsubmit=event=>{event.preventDefault();commit('配置筛选',()=>view().filters=structuredClone(filterDraft));$('filter-dialog').close();};
$('clear-filter').onclick=()=>{commit('清除筛选',()=>view().filters={logic:'and',conditions:[]});$('filter-dialog').close();};
$('group-button').onclick = () => {
  const choices=state.fields.filter(f=>['select','person','multiselect'].includes(f.type));
  setOptions($('group-field'),choices,'不分组');setOptions($('group-field-2'),choices,'无');
  $('group-field').value=groupFields(view())[0]||'';$('group-field-2').value=groupFields(view())[1]||'';
  openTool('group-dialog', $('group-button'));
};
$('group-form').onsubmit = event => { event.preventDefault();commit('配置分组',()=>{view().groupFields=[...new Set([$('group-field').value,$('group-field-2').value].filter(Boolean))];view().collapsed=[];});$('group-dialog').close(); };
$('sort-button').onclick = () => {
  setOptions($('sort-field'),state.fields,'默认顺序');setOptions($('sort-field-2'),state.fields,'无');const rules=sortRules(view());$('sort-field').value=rules[0]?.field||'';$('sort-field-2').value=rules[1]?.field||'';$('sort-form').elements.direction.value=rules[0]?.direction||'asc';$('sort-form').elements.direction2.value=rules[1]?.direction||'asc';
  openTool('sort-dialog', $('sort-button'));
};
$('sort-form').onsubmit = event => { event.preventDefault();const f=event.currentTarget.elements;commit('配置排序',()=>{view().sorts=[{field:f.field.value,direction:f.direction.value},{field:f.field2.value,direction:f.direction2.value}].filter((x,i,items)=>x.field&&items.findIndex(item=>item.field===x.field)===i);});$('sort-dialog').close(); };
function fieldTypeUI() {
  const form=$('new-field-form'),type=form.elements.type.value;
  $('new-options-label').hidden=!['select','multiselect','person'].includes(type);
  $('button-options').hidden=type!=='button';
  form.elements.format.disabled=type!=='number';
  form.elements.defaultValue.disabled=!editable({type});
  form.elements.defaultValue.type=type==='date'?'date':['number','progress','rating'].includes(type)?'number':'text';
  form.elements.defaultValue.step=type==='rating'?'1':'any';
  form.elements.defaultValue.required=false;
}
$('new-field-form').elements.type.onchange=()=>{const form=$('new-field-form');form.elements.defaultValue.value='';if(form.elements.type.value==='person'&&!form.elements.options.value)form.elements.options.value=workspace.members.join(',');fieldTypeUI();};
$('new-field-form').elements.type.replaceChildren(...Object.entries(TYPES).map(([value,label])=>new Option(label,value)),new Option('固定延期计算','computed'));
$('new-field-form').onsubmit=event=>{
  event.preventDefault();const form=event.currentTarget,previous=state.fields.find(f=>f.id===editingFieldId);
  const field={...previous,id:editingFieldId||uid('field'),name:form.elements.name.value.trim(),type:form.elements.type.value,width:Math.max(60,Math.min(600,Number(form.elements.width.value)||150)),format:form.elements.format.value,defaultValue:form.elements.defaultValue.value};
  if(['select','multiselect','person'].includes(field.type))field.options=[...new Set(form.elements.options.value.split(/[,，\n]/).map(s=>s.trim()).filter(Boolean))];
  if(field.type==='button'){field.buttonLabel=form.elements.buttonLabel.value.trim()||'查看详情';field.action=form.elements.action.value;}
  try{
    if(field.type==='computed'&&editingFieldId!=='overdue')throw new Error('固定延期计算仅用于内置延期字段');
    commit(editingFieldId?'编辑字段':'添加字段',()=>saveField(state,field,editingFieldId));
    editingFieldId=null;$('new-field-dialog').close();
  }catch(error){$('new-field-error').textContent=error.message;}
};
$('record-form').onsubmit = event => {
  event.preventDefault(); const form = event.currentTarget;
  const patch = Object.fromEntries(state.fields.filter(editable).map(f => {const input=form.elements.namedItem(f.id);return[f.id,input.multiple?[...input.selectedOptions].map(o=>o.value):input.value];}));
  try {
    // Let changing status set/clear completion date unless the user supplied a new one.
    if (editingId && patch.completed === findRecord(editingId).completed && patch.status !== findRecord(editingId).status) delete patch.completed;
    if (!editingId && patch.status === '已完成' && !patch.completed && state.fields.some(f=>f.id==='completed'&&f.type==='date')) patch.completed = dateKey();
    const draft = { ...state, records: structuredClone(state.records) };
    const id = editingId || crypto.randomUUID();
    if (!editingId) draft.records.push({ id,__creator:workspace.currentUser });
    updateRecord(draft, id, patch);
    commit(editingId ? '编辑记录' : '添加记录', () => { state.records = draft.records; });
    $('record-dialog').close();
  } catch (error) { $('record-error').textContent = error.message; }
};
$('delete-record').onclick = () => {
  if (!editingId || !confirm('删除这条任务？可以通过撤销恢复。')) return;
  commit('删除记录', () => { state.records = state.records.filter(r => r.id !== editingId); }); $('record-dialog').close();
};
function undo() { sheet.endEdit(); if (pendingBefore) flushEdits(); history.undo();state=workspace.tables.find(t=>t.id===workspace.activeTableId);render(); save(); }
function redo() { sheet.endEdit(); if (pendingBefore) flushEdits(); history.redo();state=workspace.tables.find(t=>t.id===workspace.activeTableId);render(); save(); }
$('undo-button').onclick = undo; $('redo-button').onclick = redo;
const collapseButton = $('collapse-header');
const collapsibleBlocks = document.querySelectorAll('.site-header,.page-header,.view-tabs,.view-caption');
collapseButton.onclick = () => {
  const collapsed = document.body.classList.toggle('header-collapsed');
  // Animating height keeps the blocks in the DOM; make them unreachable for
  // keyboard and screen readers while they are collapsed.
  collapsibleBlocks.forEach(block => { block.inert = collapsed; });
  collapseButton.setAttribute('aria-expanded', String(!collapsed));
  collapseButton.setAttribute('aria-label', collapsed ? '展开顶部区域' : '收起顶部区域');
  collapseButton.title = collapsed ? '展开顶部区域' : '收起顶部区域';
};
const toast = $('toast');
let toastTimer = 0;
function showToast(message) {
  $('toast-text').textContent = message;
  toast.classList.add('is-visible');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove('is-visible'), 2600);
}
$('reset-button').onclick = () => $('reset-dialog').showModal();
$('confirm-reset').onclick = () => {
  $('reset-dialog').close();
  workspace=createWorkspace();state=workspace.tables[0];history=new History(workspace);pendingBefore=null;cancelAnimationFrame(editFrame);editFrame=0;render();save();
  showToast(`已恢复示例数据，共 ${state.records.length} 条任务`);
};
// 线上发布时把这里换成视频链接（CDN / 对象存储）；本地开发用项目根目录的 intro.mp4，
// 该文件只由开发服务器提供，不会打进 dist/。
const INTRO_VIDEO_URL = '../intro.mp4';
const introVideo = $('intro-video');
$('intro-button').onclick = () => {
  $('intro-fallback').hidden = true;
  if (!introVideo.getAttribute('src')) introVideo.src = INTRO_VIDEO_URL;
  $('intro-dialog').showModal();
};
$('intro-dialog').addEventListener('close', () => { introVideo.pause(); if (introVideo.readyState > 0) introVideo.currentTime = 0; });
introVideo.addEventListener('error', () => { $('intro-fallback').hidden = false; });
$('table-select').onchange=()=>{if(!sheet.endEdit()){$('table-select').value=state.id;return;}if(pendingBefore)flushEdits();workspace.activeTableId=$('table-select').value;state=workspace.tables.find(t=>t.id===workspace.activeTableId);render();save();};
$('table-select').oncontextmenu=event=>{event.preventDefault();openManager('table');};$('manage-table').onclick=()=>openManager('table');$('manage-view').onclick=()=>openManager('view');
let manageKind='view';
function openManager(kind){manageKind=kind;$('manage-title').textContent=kind==='view'?`管理视图：${view().name}`:`管理数据表：${state.name}`;$('delete-item').hidden=kind==='view'?state.views.length<2:workspace.tables.length<2;$('manage-dialog').showModal();}
$('rename-item').onclick=()=>{const item=manageKind==='view'?view():state;const name=prompt('输入新名称',item.name);if(name?.trim())commit('重命名',()=>item.name=name.trim());$('manage-dialog').close();};
$('duplicate-item').onclick=()=>{if(manageKind==='view')commit('复制视图',()=>{const copy=structuredClone(view());copy.id=uid('view');copy.name=`${view().name} 副本`;state.views.splice(state.views.indexOf(view())+1,0,copy);state.activeView=copy.id;});else commit('复制数据表',()=>{const copy=structuredClone(state);copy.id=uid('table');copy.name=state.name+' 副本';copy.records=copy.records.map(r=>({...r,id:uid('record')}));workspace.tables.push(copy);workspace.activeTableId=copy.id;state=copy;});$('manage-dialog').close();};
$('delete-item').onclick=()=>{if(manageKind==='view'){const id=view().id;if(confirm(`删除视图“${view().name}”？`))commit('删除视图',()=>{state.views=state.views.filter(v=>v.id!==id);if(state.activeView===id)state.activeView=state.views[0].id;});}else{const id=state.id;if(confirm(`删除数据表“${state.name}”？`))commit('删除数据表',()=>{workspace.tables=workspace.tables.filter(t=>t.id!==id);workspace.activeTableId=workspace.tables[0].id;state=workspace.tables[0];});}$('manage-dialog').close();};
$('add-table').onclick=()=>{const name=prompt('新数据表名称','新数据表');if(!name?.trim())return;commit('新建数据表',()=>{const table=newTable(name.trim());workspace.tables.push(table);workspace.activeTableId=table.id;state=table;});};
$('add-view').onclick=()=>{$('view-form').reset();setOptions($('view-form').elements.type,Object.entries(VIEW_TYPES).map(([id,name])=>({id,name})));$('view-dialog').showModal();};
$('view-form').onsubmit=event=>{event.preventDefault();const form=event.currentTarget;const type=form.elements.type.value,name=form.elements.name.value.trim();if(!name)return;commit('新建视图',()=>{const item=makeView(state,type,name);state.views.push(item);state.activeView=item.id;});$('view-dialog').close();};
document.querySelectorAll('[data-close]').forEach(button => button.onclick = () => button.closest('dialog').close());
document.querySelectorAll('.tool-dialog, .center-dialog').forEach(dialog => dialog.addEventListener('click', event => {
  const rect = dialog.getBoundingClientRect();
  if (event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom) dialog.close();
}));
document.addEventListener('keydown', event => {
  const inInput = event.target.closest('input,textarea,select,[contenteditable=true]');
  if ((event.ctrlKey || event.metaKey) && !inInput && !document.querySelector('dialog[open]')) {
    if (event.key.toLowerCase() === 'z' || event.key.toLowerCase() === 'y') {
      event.preventDefault(); event.stopPropagation();
      if (event.shiftKey || event.key.toLowerCase() === 'y') redo(); else undo();
    }
  }
}, true);

function hitColumn(event) {
  const rect = $('ss').getBoundingClientRect();
  if (event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom) return -1;
  const hit = sheet.hitTest(event.clientX - rect.left, event.clientY - rect.top);
  return hit && hit.col >= 0 && hit.col < provider.groups.length ? hit.col : -1;
}
function hitCell(event){const rect=$('ss').getBoundingClientRect();if(event.clientX<rect.left||event.clientX>rect.right||event.clientY<rect.top||event.clientY>rect.bottom)return null;const hit=sheet.hitTest(event.clientX-rect.left,event.clientY-rect.top);return hit&&hit.row>=0&&hit.row<provider.rowCount&&hit.col>=0&&hit.col<provider.colCount?hit:null;}
document.addEventListener('mousemove', event => {
  if (!drag) return;
  if (!drag.moved && Math.hypot(event.clientX - drag.x, event.clientY - drag.y) > 6) {
    drag.moved = true; ghost = document.createElement('div'); ghost.className = 'drag-ghost'; ghost.textContent = findRecord(drag.id)?.title; document.body.append(ghost);
  }
  if (!drag.moved) return;
  ghost.style.left = `${event.clientX + 15}px`; ghost.style.top = `${event.clientY + 12}px`;
  if(drag.kind==='calendar'||drag.kind==='gantt')return;
  const col = hitColumn(event);
  if (col !== dropColumn) { dropColumn = col; provider.styleCache.clear(); dataRange.repaint(); }
});
document.addEventListener('mouseup', event => {
  if (!drag) return;
  const finished = drag; drag = null; ghost?.remove(); ghost = null; dropColumn = -1;
  if (!finished.moved) { openRecord(finished.id); return; }
  if(finished.kind==='calendar'){
    const hit=hitCell(event);if(hit&&hit.row>0){const date=provider.calendarDays[(hit.row-1)*7+hit.col];try{commit('日历拖动改期',()=>updateRecord(state,finished.id,{[finished.field]:date}));}catch(error){alert(error.message);dataRange.repaint();}}else dataRange.repaint();return;
  }
  if(finished.kind==='gantt'){
    const hit=hitCell(event);if(!hit||hit.col<2){dataRange.repaint();return;}
    const delta=hit.col-finished.sourceColumn,start=finished.mode==='end'?finished.start:addDays(finished.start,delta),end=finished.mode==='start'?finished.end:addDays(finished.end,delta);
    if(end<start){alert('结束日期不能早于开始日期');dataRange.repaint();return;}
    commit('甘特拖动改期',()=>updateRecord(state,finished.id,{[finished.startField]:start,[finished.endField]:end}));return;
  }
  const col = hitColumn(event);
  if (col >= 0 && col !== finished.sourceColumn) {
    const values=provider.groups[col].values;
    commit('看板拖动更新字段', () => updateRecord(state, finished.id, values));
  } else { provider.styleCache.clear(); dataRange.repaint(); }
});
window.addEventListener('blur', () => { drag = null; ghost?.remove(); ghost = null; dropColumn = -1; provider.styleCache.clear(); dataRange.repaint(); });
let resizeFrame=0,resizeTimer=0,lastSize='';
new ResizeObserver(entries=>{
 const rect=entries[0].contentRect,key=Math.round(rect.width)+'x'+Math.round(rect.height);if(key===lastSize)return;lastSize=key;
 // The header collapse transition resizes the shell every frame; keep the canvas
 // in sync immediately and defer the heavier layout rebuild until the size settles.
 cancelAnimationFrame(resizeFrame);resizeFrame=requestAnimationFrame(()=>spread.refresh());
 clearTimeout(resizeTimer);
 resizeTimer=setTimeout(()=>{if(!sheet.endEdit())return;if(pendingBefore)flushEdits();render();},180);
}).observe($('spread-shell'));
sheet.bind(S.Events.EditEnded, () => { if (pendingBefore && !editFrame) editFrame = requestAnimationFrame(flushEdits); });

render(); save();
// Console inspection entry: no extra diagnostic panels in the business UI.
window.demo = { spread, sheet, dataRange, provider, get state() { return state; }, get workspace(){return workspace;}, switchView, openRecord, undo, redo };
console.log('DataRange 多维表格已就绪', { version: S.productVersion, views: state.views.map(v => v.name), records: state.records.length });
