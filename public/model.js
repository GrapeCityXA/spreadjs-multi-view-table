// Business records and view rules deliberately have no worksheet row/column coordinates.
export const LEGACY_KEY = 'spreadjs-datarange-tasks-v1';
export const STORAGE_KEY = 'spreadjs-datarange-workspace-v2';
export const STATUSES = ['待开始', '进行中', '已完成', '已暂停'];
export const OWNERS = ['林晓', '周北北', '黄泡泡', '刘贝拉'];
export const PRIORITIES = ['重要紧急', '紧急不重要', '重要不紧急', '不紧急不重要'];

export function dateKey(date = new Date()) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

export function relativeDate(days, base = new Date()) {
  const date = new Date(base.getFullYear(), base.getMonth(), base.getDate());
  date.setDate(date.getDate() + days);
  return dateKey(date);
}

export function createDocument() {
  const fields = [
    { id: 'title', name: '任务描述', type: 'text', width: 290 },
    { id: 'owner', name: '任务执行人', type: 'select', options: OWNERS, width: 130 },
    { id: 'status', name: '进度', type: 'select', options: STATUSES, defaultValue:'待开始', width: 112 },
    { id: 'start', name: '开始日期', type: 'date', width: 132 },
    { id: 'due', name: '预计完成日期', type: 'date', width: 140 },
    { id: 'overdue', name: '是否延期', type: 'computed', width: 120 },
    { id: 'completed', name: '实际完成日期', type: 'date', width: 140 },
    { id: 'note', name: '最新进展记录', type: 'text', width: 260 },
    { id: 'priority', name: '重要紧急程度', type: 'select', options: PRIORITIES, width: 148 }
  ];
  const rows = [
    ['完成年度财务报告', '林晓', '进行中', -12, -2, '', '已收集各部门数据，正在核对报表。', 0],
    ['组织年度员工团建活动', '周北北', '已完成', -18, -3, -4, '活动顺利结束，满意度反馈已归档。', 0],
    ['更新公司网站', '黄泡泡', '进行中', -7, 5, '', '首页设计已确认，正在联调。', 0],
    ['招聘新员工', '周北北', '进行中', -9, -1, '', '已完成首轮面试，等待业务负责人反馈。', 0],
    ['多维表格应用模式员工培训', '刘贝拉', '待开始', 1, 8, '', '培训材料准备中。', 0],
    ['开发新产品原型', '黄泡泡', '待开始', -3, -1, '', '需求已梳理，等待产品评审。', 1],
    ['进行客户满意度调查', '刘贝拉', '待开始', -2, 6, '', '问卷已设计，计划本周发送。', 1],
    ['优化供应链管理', '黄泡泡', '待开始', 2, 15, '', '已识别三个关键优化环节。', 2],
    ['改善工作环境', '周北北', '进行中', -4, 10, '', '已收集员工意见和改造建议。', 2],
    ['制定新的销售策略', '刘贝拉', '进行中', -3, 12, '', '已完成市场分析，正在制定方案。', 2],
    ['提升员工培训体验', '黄泡泡', '已暂停', -5, 18, '', '等待下一阶段预算安排。', 3],
    ['整理共享资料库', '林晓', '已完成', -8, -1, -2, '历史资料已完成分类和归档。', 3]
  ];
  const records = rows.map((r, i) => ({ id: `task-${i + 1}`, title: r[0], owner: r[1], status: r[2], start: relativeDate(r[3]), due: relativeDate(r[4]), completed: r[5] === '' ? '' : relativeDate(r[5]), note: r[6], priority: PRIORITIES[r[7]] }));
  const makeView = (id, name, type, groupBy) => ({ id, name, type, groupBy, columns: fields.map(f => f.id), query: '', filter: { status: '', owner: '', overdue: false }, sort: { field: '', direction: 'asc' }, collapsed: [] });
  return { version: 1, fields, records, activeView: 'table', views: [makeView('table', '任务管理表', 'grid', 'priority'), makeView('status', '进度看板', 'board', 'status'), makeView('owner', '人员任务分配看板', 'board', 'owner')] };
}


export const TYPES = {text:'文本',number:'数字',date:'日期',select:'单选',multiselect:'多选',person:'人员（本地成员）',progress:'进度',rating:'评分',button:'按钮',creator:'创建人'};
export const VIEW_TYPES = {grid:'表格',board:'看板',calendar:'日历',gantt:'甘特',gallery:'画册'};
export const READ_ONLY = ['computed','creator','button'];
export const uid = prefix => prefix+'-'+crypto.randomUUID();
export const groupFields = view => view.groupFields ?? (view.groupBy ? [view.groupBy] : []);
export const sortRules = view => view.sorts ?? (view.sort?.field ? [view.sort] : []);
export const addDays = (key, days) => relativeDate(days, new Date(key+'T00:00:00'));
export const dayDiff = (a,b) => Math.round((Date.parse(b+'T00:00:00Z')-Date.parse(a+'T00:00:00Z'))/86400000);
export const editable = field => !READ_ONLY.includes(field.type);
export function makeView(table,type,name) {
  return {id:uid('view'),type,name:name||VIEW_TYPES[type],columns:table.fields.map(f=>f.id),query:'',filters:{logic:'and',conditions:[]},groupFields:type==='board'?[table.fields.find(f=>['select','person'].includes(f.type))?.id||'title']:[],sorts:[],collapsed:[],month:dateKey().slice(0,7),dateField:table.fields.find(f=>f.type==='date')?.id||'',startField:table.fields.find(f=>f.id==='start'&&f.type==='date')?.id||'',endField:table.fields.find(f=>f.id==='due'&&f.type==='date')?.id||'',timelineStart:addDays(dateKey(),-7),days:42};
}
export function upgradeTable(table, addViews=false) {
  table.id ||= uid('table'); table.name ||= '任务管理';
  table.views ||= [];
  if(addViews){
    if(!table.views.some(v=>v.type==='calendar'))table.views.push(makeView(table,'calendar','任务日历'));
    if(!table.views.some(v=>v.type==='gantt'))table.views.push(makeView(table,'gantt','项目甘特'));
    if(!table.views.some(v=>v.type==='gallery'))table.views.push(makeView(table,'gallery','任务画册'));
  }
  table.fields.forEach(f=>{ if(f.id==='title') f.required=true; });
  table.views.forEach(v=>{
    const defaults=makeView(table,v.type,v.name);
    const conditions=[];
    if(v.filter?.status) conditions.push({field:'status',op:'eq',value:v.filter.status});
    if(v.filter?.owner) conditions.push({field:'owner',op:'eq',value:v.filter.owner});
    if(v.filter?.overdue) conditions.push({field:'overdue',op:'eq',value:'已延期'});
    Object.assign(v,{...defaults,...v,filters:v.filters||{logic:'and',conditions},groupFields:groupFields(v),sorts:sortRules(v)});
    delete v.filter; delete v.sort; delete v.groupBy;
  });
  return table;
}
export function createWorkspace() {
  const table=upgradeTable(createDocument(),true);
  table.fields.find(f=>f.id==='owner').type='person';
  table.fields.push({id:'progress',name:'完成进度',type:'progress',width:150,defaultValue:0},{id:'tags',name:'任务标签',type:'multiselect',options:['项目','运营','内部'],width:165,defaultValue:[]},{id:'rating',name:'优先评分',type:'rating',width:140,defaultValue:0},{id:'creator',name:'创建人',type:'creator',width:120},{id:'action',name:'快捷操作',type:'button',action:'detail',buttonLabel:'查看详情',width:120});
  table.records.forEach((r,i)=>Object.assign(r,{progress:r.status==='已完成'?100:r.status==='进行中'?45:0,tags:[i%2?'运营':'项目'],rating:i%5+1,__creator:'林晓'}));
  table.views.forEach(v=>v.columns=table.fields.map(f=>f.id));
  table.views.find(v=>v.type==='calendar').dateField='due';
  return {version:2,members:OWNERS,currentUser:'林晓',activeTableId:table.id,tables:[table]};
}
function readWorkspace(storage) {
  const text=storage.getItem(STORAGE_KEY);
  if(text) {
    const value=JSON.parse(text);
    if(value.version===1&&Array.isArray(value.records)&&Array.isArray(value.fields)&&value.views?.length){
      upgradeTable(value,true);return {version:2,members:OWNERS,currentUser:'林晓',activeTableId:value.id,tables:[value]};
    }
    if(value.version!==2||!value.tables?.length) throw new Error('本地工作区格式不正确，原数据已保留');
    value.tables.forEach(t=>upgradeTable(t));
    if(!value.tables.some(t=>t.id===value.activeTableId)) value.activeTableId=value.tables[0].id;
    return value;
  }
  const old=storage.getItem(LEGACY_KEY);
  if(old) {
    const table=JSON.parse(old);
    if(!Array.isArray(table.fields)||!Array.isArray(table.records)||!table.views?.length) throw new Error('旧版本数据格式不正确，原数据已保留');
    upgradeTable(table,true);
    return {version:2,members:OWNERS,currentUser:'林晓',activeTableId:table.id,tables:[table]};
  }
  return createWorkspace();
}
// Rename the sample member in saved records, options and view rules without
// resetting the user's workspace. The legacy storage key remains a backup.
export function loadWorkspace(storage) {
  const workspace=readWorkspace(storage);
  if(workspace.sampleMemberRename===1)return workspace;
  function rename(value) {
    if(typeof value==='string')return value.replaceAll('段涵','林晓').replaceAll('段函','林晓');
    if(Array.isArray(value))return value.map(rename);
    if(value&&typeof value==='object')return Object.fromEntries(Object.entries(value).map(([key,item])=>[key,key==='id'?item:rename(item)]));
    return value;
  }
  const renamed=rename(workspace);renamed.sampleMemberRename=1;return renamed;
}
// Backward-compatible table loader for integrations that only need one table.
export function loadDocument(storage) { const workspace=loadWorkspace(storage); return workspace.tables.find(t=>t.id===workspace.activeTableId); }
export function overdueState(record,today=dateKey()) {
  if(record.status==='已暂停') return '已暂停';
  if(record.status==='已完成') return record.due&&record.completed&&record.completed>record.due?'延期完成':'已完成';
  return record.due&&record.due<today?'已延期':'正常';
}
export function fieldValue(record,fieldId,today) { return fieldId==='overdue'?overdueState(record,today):record[fieldId]??''; }
export function valueOf(record,field) { return field.type==='creator'?record.__creator||'未知（历史数据）':field.type==='button'?field.buttonLabel||'查看详情':fieldValue(record,field.id); }
export function normalizeValue(field,value) {
  if(!editable(field)) throw new Error('此字段由系统提供，不能直接修改');
  if(value instanceof Date && field.type==='date') value=dateKey(value);
  const text=value==null?'':String(value).trim();
  if((field.required||field.id==='title')&&!text) throw new Error(field.name+'不能为空');
  if(['number','progress','rating'].includes(field.type)) {
    if(!text) return field.type==='number'?'':0;
    const number=Number(text);
    if(!Number.isFinite(number)) throw new Error(field.name+'需要有效数字');
    if(field.type==='progress'&&(number<0||number>100)) throw new Error('进度应为 0–100');
    if(field.type==='rating'&&(!Number.isInteger(number)||number<0||number>5)) throw new Error('评分应为 0–5 的整数');
    return number;
  }
  if(field.type==='date'&&text) {
    const normalized=text.replaceAll('/','-'); const d=new Date(normalized+'T00:00:00');
    if(!/^\d{4}-\d{2}-\d{2}$/.test(normalized)||Number.isNaN(d.getTime())||dateKey(d)!==normalized) throw new Error(field.name+'请使用有效 YYYY-MM-DD 日期');
    return normalized;
  }
  if(field.type==='multiselect') {
    const values=[...new Set((Array.isArray(value)?value:text.split(/[,，]/)).map(s=>String(s).trim()).filter(Boolean))];
    if(values.some(v=>!field.options?.includes(v))) throw new Error(field.name+'包含无效选项');
    return values;
  }
  if(['select','person'].includes(field.type)&&text&&!(field.options||OWNERS).includes(text)) throw new Error(field.name+'请选择已有选项');
  return text;
}
export function updateRecord(table,id,patch) {
  const record=table.records.find(r=>r.id===id); if(!record) throw new Error('记录已不存在');
  const normalized={};
  for(const [key,value] of Object.entries(patch)) {
    const field=table.fields.find(f=>f.id===key); if(!field) throw new Error('未知字段：'+key);
    normalized[key]=normalizeValue(field,value);
  }
  if('status' in normalized&&normalized.status!==record.status&&!('completed' in normalized)&&table.fields.some(f=>f.id==='completed'&&f.type==='date')) normalized.completed=normalized.status==='已完成'?(record.completed||dateKey()):'';
  const next={...record,...normalized};
  if(('start' in normalized||'due' in normalized)&&table.fields.some(f=>f.id==='start'&&f.type==='date')&&table.fields.some(f=>f.id==='due'&&f.type==='date')&&next.start&&next.due&&next.due<next.start)throw new Error('截止日期不能早于开始日期');
  Object.assign(record,normalized); return record;
}
export function defaultsFor(table,user='林晓') {
  return Object.assign(Object.fromEntries(table.fields.filter(editable).map(f=>[f.id,structuredClone(f.defaultValue??(f.type==='multiselect'?[]:['progress','rating'].includes(f.type)?0:''))])),{__creator:user});
}
export const OPERATORS={eq:'等于',ne:'不等于',contains:'包含',notContains:'不包含',gt:'大于 / 晚于',gte:'大于等于',lt:'小于 / 早于',lte:'小于等于',empty:'为空',notEmpty:'不为空'};
export function matches(table,record,node,today) {
  if(node.conditions) { const results=node.conditions.map(n=>matches(table,record,n,today)); return !results.length||(node.logic==='or'?results.some(Boolean):results.every(Boolean)); }
  const field=table.fields.find(f=>f.id===node.field); if(!field) return true;
  const raw=field.type==='computed'?fieldValue(record,field.id,today):valueOf(record,field);
  const empty=raw===''||raw==null||Array.isArray(raw)&&!raw.length;
  if(node.op==='empty') return empty; if(node.op==='notEmpty') return !empty;
  const numeric=['number','progress','rating'].includes(field.type);
  const a=numeric?Number(raw):Array.isArray(raw)?raw.join(','):String(raw).toLocaleLowerCase();
  const b=numeric?Number(node.value):String(node.value??'').toLocaleLowerCase();
  if(numeric&&(empty||node.value===''||!Number.isFinite(b))) return false;
  switch(node.op) {
    case 'eq': return Array.isArray(raw)?raw.includes(node.value):a===b;
    case 'ne': return Array.isArray(raw)?!raw.includes(node.value):a!==b;
    case 'contains': return String(a).includes(String(b));
    case 'notContains': return !String(a).includes(String(b));
    case 'gt': return a>b; case 'gte': return a>=b; case 'lt': return a<b; case 'lte': return a<=b;
    default:return true;
  }
}
export function projectRecords(table,view,today=dateKey()) {
  const query=(view.query||'').trim().toLocaleLowerCase();
  const records=table.records.filter(r=>(!view.filters||matches(table,r,view.filters,today))&&(!view.filter||(!view.filter.status||r.status===view.filter.status)&&(!view.filter.owner||r.owner===view.filter.owner)&&(!view.filter.overdue||overdueState(r,today)==='已延期'))&&(!query||table.fields.some(f=>String(valueOf(r,f)).toLocaleLowerCase().includes(query))));
  const rules=sortRules(view);
  return records.sort((a,b)=>{
    for(const rule of rules) {
      const f=table.fields.find(f=>f.id===rule.field); if(!f) continue;
      const x=valueOf(a,f),y=valueOf(b,f);
      const order=['number','progress','rating'].includes(f.type)?(x===''?-Infinity:Number(x))-(y===''?-Infinity:Number(y)):String(x).localeCompare(String(y),'zh-CN');
      if(order) return order*(rule.direction==='desc'?-1:1);
    }
    return 0;
  });
}
export function groupRecords(table,records,fieldId,includeEmpty=false) {
  if(!fieldId) return [{key:'',records}];
  const field=table.fields.find(f=>f.id===fieldId);
  const keyFor=r=>String(field?valueOf(r,field):'');
  const keys=includeEmpty?[...(field?.options||[])]:[];
  records.forEach(r=>{const key=keyFor(r);if(!keys.includes(key)) keys.push(key);});
  if(field?.options) keys.sort((a,b)=>(field.options.indexOf(a)<0?999:field.options.indexOf(a))-(field.options.indexOf(b)<0?999:field.options.indexOf(b)));
  return keys.map(key=>({key,records:records.filter(r=>keyFor(r)===key)}));
}
export function leafGroups(table,view) {
  const fields=groupFields(view); const result=[];
  function walk(records,depth,values,labels) {
    if(depth===fields.length) {result.push({key:JSON.stringify(values),values,labels,records});return;}
    groupRecords(table,records,fields[depth],depth===0&&view.type==='board').forEach(g=>walk(g.records,depth+1,{...values,[fields[depth]]:g.key},[...labels,g.key||'未设置']));
  }
  walk(projectRecords(table,view),0,{},[]);return result;
}
export function gridRows(table,view) {
  const fields=groupFields(view),rows=[{kind:'header'}];
  function walk(records,depth,values) {
    if(depth===fields.length) {
      records.forEach((r,i)=>rows.push({kind:'record',id:r.id,index:i+1}));rows.push({kind:'add',values});return;
    }
    groupRecords(table,records,fields[depth]).forEach(g=>{
      const next={...values,[fields[depth]]:g.key},key=JSON.stringify(next);
      rows.push({kind:'group',key,label:Object.values(next).map(v=>v||'未设置').join(' / '),depth,count:g.records.length});
      if(!view.collapsed.includes(key)&&!view.collapsed.includes(g.key)) walk(g.records,depth+1,next);
    });
  }
  const records=projectRecords(table,view); walk(records,0,{});
  if(!records.length) rows.push({kind:'empty'},{kind:'add',values:{}});
  return rows;
}
function cleanFilters(node,ids) { if(node.conditions) return {...node,conditions:node.conditions.filter(n=>n.conditions||ids.has(n.field)).map(n=>cleanFilters(n,ids))}; return node; }
export function repairViews(table) {
  const ids=new Set(table.fields.map(f=>f.id)),dates=table.fields.filter(f=>f.type==='date').map(f=>f.id);
  table.views.forEach(v=>{
    v.columns=v.columns.filter(id=>ids.has(id));if(!v.columns.includes('title')) v.columns.unshift('title');
    v.groupFields=groupFields(v).filter(id=>ids.has(id)&&['select','person','multiselect'].includes(table.fields.find(f=>f.id===id).type));
    v.sorts=sortRules(v).filter(s=>ids.has(s.field));v.filters=cleanFilters(v.filters,ids);v.collapsed=[];
    for(const key of ['dateField','startField','endField']) if(!dates.includes(v[key])) v[key]=dates[0]||'';
  });
}
export function saveField(table,field,editingId) {
  if(!field.name?.trim()||table.fields.some(f=>f.name===field.name&&f.id!==editingId)) throw new Error('请填写不重复的字段名称');
  if(editingId==='title'&&field.type!=='text') throw new Error('主字段保留为文本');
  if(editingId==='overdue'&&field.type!=='computed') throw new Error('延期字段使用固定计算规则');
  const draft=structuredClone(table),previous=draft.fields.find(f=>f.id===editingId);
  if(editingId&&!previous) throw new Error('字段已不存在');
  field={...field,id:editingId||field.id||uid('field')};
  if(['select','multiselect'].includes(field.type)&&!field.options?.length) throw new Error('请至少填写一个选项');
  if(field.type==='person') field.options=field.options?.length?field.options:[...OWNERS];
  if(field.type==='creator'&&previous&&previous.type!=='creator') throw new Error('创建人请新增字段，历史记录不能推断真实身份');
  if(editable(field)) field.defaultValue=normalizeValue({...field,required:false,id:field.id==='title'?'default-title':field.id},field.defaultValue??'');
  const converted=draft.records.map(r=>{
    if(!editable(field)) return r;
    let value=previous?valueOf(r,previous):structuredClone(field.defaultValue);
    if(previous?.type==='multiselect'&&field.type!=='multiselect') value=value.join(',');
    return {...r,[field.id]:normalizeValue(field,value)};
  });
  if(previous) draft.fields[draft.fields.indexOf(previous)]=field;
  else {draft.fields.push(field);draft.views.forEach(v=>v.columns.push(field.id));}
  draft.records=converted;repairViews(draft);Object.assign(table,draft);return field;
}
export function deleteField(table,id) {
  if(id==='title') throw new Error('主字段不能删除');
  table.fields=table.fields.filter(f=>f.id!==id);table.records.forEach(r=>delete r[id]);repairViews(table);
}
export function newTable(name) {
  const table={id:uid('table'),name,fields:[{id:'title',name:'标题',type:'text',required:true,width:290},{id:'status',name:'状态',type:'select',options:[...STATUSES],defaultValue:'待开始',width:140},{id:'start',name:'开始日期',type:'date',width:140},{id:'due',name:'截止日期',type:'date',width:140}],records:[],views:[]};
  table.views.push(makeView(table,'grid','表格'));table.activeView=table.views[0].id;return table;
}
// One history covers the entire workspace, including schema and view/table operations.
export class History {
  constructor(document) {this.document=document;this.past=[];this.future=[];}
  snapshot() {return structuredClone(this.document);}
  record(before) {
    if(JSON.stringify(before)===JSON.stringify(this.snapshot())) return false;
    this.past.push(before);if(this.past.length>60)this.past.shift();this.future=[];return true;
  }
  restore(snapshot) {
    // Navigation is preserved when the currently selected table/view still exists.
    const activeTableId=this.document.activeTableId,activeView=this.document.activeView;
    const currentViews=new Map((this.document.tables||[]).map(t=>[t.id,t.activeView]));
    for(const key of Object.keys(this.document)) delete this.document[key];
    Object.assign(this.document,structuredClone(snapshot));
    if(this.document.tables?.some(t=>t.id===activeTableId)) this.document.activeTableId=activeTableId;
    for(const t of this.document.tables||[]) if(t.views.some(v=>v.id===currentViews.get(t.id))) t.activeView=currentViews.get(t.id);
    if(!this.document.tables&&this.document.views?.some(v=>v.id===activeView))this.document.activeView=activeView;
  }
  undo() {if(!this.past.length)return false;this.future.push(this.snapshot());this.restore(this.past.pop());return true;}
  redo() {if(!this.future.length)return false;this.past.push(this.snapshot());this.restore(this.future.pop());return true;}
}
