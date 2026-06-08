// ═══════════════════ 通用函数 ═══════════════════
async function logout(){await fetch('/api/logout',{method:'POST'}); location.href='/login';}
function $(id){return document.getElementById(id);}
function show(el){if(el) el.style.display='block';}
function hide(el){if(el) el.style.display='none';}
function setStatus(id, msg, color){
  const el=$(id);
  if(!el) return;
  el.textContent=msg||'';
  if(color) el.style.color=color;
}
function showInlineError(errorId, msg, statusId){
  const el=$(errorId);
  if(el){
    el.textContent=msg||'操作失败';
    show(el);
  }
  if(statusId) setStatus(statusId,'');
}
function clearInlineError(errorId){
  const el=$(errorId);
  if(!el) return;
  el.textContent='';
  hide(el);
}

function excelLoadingHtml(msg){
  return `<div style="padding:14px 16px;color:#5f6d82;background:#f7f9fc;border:1px dashed #cbd5e1;border-radius:8px;font-size:13px"><span class="spinner">&#8635;</span> ${_escH(msg||'\u6b63\u5728\u8bfb\u53d6 Excel...')}</div>`;
}
function setLoadingStatus(id,msg){
  const el=$(id); if(!el) return;
  el.innerHTML=`<span class="spinner">&#8635;</span> ${_escH(msg||'\u52a0\u8f7d\u4e2d...')}`;
}
function setPlainStatus(id,msg){
  const el=$(id); if(!el) return;
  el.textContent=msg||'';
}
function setSelectPlaceholder(id,msg){
  const el=$(id); if(!el) return;
  el.innerHTML=`<option>${_escH(msg||'\u52a0\u8f7d\u4e2d...')}</option>`;
}
function clearSelectOptions(id,msg){
  const el=$(id); if(!el) return;
  el.innerHTML=`<option>${_escH(msg||'\u5148\u52a0\u8f7d\u5217')}</option>`;
}
async function postFormJson(url, fd){
  const r=await fetch(url,{method:'POST',body:fd});
  let d;
  try{
    d=await r.json();
  }catch(e){
    throw new Error('服务器返回格式异常');
  }
  if(!r.ok && d && !d.error) d.error='请求失败：HTTP '+r.status;
  return d;
}

// ═══════════════════ 左侧导航切换 ═══════════════════
const TOOLS = {
  bom:    {title:'BOM 格式转换',        badge:'v5.10', tpl:'tpl-bom',
           desc:'将客户提供的多种格式 BOM 表自动识别列映射，支持品牌型号合并列/分开列（格式A/B/C），展开为多供应商独立行。'},
  feishu: {title:'飞书优选库+关系库匹配', badge:'v3.0',  tpl:'tpl-feishu',
           desc:'连接飞书内部 API 网关，支持 15 个预置库（优选库 + 对应关系库），多键 AND 匹配（最多 3 对），批量提取字段，输出含来源表格的匹配结果。'},
  'manufacturer-alias': {title:'\u5382\u5546\u547d\u540d\u6620\u5c04\u8868', badge:'\u6620\u5c04\u5e93', tpl:'tpl-manufacturer-alias',
           desc:'\u7ef4\u62a4\u5ba2\u6237\u5382\u5546\u522b\u540d\u3001\u5927\u5c0f\u5199\u53d8\u4f53\u3001\u4e2d\u6587\u540d\u548c\u97f3\u8bd1\u540d\u5230 HQ \u89c4\u8303\u5382\u5546\u540d\u7684\u7cbe\u786e\u6620\u5c04\uff0c\u4f9b\u540e\u7eed\u5339\u914d\u6d41\u7a0b\u590d\u7528\u3002'},
  'pref-rate': {title:'查询BOM优选率', badge:'v1.0', tpl:'tpl-pref-rate',
               desc:'按 HQ料号 在所有优选库缓存中查找优选等级，输出含优选率统计的 Excel 结果文件。'},
  plm:    {title:'转换为上传PLM系统格式', badge:'v1.6',  tpl:'tpl-plm',
           desc:'将整机 BOM 配置表转换为 PLM 系统可导入的标准格式：序号、料号、单耗等25列，主供行填单耗，替代料自动标记主辅BOM标记。'},
  'plm-auto': {title:'PLM网页自动化', badge:'自动化', tpl:'tpl-plm-auto',
           desc:'自动登录 EIP/PLM，按标准流程上传文件、查询并导出结果。当前包含规格型号反查物料。'},
  'bom-compare': {title:'BOM比对工具合集', badge:'v0.1', tpl:'tpl-bom-compare',
           desc:'提供单板HQ BOM版本对比、整机HQ BOM版本对比、Cadence导出BOM对比HQ BOM三个比对子功能。'},
  toolbox: {title:'小工具合集', badge:'工具', tpl:'tpl-toolbox',
           desc:'提供轻量级本地小工具。当前包含文件哈希值计算，可直接在浏览器内得到 MD5。'},
  'bug-report': {title:'Bug提交栏目', badge:'反馈', tpl:'tpl-bug-report',
             desc:'提交工具问题、附件和复现信息，所有记录保存在服务端，团队成员均可查看。'},
  'feature-request': {title:'需求开发工单', badge:'需求', tpl:'tpl-feature-request',
            desc:'提交新功能、流程优化和开发需求，便于后续评估、排期和跟进。'},
  'admin-users': {title:'用户管理', badge:'ADMIN', tpl:'tpl-admin-users',
           desc:'查看注册用户、角色权限、启用状态和最近使用情况。'},
  manual: {title:'工具说明书', badge:'说明', tpl:'tpl-manual',
           desc:'集中展示每个工具的处理逻辑、输入要求、关键规则和推荐使用步骤。'},
};
let curTool = null;
const APP_NOTICE_VERSION = '2026-05-19-manual-docs';

function initRefreshNotice(){
  const notice = $('refreshNotice');
  const ack = $('refreshNoticeAck');
  if(!notice || !ack) return;
  const key = 'bomToolsRefreshNoticeSeen';
  let seen = '';
  try{ seen = localStorage.getItem(key) || ''; }catch(e){}
  if(seen !== APP_NOTICE_VERSION) notice.classList.add('show');
  ack.onclick = function(){
    notice.classList.remove('show');
    try{ localStorage.setItem(key, APP_NOTICE_VERSION); }catch(e){}
  };
}

function showOverview(){
  curTool = null;
  if(location.hash) history.replaceState(null, '', location.pathname + location.search);
  document.querySelectorAll('.sidebar nav a[data-tool]').forEach(x=>x.classList.remove('active'));
  $('toolTitle').textContent = 'BOM Tools';
  $('toolBadge').textContent = '硬件设计辅助平台';
  let html = '<div class="overview-grid">';
  for(const [key, t] of Object.entries(TOOLS)){
    const iconEl = document.querySelector('.sidebar nav a[data-tool="'+key+'"] .icon');
    const icon = iconEl ? iconEl.textContent : '🛠';
    html += '<div class="overview-card" onclick="switchTool(\''+key+'\')">' +
      '<div class="oc-icon">'+icon+'</div>' +
      '<div class="oc-body">' +
        '<div class="oc-title">'+t.title+' <span class="oc-badge">'+t.badge+'</span></div>' +
        '<div class="oc-desc">'+t.desc+'</div>' +
      '</div></div>';
  }
  html += '</div>';
  $('contentArea').innerHTML = html;
}

function switchTool(key){
  const a = document.querySelector('.sidebar nav a[data-tool="'+key+'"]');
  if(a) a.click();
}

document.querySelectorAll('.sidebar nav a[data-tool]').forEach(a=>{
  a.onclick=function(e){
    e.preventDefault();
    const key = this.dataset.tool;
    if(location.hash !== '#'+key) history.pushState(null, '', '#'+key);
    document.querySelectorAll('.sidebar nav a[data-tool]').forEach(x=>x.classList.remove('active'));
    this.classList.add('active');
    curTool = key;
    const t=TOOLS[curTool];
    $('toolTitle').textContent=t.title;
    $('toolBadge').textContent=t.badge;
    $('contentArea').innerHTML=$(t.tpl).innerHTML;
    initTool(curTool);
  };
});

window.addEventListener('popstate', ()=>{
  const key = decodeURIComponent((location.hash || '').replace(/^#/, ''));
  if(key && TOOLS[key]) switchTool(key);
  else showOverview();
});

function initTool(tool){
  if(tool==='bom') initBom();
  else if(tool==='feishu') initFeishu();
  else if(tool==='plm') initPlm();
  else if(tool==='plm-auto') initPlmAuto();
  else if(tool==='bom-compare') initBomCompare();
  else if(tool==='toolbox') initToolbox();
  else if(tool==='bug-report') initBugReport();
  else if(tool==='feature-request') initFeatureRequest();
  else if(tool==='admin-users') initAdminUsers();
  else if(tool==='manufacturer-alias') initManufacturerAlias();
  else if(window._toolInits&&window._toolInits[tool]) window._toolInits[tool]();
}

const HASH_MD5_K = [
  0xd76aa478,0xe8c7b756,0x242070db,0xc1bdceee,0xf57c0faf,0x4787c62a,0xa8304613,0xfd469501,
  0x698098d8,0x8b44f7af,0xffff5bb1,0x895cd7be,0x6b901122,0xfd987193,0xa679438e,0x49b40821,
  0xf61e2562,0xc040b340,0x265e5a51,0xe9b6c7aa,0xd62f105d,0x02441453,0xd8a1e681,0xe7d3fbc8,
  0x21e1cde6,0xc33707d6,0xf4d50d87,0x455a14ed,0xa9e3e905,0xfcefa3f8,0x676f02d9,0x8d2a4c8a,
  0xfffa3942,0x8771f681,0x6d9d6122,0xfde5380c,0xa4beea44,0x4bdecfa9,0xf6bb4b60,0xbebfbc70,
  0x289b7ec6,0xeaa127fa,0xd4ef3085,0x04881d05,0xd9d4d039,0xe6db99e5,0x1fa27cf8,0xc4ac5665,
  0xf4292244,0x432aff97,0xab9423a7,0xfc93a039,0x655b59c3,0x8f0ccc92,0xffeff47d,0x85845dd1,
  0x6fa87e4f,0xfe2ce6e0,0xa3014314,0x4e0811a1,0xf7537e82,0xbd3af235,0x2ad7d2bb,0xeb86d391
];
const HASH_MD5_S = [
  7,12,17,22,7,12,17,22,7,12,17,22,7,12,17,22,
  5,9,14,20,5,9,14,20,5,9,14,20,5,9,14,20,
  4,11,16,23,4,11,16,23,4,11,16,23,4,11,16,23,
  6,10,15,21,6,10,15,21,6,10,15,21,6,10,15,21
];

function hashRotL(value, shift){ return (value << shift) | (value >>> (32 - shift)); }
function hashAdd32(){ return Array.from(arguments).reduce((sum,value)=>(sum + value) >>> 0, 0); }
function hashMd5(buffer){
  const bytes = new Uint8Array(buffer);
  const originalLength = bytes.length;
  const paddedLength = (((originalLength + 8) >>> 6) + 1) << 6;
  const padded = new Uint8Array(paddedLength);
  padded.set(bytes);
  padded[originalLength] = 0x80;
  const bitLength = originalLength * 8;
  for(let i=0;i<8;i++) padded[paddedLength - 8 + i] = Math.floor(bitLength / Math.pow(2, 8 * i)) & 0xff;
  let a0=0x67452301,b0=0xefcdab89,c0=0x98badcfe,d0=0x10325476;
  for(let offset=0;offset<paddedLength;offset+=64){
    const m = new Uint32Array(16);
    for(let i=0;i<16;i++){
      const j = offset + i * 4;
      m[i] = padded[j] | (padded[j+1] << 8) | (padded[j+2] << 16) | (padded[j+3] << 24);
    }
    let a=a0,b=b0,c=c0,d=d0;
    for(let i=0;i<64;i++){
      let f,g;
      if(i<16){ f=(b & c) | ((~b) & d); g=i; }
      else if(i<32){ f=(d & b) | ((~d) & c); g=(5*i+1)%16; }
      else if(i<48){ f=b ^ c ^ d; g=(3*i+5)%16; }
      else { f=c ^ (b | (~d)); g=(7*i)%16; }
      const temp=d;
      d=c; c=b;
      b=hashAdd32(b, hashRotL(hashAdd32(a, f, HASH_MD5_K[i], m[g]), HASH_MD5_S[i]));
      a=temp;
    }
    a0=hashAdd32(a0,a); b0=hashAdd32(b0,b); c0=hashAdd32(c0,c); d0=hashAdd32(d0,d);
  }
  return [a0,b0,c0,d0].flatMap(word=>[word&0xff,(word>>>8)&0xff,(word>>>16)&0xff,(word>>>24)&0xff])
    .map(byte=>byte.toString(16).padStart(2,'0')).join('');
}
function hashFormatSize(bytes){
  if(bytes === 0) return '0 B';
  const units = ['B','KB','MB','GB','TB'];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return (bytes / Math.pow(1024, index)).toFixed(index === 0 ? 0 : 2) + ' ' + units[index];
}
function hashReadFileBuffer(file, onProgress){
  return new Promise((resolve, reject)=>{
    const reader = new FileReader();
    reader.onprogress = function(event){
      if(event.lengthComputable && onProgress){
        onProgress('读取文件', Math.round(event.loaded / event.total * 100));
      }
    };
    reader.onload = function(){ resolve(reader.result); };
    reader.onerror = function(){ reject(reader.error || new Error('文件读取失败')); };
    reader.readAsArrayBuffer(file);
  });
}
function hashCreateMd5Worker(){
  const workerCode = `
const HASH_MD5_K=${JSON.stringify(HASH_MD5_K)};
const HASH_MD5_S=${JSON.stringify(HASH_MD5_S)};
function hashRotL(value, shift){ return (value << shift) | (value >>> (32 - shift)); }
function hashAdd32(){ return Array.from(arguments).reduce((sum,value)=>(sum + value) >>> 0, 0); }
function hashMd5(buffer){
  const bytes = new Uint8Array(buffer);
  const originalLength = bytes.length;
  const paddedLength = (((originalLength + 8) >>> 6) + 1) << 6;
  const padded = new Uint8Array(paddedLength);
  padded.set(bytes);
  padded[originalLength] = 0x80;
  const bitLength = originalLength * 8;
  for(let i=0;i<8;i++) padded[paddedLength - 8 + i] = Math.floor(bitLength / Math.pow(2, 8 * i)) & 0xff;
  let a0=0x67452301,b0=0xefcdab89,c0=0x98badcfe,d0=0x10325476;
  for(let offset=0;offset<paddedLength;offset+=64){
    const m = new Uint32Array(16);
    for(let i=0;i<16;i++){
      const j = offset + i * 4;
      m[i] = padded[j] | (padded[j+1] << 8) | (padded[j+2] << 16) | (padded[j+3] << 24);
    }
    let a=a0,b=b0,c=c0,d=d0;
    for(let i=0;i<64;i++){
      let f,g;
      if(i<16){ f=(b & c) | ((~b) & d); g=i; }
      else if(i<32){ f=(d & b) | ((~d) & c); g=(5*i+1)%16; }
      else if(i<48){ f=b ^ c ^ d; g=(3*i+5)%16; }
      else { f=c ^ (b | (~d)); g=(7*i)%16; }
      const temp=d;
      d=c; c=b;
      b=hashAdd32(b, hashRotL(hashAdd32(a, f, HASH_MD5_K[i], m[g]), HASH_MD5_S[i]));
      a=temp;
    }
    a0=hashAdd32(a0,a); b0=hashAdd32(b0,b); c0=hashAdd32(c0,c); d0=hashAdd32(d0,d);
    if(offset % 1048576 === 0) postMessage({type:'progress', pct:Math.min(99, Math.round(offset / paddedLength * 100))});
  }
  return [a0,b0,c0,d0].flatMap(word=>[word&0xff,(word>>>8)&0xff,(word>>>16)&0xff,(word>>>24)&0xff])
    .map(byte=>byte.toString(16).padStart(2,'0')).join('');
}
onmessage = function(event){
  try{
    postMessage({type:'progress', pct:0});
    const md5 = hashMd5(event.data);
    postMessage({type:'done', md5});
  }catch(err){
    postMessage({type:'error', error:err && err.message ? err.message : String(err)});
  }
};`;
  const url = URL.createObjectURL(new Blob([workerCode], {type:'application/javascript'}));
  const worker = new Worker(url);
  worker._hashWorkerUrl = url;
  return worker;
}
function hashMd5InWorker(buffer, onProgress){
  return new Promise((resolve, reject)=>{
    if(!window.Worker){
      resolve(hashMd5(buffer));
      return;
    }
    const worker = hashCreateMd5Worker();
    worker.onmessage = function(event){
      const data = event.data || {};
      if(data.type === 'progress' && onProgress) onProgress('计算 MD5', data.pct || 0);
      if(data.type === 'done'){
        if(onProgress) onProgress('计算 MD5', 100);
        URL.revokeObjectURL(worker._hashWorkerUrl);
        worker.terminate();
        resolve(data.md5);
      }
      if(data.type === 'error'){
        URL.revokeObjectURL(worker._hashWorkerUrl);
        worker.terminate();
        reject(new Error(data.error || 'MD5 计算失败'));
      }
    };
    worker.onerror = function(event){
      URL.revokeObjectURL(worker._hashWorkerUrl);
      worker.terminate();
      reject(new Error(event.message || 'MD5 计算失败'));
    };
    worker.postMessage(buffer, [buffer]);
  });
}
function hashResultCard(file, hashes){
  const card = document.createElement('article');
  card.className = 'hash-card';
  card.innerHTML =
    '<div class="hash-meta">' +
      '<div class="hash-meta-item"><span class="hash-label">文件名</span><span class="hash-value" data-role="name"></span></div>' +
      '<div class="hash-meta-item"><span class="hash-label">大小</span><span class="hash-value">'+hashFormatSize(file.size)+'</span></div>' +
      '<div class="hash-meta-item"><span class="hash-label">类型</span><span class="hash-value">'+_escH(file.type || '未知')+'</span></div>' +
    '</div>' +
    '<div class="hash-row hash-row-md5"><strong>MD5</strong><code>'+hashes.md5+'</code><button class="btn btn-sm btn-outline" type="button" data-copy="'+hashes.md5+'">复制</button></div>';
  card.querySelector('[data-role="name"]').textContent = file.name;
  return card;
}
async function hashCalculateFile(file){
  return hashCalculateFileWithProgress(file);
}
async function hashCalculateFileWithProgress(file, onProgress){
  const buffer = await hashReadFileBuffer(file, onProgress);
  const md5 = await hashMd5InWorker(buffer, onProgress);
  return {md5:md5};
}
async function hashCopyText(text){
  if(navigator.clipboard && navigator.clipboard.writeText){
    await navigator.clipboard.writeText(text);
    return;
  }
  const ta = document.createElement('textarea');
  ta.value = text;
  ta.style.position = 'fixed';
  ta.style.left = '-9999px';
  document.body.appendChild(ta);
  ta.select();
  document.execCommand('copy');
  ta.remove();
}
function hashUpdateProgress(stage, fileName, pct, note, isError){
  const panel = $('hashProgressPanel');
  const stageEl = $('hashProgressStage');
  const fileEl = $('hashProgressFile');
  const percentEl = $('hashProgressPercent');
  const barEl = $('hashProgressBar');
  const noteEl = $('hashProgressNote');
  if(!panel || !stageEl || !fileEl || !percentEl || !barEl || !noteEl) return;
  const safePct = Math.max(0, Math.min(100, Number.isFinite(pct) ? pct : 0));
  panel.classList.add('show');
  panel.classList.toggle('error', !!isError);
  stageEl.textContent = stage || '处理中';
  fileEl.textContent = fileName || '';
  percentEl.textContent = safePct + '%';
  barEl.style.width = safePct + '%';
  noteEl.textContent = note || '文件不会上传服务器，所有计算都在本地浏览器完成。';
}
async function hashHandleFiles(files){
  const fileList = Array.from(files || []);
  const statusEl = $('hashStatus');
  const resultsEl = $('hashResults');
  const sectionEl = $('hashResultsSection');
  if(!fileList.length || !statusEl || !resultsEl || !sectionEl) return;
  resultsEl.innerHTML = '';
  sectionEl.style.display = 'block';
  statusEl.classList.remove('error');
  statusEl.textContent = '正在计算 '+fileList.length+' 个文件...';
  hashUpdateProgress('准备开始', '共 '+fileList.length+' 个文件', 0, '正在准备读取文件，本工具不会上传文件。');
  try{
    for(let i=0;i<fileList.length;i++){
      const file = fileList[i];
      const fileLabel = '文件 '+(i+1)+'/'+fileList.length+'：'+file.name;
      statusEl.textContent = fileLabel+' | 准备读取';
      const hashes = await hashCalculateFileWithProgress(file, function(stage, pct){
        const pctText = Number.isFinite(pct) ? ' '+pct+'%' : '';
        statusEl.textContent = fileLabel+' | '+stage+pctText;
        const overall = Math.round(((i + (Math.max(0, Math.min(100, pct || 0)) / 100)) / fileList.length) * 100);
        hashUpdateProgress(stage, fileLabel, overall, stage+'：'+Math.max(0, Math.min(100, pct || 0))+'%，总进度 '+overall+'%');
      });
      resultsEl.appendChild(hashResultCard(file, hashes));
    }
    statusEl.textContent = '完成：已计算 '+fileList.length+' 个文件';
    hashUpdateProgress('计算完成', '已计算 '+fileList.length+' 个文件', 100, '结果已生成，可复制 MD5。');
  }catch(err){
    statusEl.classList.add('error');
    statusEl.textContent = '计算失败：' + (err && err.message ? err.message : err);
    hashUpdateProgress('计算失败', '', 100, err && err.message ? err.message : String(err), true);
  }
}
function initToolbox(){
  const dropZone = $('hashDropZone');
  const fileInput = $('hashFileInput');
  const resultsEl = $('hashResults');
  if(!dropZone || !fileInput || !resultsEl) return;
  fileInput.onchange = function(event){
    hashHandleFiles(event.target.files);
    event.target.value = '';
  };
  dropZone.ondragover = function(event){
    event.preventDefault();
    dropZone.classList.add('dragging');
  };
  dropZone.ondragleave = function(){ dropZone.classList.remove('dragging'); };
  dropZone.ondrop = function(event){
    event.preventDefault();
    dropZone.classList.remove('dragging');
    hashHandleFiles(event.dataTransfer.files);
  };
  resultsEl.onclick = async function(event){
    const button = event.target.closest('[data-copy]');
    if(!button) return;
    const original = button.textContent;
    try{
      await hashCopyText(button.dataset.copy || '');
      button.textContent = '已复制';
      setTimeout(()=>{ button.textContent = original; }, 1200);
    }catch(e){
      button.textContent = '复制失败';
      setTimeout(()=>{ button.textContent = original; }, 1200);
    }
  };
}


function adminFmtTime(ts){
  if(!ts) return '';
  const d=new Date(ts*1000);
  const pad=n=>String(n).padStart(2,'0');
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function adminActionLabel(action){
  const map={login:'登录',submit_bug:'提交Bug',submit_feature:'提交需求',like_feature:'点赞需求',update_bug_status:'修改Bug状态',update_feature_status:'修改需求状态',admin_update_user_role:'修改用户角色',admin_update_user_active:'修改用户状态',tool_run:'使用工具',tool_export:'导出文件'};
  return map[action]||action;
}
function adminActivityDetail(a){
  const d=a.detail||{};
  const parts=[];
  if(d.tool) parts.push('工具：'+d.tool);
  if(d.filename) parts.push('文件：'+d.filename);
  if(d.files&&d.files.length) parts.push('文件：'+d.files.map(f=>f.name||f.filename||'').filter(Boolean).join('、'));
  if(d.total!==undefined) parts.push('总数：'+d.total);
  if(d.matched!==undefined) parts.push('命中：'+d.matched);
  if(d.changed!==undefined) parts.push('变更：'+d.changed);
  if(d.error) parts.push('错误：'+d.error);
  return parts.join(' | ');
}
async function adminLoadUsers(){
  const status=$('adminUserStatus');
  const q=$('adminUserQuery')?$('adminUserQuery').value.trim():'';
  if(status) status.textContent='加载中...';
  try{
    const r=await fetch('/api/admin/users?q='+encodeURIComponent(q));
    const d=await r.json();
    if(!d.success) throw new Error(d.error||'加载失败');
    const s=d.summary||{};
    $('adminUserTotal').textContent='总用户 '+(s.total||0);
    $('adminUserActive').textContent='启用 '+(s.active||0);
    $('adminUserDisabled').textContent='禁用 '+(s.disabled||0);
    $('adminUserAdmins').textContent='管理员 '+(s.admins||0);
    const rows=(d.users||[]).map(u=>`<tr>
      <td>${_escH(u.employee_id)}</td><td>${_escH(u.display_name)}</td>
      <td><select data-admin-role="${_escH(u.id)}"><option value="user" ${u.role==='user'?'selected':''}>普通用户</option><option value="admin" ${u.role==='admin'?'selected':''}>管理员</option></select></td>
      <td>${u.is_active?'<span class="badge-sm badge-green">启用</span>':'<span class="badge-sm badge-gray">禁用</span>'}</td>
      <td>${adminFmtTime(u.created_at)}</td><td>${adminFmtTime(u.last_login_at)}</td>
      <td>${u.login_count||0}</td><td>${u.bug_submit_count||0}</td><td>${u.feature_submit_count||0}</td><td>${u.feature_like_count||0}</td><td>${u.status_update_count||0}</td>
      <td><button class="btn btn-sm btn-gray" data-admin-active="${_escH(u.id)}" data-active="${u.is_active?0:1}">${u.is_active?'禁用':'启用'}</button></td>
    </tr>`).join('');
    $('adminUserRows').innerHTML=rows||'<tr><td colspan="12">暂无用户</td></tr>';
    document.querySelectorAll('[data-admin-role]').forEach(el=>{el.onchange=()=>adminSetRole(el.dataset.adminRole,el.value);});
    document.querySelectorAll('[data-admin-active]').forEach(el=>{el.onclick=()=>adminSetActive(el.dataset.adminActive,el.dataset.active==='1');});
    if(status) status.textContent='已加载 '+(d.users||[]).length+' 个用户';
  }catch(e){ if(status) status.textContent=e.message; }
}
async function adminLoadActivity(){
  try{
    const r=await fetch('/api/admin/activity?limit=30');
    const d=await r.json();
    if(!d.success) throw new Error(d.error||'加载失败');
    const rows=(d.activities||[]).map(a=>{
      const detail=adminActivityDetail(a);
      const target=[a.target_type,a.target_id].filter(Boolean).join(' / ');
      return `<tr>
        <td>${adminFmtTime(a.created_at)}</td>
        <td>${_escH(a.employee_id)}</td>
        <td>${_escH(a.display_name)}</td>
        <td>${_escH(adminActionLabel(a.action))}</td>
        <td>${_escH(target)}</td>
        <td>${_escH(detail)}</td>
      </tr>`;
    }).join('');
    $('adminActivityRows').innerHTML=rows||'<tr><td colspan="6">暂无活动</td></tr>';
  }catch(e){ $('adminActivityRows').innerHTML='<tr><td colspan="6" style="color:#c00000">'+_escH(e.message)+'</td></tr>'; }
}
async function adminSetRole(userId,role){
  if(!confirm('确认修改用户角色？')){adminLoadUsers();return;}
  const r=await fetch('/api/admin/users/'+encodeURIComponent(userId)+'/role',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({role})});
  const d=await r.json();
  if(!d.success) alert(d.error||'修改失败');
  adminLoadUsers();adminLoadActivity();
}
async function adminSetActive(userId,isActive){
  if(!confirm(isActive?'确认启用该用户？':'确认禁用该用户？')) return;
  const r=await fetch('/api/admin/users/'+encodeURIComponent(userId)+'/active',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({is_active:isActive})});
  const d=await r.json();
  if(!d.success) alert(d.error||'修改失败');
  adminLoadUsers();adminLoadActivity();
}
function initAdminUsers(){
  if($('adminUserRefresh')) $('adminUserRefresh').onclick=function(){adminLoadUsers();adminLoadActivity();};
  if($('adminUserQuery')) $('adminUserQuery').onkeydown=e=>{if(e.key==='Enter')adminLoadUsers();};
  adminLoadUsers();
  adminLoadActivity();
}
initRefreshNotice();

function bugFmtTime(ts){
  if(!ts) return '';
  const d=new Date(ts*1000);
  const pad=n=>String(n).padStart(2,'0');
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function bugStatusClass(status){
  const s=String(status||'');
  if(s==='处理中') return 'badge-status-progress';
  if(s==='已修复') return 'badge-status-fixed';
  if(s==='已关闭') return 'badge-status-closed';
  if(s==='暂缓') return 'badge-status-paused';
  if(s==='无法复现') return 'badge-status-invalid';
  return 'badge-status-pending';
}

function featureStatusClass(status){
  const s=String(status||'');
  if(s==='已纳入') return 'badge-status-progress';
  if(s==='开发中') return 'badge-status-progress';
  if(s==='已完成') return 'badge-status-fixed';
  if(s==='已关闭') return 'badge-status-closed';
  if(s==='暂缓') return 'badge-status-paused';
  return 'badge-status-pending';
}



async function mfgLoadAliases(query){
  const q = query !== undefined ? query : $('mfgSearch').value.trim();
  $('mfgListStatus').textContent='\u52a0\u8f7d\u4e2d...';
  try{
    const r=await fetch('/api/manufacturer_aliases?q='+encodeURIComponent(q));
    const d=await r.json();
    if(!d.success) throw new Error(d.error||'\u52a0\u8f7d\u5931\u8d25');
    const rows=d.aliases||[];
    $('mfgListStatus').textContent=`\u5171 ${rows.length} \u6761\u8bb0\u5f55`;
    if(d.match){
      $('mfgMatch').style.display='block';
      $('mfgMatch').innerHTML=`\u5339\u914d\u5230 HQ \u89c4\u8303\u5382\u5546\u540d\uff1a<b>${_escH(d.match.canonical_name)}</b> &nbsp; \u522b\u540d\uff1a${_escH(d.match.alias)} &nbsp; \u6765\u6e90\uff1a${_escH(d.match.source||'')}`;
    }else{
      $('mfgMatch').style.display=q?'block':'none';
      $('mfgMatch').innerHTML=q?`\u672a\u627e\u5230\u7cbe\u786e\u6620\u5c04\uff0c\u5f52\u4e00\u5316\u952e\uff1a<b>${_escH(d.normalized_query||'')}</b>`:'';
    }
    if(!rows.length){
      $('mfgRows').innerHTML='<tr><td colspan="6" style="color:#888;text-align:center;padding:18px">\u6682\u65e0\u6620\u5c04\u8bb0\u5f55</td></tr>';
      return;
    }
    $('mfgRows').innerHTML=rows.map(item=>`<tr>
      <td>${_escH(item.canonical_name)}</td>
      <td>${_escH(item.alias)}</td>
      <td style="color:#667085">${_escH(item.normalized_alias)}</td>
      <td>${_escH(item.source||'')}</td>
      <td>${_escH(item.note||'')}</td>
      <td><button class="btn btn-sm btn-gray" type="button" onclick="mfgDeleteAlias('${_escH(item.id)}')" style="color:#c00000">\u5220\u9664</button></td>
    </tr>`).join('');
  }catch(e){
    $('mfgListStatus').textContent='';
    $('mfgRows').innerHTML=`<tr><td colspan="6"><div class="error show">${_escH(e.message)}</div></td></tr>`;
  }
}

async function mfgAddAlias(){
  const fd=new FormData();
  fd.append('canonical_name',$('mfgCanonical').value.trim());
  fd.append('alias',$('mfgAlias').value.trim());
  fd.append('source',$('mfgSource').value.trim());
  fd.append('note',$('mfgNote').value.trim());
  $('mfgAdd').disabled=true;$('mfgAddStatus').textContent='\u4fdd\u5b58\u4e2d...';clearInlineError('mfgError');
  try{
    const r=await fetch('/api/manufacturer_aliases',{method:'POST',body:fd});
    const d=await r.json();
    if(!d.success){
      let msg=d.error||'\u4fdd\u5b58\u5931\u8d25';
      if(d.existing) msg += `\uff1a${d.existing.alias} \u2192 ${d.existing.canonical_name}`;
      throw new Error(msg);
    }
    $('mfgAddStatus').textContent='\u5df2\u4fdd\u5b58';
    $('mfgAlias').value='';$('mfgNote').value='';
    await mfgLoadAliases($('mfgSearch').value.trim());
  }catch(e){showInlineError('mfgError',e.message,'mfgAddStatus');}
  $('mfgAdd').disabled=false;
}

async function mfgDeleteAlias(id){
  if(!id || !confirm('\u786e\u5b9a\u5220\u9664\u8fd9\u6761\u5382\u5546\u6620\u5c04\uff1f')) return;
  $('mfgListStatus').textContent='\u5220\u9664\u4e2d...';
  try{
    const r=await fetch('/api/manufacturer_aliases/'+encodeURIComponent(id),{method:'DELETE'});
    const d=await r.json();
    if(!d.success) throw new Error(d.error||'\u5220\u9664\u5931\u8d25');
    await mfgLoadAliases($('mfgSearch').value.trim());
  }catch(e){$('mfgListStatus').textContent=e.message;}
}

function initManufacturerAlias(){
  mfgLoadAliases('');
  $('mfgSearchBtn').onclick=()=>mfgLoadAliases($('mfgSearch').value.trim());
  $('mfgRefresh').onclick=()=>mfgLoadAliases('');
  $('mfgSearch').onkeydown=e=>{if(e.key==='Enter')mfgLoadAliases($('mfgSearch').value.trim());};
}

async function bugLoadReports(){
  $('bugListStatus').textContent='加载中...';
  try{
    const params=new URLSearchParams();
    const status=$('bugFilterStatus') ? $('bugFilterStatus').value.trim() : '';
    const module=$('bugFilterModule') ? $('bugFilterModule').value.trim() : '';
    const q=$('bugFilterQuery') ? $('bugFilterQuery').value.trim() : '';
    if(status) params.set('status',status);
    if(module) params.set('module',module);
    if(q) params.set('q',q);
    const r=await fetch('/api/bug_reports'+(params.toString()?('?'+params.toString()):''));
    const d=await r.json();
    if(!d.success) throw new Error(d.error||'加载失败');
    const reports=d.reports||[];
    $('bugListStatus').textContent=`共 ${reports.length} 条记录`;
    if(!reports.length){
      $('bugList').innerHTML='<div style="font-size:13px;color:#888;padding:16px;background:#f7f9fc;border-radius:8px">暂无问题记录</div>';
      return;
    }
    window._bugReportsById={};
    reports.forEach(item=>{window._bugReportsById[item.id]=item;});
    $('bugList').innerHTML=reports.map(item=>{
      const imgs=(item.attachments||[]).map(a=>`<a href="${_escH(a.url)}" target="_blank" onclick="event.stopPropagation()" style="font-size:12px;color:#1a5ad4;margin-right:8px">${_escH(a.name||'截图')}</a>`).join('');
      const statusClass=bugStatusClass(item.status);
      return `<div onclick="bugOpenReport('${_escH(item.id)}')" style="border:1px solid #e2e8f3;border-radius:8px;padding:12px;background:#fff;cursor:pointer">
        <div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start">
          <div style="font-weight:700;color:#1a3a5c">${_escH(item.title)}</div>
          <span class="badge-sm ${statusClass}">${_escH(item.status||'待处理')}</span>
        </div>
        <div style="font-size:12px;color:#667085;margin-top:5px">${_escH(item.module)} | ${_escH(item.severity)} | ${_escH(item.reporter)}（${_escH(item.employee_id)}） | ${bugFmtTime(item.submitted_at)}</div>
        <div style="font-size:13px;color:#333;line-height:1.6;margin-top:8px;white-space:pre-wrap">${_escH(item.description)}</div>
        ${item.steps?`<div style="font-size:12px;color:#555;margin-top:8px"><b>复现步骤：</b><div style="white-space:pre-wrap">${_escH(item.steps)}</div></div>`:''}
        ${item.expected?`<div style="font-size:12px;color:#555;margin-top:8px"><b>期望结果：</b><div style="white-space:pre-wrap">${_escH(item.expected)}</div></div>`:''}
        ${imgs?`<div style="margin-top:8px"><b style="font-size:12px;color:#555">附件：</b>${imgs}</div>`:''}
        <div style="font-size:12px;color:#1a5ad4;margin-top:8px">点击查看详情 / 修改处理状态</div>
      </div>`;
    }).join('');
  }catch(e){
    $('bugListStatus').textContent='';
    $('bugList').innerHTML=`<div class="error show">${_escH(e.message)}</div>`;
  }
}

function bugOpenReport(id){
  const item=(window._bugReportsById||{})[id];
  if(!item) return;
  const files=(item.attachments||[]).map(a=>`<a href="${_escH(a.url)}" target="_blank" style="font-size:12px;color:#1a5ad4;margin-right:8px">${_escH(a.name||'附件')}</a>`).join('');
  $('bugDetailTitle').textContent=item.title||'';
  const detailStatus=item.status||'待处理';
  $('bugDetailMeta').innerHTML=`${_escH(item.module||'')} | ${_escH(item.severity||'')} | ${_escH(item.reporter||'')}（${_escH(item.employee_id||'')}） | ${bugFmtTime(item.submitted_at)} <span class="badge-sm ${bugStatusClass(detailStatus)}" style="margin-left:8px">${_escH(detailStatus)}</span>`;
  $('bugDetailBody').innerHTML=`<div style="font-size:13px;color:#333;line-height:1.7;white-space:pre-wrap">${_escH(item.description)}</div>${item.steps?`<div style="font-size:13px;color:#555;margin-top:10px"><b>复现步骤：</b><div style="white-space:pre-wrap">${_escH(item.steps)}</div></div>`:''}${item.expected?`<div style="font-size:13px;color:#555;margin-top:10px"><b>期望结果：</b><div style="white-space:pre-wrap">${_escH(item.expected)}</div></div>`:''}${files?`<div style="margin-top:10px"><b style="font-size:13px;color:#555">附件：</b>${files}</div>`:''}`;
  $('bugDetailStatus').value=item.status||'待处理';
  $('bugDetailSave').dataset.id=id;
  $('bugDetailMsg').textContent='';
  $('bugDetailModal').style.display='flex';
}
function bugCloseReport(){
  if($('bugDetailModal')) $('bugDetailModal').style.display='none';
}
async function bugSaveStatus(){
  const id=$('bugDetailSave').dataset.id;
  const status=$('bugDetailStatus').value;
  $('bugDetailSave').disabled=true;$('bugDetailMsg').textContent='保存中...';
  try{
    const r=await fetch('/api/bug_reports/'+encodeURIComponent(id)+'/status',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({status})});
    const d=await r.json();
    if(!d.success) throw new Error(d.error||'保存失败');
    $('bugDetailMsg').textContent='已保存';
    await bugLoadReports();
    if(d.report) bugOpenReport(d.report.id);
  }catch(e){$('bugDetailMsg').textContent=e.message;}
  $('bugDetailSave').disabled=false;
}

function initBugReport(){
  bugLoadReports();
  $('bugRefresh').onclick=bugLoadReports;
  ['bugFilterStatus','bugFilterModule'].forEach(id=>{ if($(id)) $(id).onchange=bugLoadReports; });
  if($('bugFilterQuery')) $('bugFilterQuery').onkeydown=e=>{if(e.key==='Enter')bugLoadReports();};
  $('bugDetailClose').onclick=bugCloseReport;
  $('bugDetailSave').onclick=bugSaveStatus;
  $('bugDetailModal').onclick=function(e){if(e.target===this) bugCloseReport();};
  $('bugSubmit').onclick=async function(){
    const fd=new FormData();
    fd.append('reporter',$('bugReporter').value.trim());
    fd.append('employee_id',$('bugEmployeeId').value.trim());
    fd.append('module',$('bugModule').value);
    fd.append('severity',$('bugSeverity').value);
    fd.append('title',$('bugTitle').value.trim());
    fd.append('description',$('bugDesc').value.trim());
    fd.append('steps',$('bugSteps').value.trim());
    fd.append('expected',$('bugExpected').value.trim());
    [...$('bugImages').files].forEach(f=>fd.append('images',f));
    $('bugSubmit').disabled=true;$('bugStatus').textContent='提交中...';clearInlineError('bugError');
    try{
      const r=await fetch('/api/bug_reports',{method:'POST',body:fd});
      const d=await r.json();
      if(d.success){
        $('bugStatus').textContent='已提交';
        ['bugTitle','bugDesc','bugSteps','bugExpected'].forEach(id=>$(id).value='');
        $('bugImages').value='';
        await bugLoadReports();
      }else{
        showInlineError('bugError',d.error||'提交失败','bugStatus');
      }
    }catch(e){showInlineError('bugError',e.message,'bugStatus');}
    $('bugSubmit').disabled=false;
  };
}




async function featureLoadRequests(){
  $('featureListStatus').textContent='加载中...';
  try{
    const params=new URLSearchParams();
    const status=$('featureFilterStatus') ? $('featureFilterStatus').value.trim() : '';
    const module=$('featureFilterModule') ? $('featureFilterModule').value.trim() : '';
    const q=$('featureFilterQuery') ? $('featureFilterQuery').value.trim() : '';
    const sort=$('featureSort') ? $('featureSort').value.trim() : '';
    if(status) params.set('status',status);
    if(module) params.set('module',module);
    if(q) params.set('q',q);
    if(sort) params.set('sort',sort);
    const r=await fetch('/api/feature_requests'+(params.toString()?('?'+params.toString()):''));
    const d=await r.json();
    if(!d.success) throw new Error(d.error||'加载失败');
    const requests=d.requests||[];
    $('featureListStatus').textContent=`共 ${requests.length} 条记录`;
    if(!requests.length){$('featureList').innerHTML='<div style="font-size:13px;color:#888;padding:16px;background:#f7f9fc;border-radius:8px">暂无需求记录</div>';return;}
    window._featureRequestsById={};
    requests.forEach(item=>{window._featureRequestsById[item.id]=item;});
    $('featureList').innerHTML=requests.map(item=>{
      const files=(item.attachments||[]).map(a=>`<a href="${_escH(a.url)}" target="_blank" onclick="event.stopPropagation()" style="font-size:12px;color:#1a5ad4;margin-right:8px">${_escH(a.name||'附件')}</a>`).join('');
      const likeCount=Number(item.likes||0);
      const status=item.status||'\u5f85\u8bc4\u4f30';
      return `<div onclick="featureOpenRequest('${_escH(item.id)}')" style="border:1px solid #e2e8f3;border-radius:8px;padding:12px;background:#fff;cursor:pointer"><div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start"><div style="font-weight:700;color:#1a3a5c">${_escH(item.title)}</div><div style="display:flex;gap:8px;align-items:center;flex-shrink:0"><span class="badge-sm ${featureStatusClass(status)}">${_escH(status)}</span><button class="btn btn-sm btn-gray" type="button" onclick="event.stopPropagation();featureLikeRequest('${_escH(item.id)}')" style="padding:2px 8px;font-size:12px">\u70b9\u8d5e ${likeCount}</button></div></div><div style="font-size:12px;color:#667085;margin-top:5px">${_escH(item.module)} | ${_escH(item.request_type)} | ${_escH(item.priority)} | ${_escH(item.requester)}\uff08${_escH(item.employee_id)}\uff09 | ${bugFmtTime(item.submitted_at)}</div>${item.background?`<div style="font-size:12px;color:#555;margin-top:8px"><b>\u9700\u6c42\u80cc\u666f\uff1a</b><div style="white-space:pre-wrap">${_escH(item.background)}</div></div>`:''}<div style="font-size:13px;color:#333;line-height:1.6;margin-top:8px;white-space:pre-wrap">${_escH(item.requirement)}</div>${item.value?`<div style="font-size:12px;color:#555;margin-top:8px"><b>\u9884\u671f\u4ef7\u503c\uff1a</b><div style="white-space:pre-wrap">${_escH(item.value)}</div></div>`:''}${item.acceptance?`<div style="font-size:12px;color:#555;margin-top:8px"><b>\u9a8c\u6536\u6807\u51c6\uff1a</b><div style="white-space:pre-wrap">${_escH(item.acceptance)}</div></div>`:''}${files?`<div style="margin-top:8px"><b style="font-size:12px;color:#555">\u9644\u4ef6\uff1a</b>${files}</div>`:''}<div style="font-size:12px;color:#1a5ad4;margin-top:8px">\u70b9\u51fb\u67e5\u770b\u8be6\u60c5 / \u4fee\u6539\u9700\u6c42\u72b6\u6001</div></div>`;
    }).join('');
  }catch(e){$('featureListStatus').textContent='';$('featureList').innerHTML=`<div class="error show">${_escH(e.message)}</div>`;}
}

async function featureLikeRequest(id){
  try{
    const employeeId=($('featureEmployeeId')&&$('featureEmployeeId').value.trim())||'';
    const r=await fetch('/api/feature_requests/'+encodeURIComponent(id)+'/like',{
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({employee_id:employeeId})
    });
    const d=await r.json();
    if(!d.success) throw new Error(d.error||'点赞失败');
    await featureLoadRequests();
    if(d.already_liked) $('featureListStatus').textContent=d.message||'你已经点赞过该需求';
  }catch(e){
    $('featureListStatus').textContent=e.message;
  }
}

function featureOpenRequest(id){
  const item=(window._featureRequestsById||{})[id];
  if(!item) return;
  const files=(item.attachments||[]).map(a=>`<a href="${_escH(a.url)}" target="_blank" style="font-size:12px;color:#1a5ad4;margin-right:8px">${_escH(a.name||'\u9644\u4ef6')}</a>`).join('');
  const status=item.status||'\u5f85\u8bc4\u4f30';
  $('featureDetailTitle').textContent=item.title||'';
  $('featureDetailMeta').innerHTML=`${_escH(item.module||'')} | ${_escH(item.request_type||'')} | ${_escH(item.priority||'')} | ${_escH(item.requester||'')}\uff08${_escH(item.employee_id||'')}\uff09 | ${bugFmtTime(item.submitted_at)} <span class="badge-sm ${featureStatusClass(status)}" style="margin-left:8px">${_escH(status)}</span>`;
  $('featureDetailBody').innerHTML=`${item.background?`<div style="font-size:13px;color:#555;margin-top:10px"><b>\u9700\u6c42\u80cc\u666f\uff1a</b><div style="white-space:pre-wrap">${_escH(item.background)}</div></div>`:''}<div style="font-size:13px;color:#333;line-height:1.7;white-space:pre-wrap;margin-top:10px">${_escH(item.requirement||'')}</div>${item.value?`<div style="font-size:13px;color:#555;margin-top:10px"><b>\u9884\u671f\u4ef7\u503c\uff1a</b><div style="white-space:pre-wrap">${_escH(item.value)}</div></div>`:''}${item.acceptance?`<div style="font-size:13px;color:#555;margin-top:10px"><b>\u9a8c\u6536\u6807\u51c6\uff1a</b><div style="white-space:pre-wrap">${_escH(item.acceptance)}</div></div>`:''}${files?`<div style="margin-top:10px"><b style="font-size:13px;color:#555">\u9644\u4ef6\uff1a</b>${files}</div>`:''}`;
  $('featureDetailStatus').value=status;
  $('featureDetailSave').dataset.id=id;
  $('featureDetailMsg').textContent='';
  $('featureDetailModal').style.display='flex';
}
function featureCloseRequest(){
  if($('featureDetailModal')) $('featureDetailModal').style.display='none';
}
async function featureSaveStatus(){
  const id=$('featureDetailSave').dataset.id;
  const status=$('featureDetailStatus').value;
  $('featureDetailSave').disabled=true;$('featureDetailMsg').textContent='\u4fdd\u5b58\u4e2d...';
  try{
    const r=await fetch('/api/feature_requests/'+encodeURIComponent(id)+'/status',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({status})});
    const d=await r.json();
    if(!d.success) throw new Error(d.error||'\u4fdd\u5b58\u5931\u8d25');
    $('featureDetailMsg').textContent='\u5df2\u4fdd\u5b58';
    await featureLoadRequests();
    if(d.request) featureOpenRequest(d.request.id);
  }catch(e){$('featureDetailMsg').textContent=e.message;}
  $('featureDetailSave').disabled=false;
}

function initFeatureRequest(){
  featureLoadRequests();
  $('featureRefresh').onclick=featureLoadRequests;
  ['featureFilterStatus','featureFilterModule','featureSort'].forEach(id=>{ if($(id)) $(id).onchange=featureLoadRequests; });
  if($('featureFilterQuery')) $('featureFilterQuery').onkeydown=e=>{if(e.key==='Enter')featureLoadRequests();};
  $('featureDetailClose').onclick=featureCloseRequest;
  $('featureDetailSave').onclick=featureSaveStatus;
  $('featureDetailModal').onclick=function(e){if(e.target===this) featureCloseRequest();};
  $('featureSubmit').onclick=async function(){
    const fd=new FormData();
    fd.append('requester',$('featureRequester').value.trim());fd.append('employee_id',$('featureEmployeeId').value.trim());fd.append('module',$('featureModule').value);fd.append('priority',$('featurePriority').value);fd.append('request_type',$('featureType').value);fd.append('title',$('featureTitle').value.trim());fd.append('background',$('featureBackground').value.trim());fd.append('requirement',$('featureRequirement').value.trim());fd.append('value',$('featureValue').value.trim());fd.append('acceptance',$('featureAcceptance').value.trim());[...$('featureFiles').files].forEach(f=>fd.append('attachments',f));
    $('featureSubmit').disabled=true;$('featureStatus').textContent='提交中...';clearInlineError('featureError');
    try{const r=await fetch('/api/feature_requests',{method:'POST',body:fd});const d=await r.json();if(d.success){$('featureStatus').textContent='已提交';['featureTitle','featureBackground','featureRequirement','featureValue','featureAcceptance'].forEach(id=>$(id).value='');$('featureFiles').value='';await featureLoadRequests();}else{showInlineError('featureError',d.error||'提交失败','featureStatus');}}
    catch(e){showInlineError('featureError',e.message,'featureStatus');}
    $('featureSubmit').disabled=false;
  };
}

function cmpState(prefix){
  const key='_cmp_'+prefix;
  if(!window[key]) window[key]={leftHeaders:[],rightHeaders:[],manualPairs:[],keyPairs:[]};
  return window[key];
}
function cmpOptions(headers, selected){
  return (headers||[]).map(h=>`<option value="${_escH(h)}"${h===selected?' selected':''}>${_escH(h)}</option>`).join('');
}

function cmpLooksLikeManufacturerCol(name){
  return /\u5382\u5546|\u5382\u5bb6|\u5236\u9020\u5546|\u751f\u4ea7\u5382\u5bb6|\u4f9b\u5e94\u5546|brand|manufacturer|maker/i.test(name||'');
}
function cmpDefaultKeyPairs(prefix){
  const st=cmpState(prefix);
  if(prefix!=='cust') return [];
  const left=st.leftHeaders||[], right=st.rightHeaders||[];
  const pick=(headers, patterns)=>headers.find(h=>patterns.some(p=>p.test(h||'')))||'';
  const leftModel=pick(left, [/\u89c4\u683c\u578b\u53f7/, /\u578b\u53f7/, /part\s*number/i, /model/i]);
  const rightModel=pick(right, [/\u578b\u53f7/, /\u89c4\u683c\u578b\u53f7/, /part\s*number/i, /model/i]);
  const leftMaker=pick(left, [/\u5382\u5546/, /\u5382\u5bb6/, /\u5236\u9020\u5546/, /\u751f\u4ea7\u5382\u5bb6/, /\u4f9b\u5e94\u5546/, /brand/i, /manufacturer/i, /maker/i]);
  const rightMaker=pick(right, [/\u751f\u4ea7\u5382\u5bb6/, /\u5382\u5546/, /\u5382\u5bb6/, /\u5236\u9020\u5546/, /\u4f9b\u5e94\u5546/, /brand/i, /manufacturer/i, /maker/i]);
  const rows=[];
  if(leftModel&&rightModel) rows.push({left:leftModel,right:rightModel,transform:''});
  if(leftMaker&&rightMaker) rows.push({left:leftMaker,right:rightMaker,transform:'manufacturer_alias'});
  return rows.length?rows:[{left:left[0]||'',right:right[0]||'',transform:''}];
}
function cmpEnsureKeyPairs(prefix){
  const st=cmpState(prefix);
  if(prefix!=='cust') return;
  st.keyPairs=(st.keyPairs||[]).filter(p=>p.left||p.right);
  if(!st.keyPairs.length) st.keyPairs=cmpDefaultKeyPairs(prefix);
  if(!st.keyPairs.length) st.keyPairs=[{left:'',right:'',transform:''}];
}
function cmpRenderKeyPairs(prefix){
  if(prefix!=='cust' || !$(prefix+'KeyRows')) return;
  const st=cmpState(prefix);
  cmpEnsureKeyPairs(prefix);
  const left=st.leftHeaders||[], right=st.rightHeaders||[];
  $(prefix+'KeyRows').innerHTML=(st.keyPairs||[]).map((pair,i)=>{
    const showMap=cmpLooksLikeManufacturerCol(pair.left);
    const checked=pair.transform==='manufacturer_alias'?' checked':'';
    return `<div class="row" style="margin:0">
      <label style="min-width:90px">\u5339\u914d\u952e${i+1}\uff1a</label>
      <select data-i="${i}" data-side="left" style="width:220px"><option value="">\u5148\u52a0\u8f7d\u5217</option>${cmpOptions(left,pair.left)}</select>
      <span>&harr;</span>
      <select data-i="${i}" data-side="right" style="width:220px"><option value="">\u5148\u52a0\u8f7d\u5217</option>${cmpOptions(right,pair.right)}</select>
      <label style="font-size:12px;color:#555;display:${showMap?'flex':'none'};align-items:center;gap:3px;white-space:nowrap"><input type="checkbox" data-i="${i}" data-transform="manufacturer_alias"${checked}>\u5382\u5546\u6620\u5c04</label>
      ${(st.keyPairs||[]).length>1?`<button class="btn btn-sm btn-gray" type="button" data-remove="${i}" style="color:#c00000">\u5220\u9664</button>`:''}
    </div>`;
  }).join('');
  $(prefix+'KeyRows').querySelectorAll('select').forEach(sel=>{
    sel.onchange=function(){const p=st.keyPairs[parseInt(this.dataset.i)];p[this.dataset.side]=this.value;if(!cmpLooksLikeManufacturerCol(p.left))p.transform='';cmpRenderKeyPairs(prefix);cmpRenderPairs(prefix);};
  });
  $(prefix+'KeyRows').querySelectorAll('input[data-transform]').forEach(chk=>{
    chk.onchange=function(){st.keyPairs[parseInt(this.dataset.i)].transform=this.checked?'manufacturer_alias':'';};
  });
  $(prefix+'KeyRows').querySelectorAll('button[data-remove]').forEach(btn=>{
    btn.onclick=function(){st.keyPairs.splice(parseInt(this.dataset.remove),1);cmpRenderKeyPairs(prefix);cmpRenderPairs(prefix);};
  });
  if($(prefix+'LeftKey')) $(prefix+'LeftKey').value=(st.keyPairs[0]||{}).left||'';
  if($(prefix+'RightKey')) $(prefix+'RightKey').value=(st.keyPairs[0]||{}).right||'';
}
function cmpCollectKeyConfig(prefix){
  if(prefix==='cust') return {left_key_cols:[],right_key_cols:[],left_key_transforms:[]};
  return {
    left_key_cols: [$(prefix+'LeftKey')?.value].filter(Boolean),
    right_key_cols: [$(prefix+'RightKey')?.value].filter(Boolean),
    left_key_transforms: [''],
  };
}
function cmpPopulateCustomerStandardMapping(prefix){
  if(prefix!=='cust') return;
  const headers=cmpState(prefix).leftHeaders||[];
  ['Refdes','Manufacturer','Model'].forEach(name=>{
    const el=$(prefix+'Std'+name);
    if(el) el.innerHTML=`<option value="">${name==='Refdes'?'不使用位号':'请选择客户列'}</option>`+cmpOptions(headers,'');
  });
}
function cmpCollectCustomerStandardMapping(prefix){
  return {
    refdes: $(prefix+'StdRefdes')?.value||'',
    manufacturer: $(prefix+'StdManufacturer')?.value||'',
    model: $(prefix+'StdModel')?.value||'',
  };
}
function cmpRenderPairs(prefix){
  const st=cmpState(prefix);
  const left=st.leftHeaders||[], right=st.rightHeaders||[];
  if(prefix==='cust'){
    if($(prefix+'CommonPairs')) $(prefix+'CommonPairs').innerHTML='';
    if($(prefix+'ManualPairs')) $(prefix+'ManualPairs').innerHTML='';
    return;
  }
  const common=left.filter(h=>h && right.includes(h));
  const keyCfg=cmpCollectKeyConfig(prefix);
  const leftKey=(keyCfg.left_key_cols||[])[0]||'';
  const rightKey=(keyCfg.right_key_cols||[])[0]||'';
  let html='';
  common.forEach(col=>{
    const checked=(col===leftKey && col===rightKey)?'':' checked';
    html+=`<label><input type="checkbox" value="${_escH(col)}"${checked}>${_escH(col)}</label>`;
  });
  $(prefix+'CommonPairs').innerHTML=html || '<span style="font-size:12px;color:#888">没有同名字段，可在下方添加自定义映射。</span>';
  cmpRenderManualPairs(prefix);
}
function cmpRenderManualPairs(prefix){
  const st=cmpState(prefix);
  const left=st.leftHeaders||[], right=st.rightHeaders||[];
  const rows=st.manualPairs||[];
  $(prefix+'ManualPairs').innerHTML=rows.map((pair,i)=>`<div class="row" style="margin:0">
    <label style="min-width:90px">自定义字段：</label>
    <select data-i="${i}" data-side="left" style="width:220px"><option value="">不选择</option>${cmpOptions(left,pair.left)}</select>
    <span>&harr;</span>
    <select data-i="${i}" data-side="right" style="width:220px"><option value="">不选择</option>${cmpOptions(right,pair.right)}</select>
    <button class="btn btn-sm btn-gray" type="button" data-remove="${i}" style="color:#c00000">删除</button>
  </div>`).join('');
  $(prefix+'ManualPairs').querySelectorAll('select').forEach(sel=>{
    sel.onchange=function(){st.manualPairs[parseInt(this.dataset.i)][this.dataset.side]=this.value;};
  });
  $(prefix+'ManualPairs').querySelectorAll('button[data-remove]').forEach(btn=>{
    btn.onclick=function(){st.manualPairs.splice(parseInt(this.dataset.remove),1);cmpRenderManualPairs(prefix);};
  });
}
async function cmpRefresh(prefix, compareType, apiPrefix='/api/bom_compare'){
  const leftFile=$(prefix+'LeftFile').files[0], rightFile=$(prefix+'RightFile').files[0];
  const st=cmpState(prefix);
  st.leftHeaders=[];st.rightHeaders=[];st.manualPairs=[];st.keyPairs=[];
  if($(prefix+'LeftKey')) clearSelectOptions(prefix+'LeftKey','加载中...');
  if($(prefix+'RightKey')) clearSelectOptions(prefix+'RightKey','加载中...');
  if($(prefix+'CommonPairs')) $(prefix+'CommonPairs').innerHTML=excelLoadingHtml('正在读取两份 Excel 列...');
  if($(prefix+'ManualPairs')) $(prefix+'ManualPairs').innerHTML='';
  hide($(prefix+'Result'));clearInlineError(prefix+'Error');
  if(!leftFile||!rightFile){
    $(prefix+'LoadStatus').textContent='';
    if($(prefix+'CommonPairs')) $(prefix+'CommonPairs').innerHTML='<span style="font-size:12px;color:#888">请先上传两份 BOM 文件</span>';
    return;
  }
  setLoadingStatus(prefix+'LoadStatus','正在读取列...');
  const fd=new FormData();
  fd.append('left_file',leftFile);fd.append('right_file',rightFile);fd.append('compare_type',compareType||'cadence_hq');
  fd.append('left_header_row',$(prefix+'LeftHdr').value||1);fd.append('right_header_row',$(prefix+'RightHdr').value||1);
  const ls=$(prefix+'LeftSheet').value, rs=$(prefix+'RightSheet').value;
  if(ls&&ls!=='先选择文件'&&ls!=='加载中...') fd.append('left_sheet',ls);
  if(rs&&rs!=='先选择文件'&&rs!=='加载中...') fd.append('right_sheet',rs);
  try{
    const d=await postFormJson(apiPrefix+'/generic_sheets',fd);
    if(!d.success) throw new Error(d.error||'读取列失败');
    $(prefix+'LeftSheet').innerHTML=(d.left_sheets||[]).map(s=>`<option${s===d.left_current_sheet?' selected':''}>${_escH(s)}</option>`).join('');
    $(prefix+'RightSheet').innerHTML=(d.right_sheets||[]).map(s=>`<option${s===d.right_current_sheet?' selected':''}>${_escH(s)}</option>`).join('');
    st.leftHeaders=d.left_headers||[];st.rightHeaders=d.right_headers||[];
    st.leftFormat=d.left_format||'';st.rightFormat=d.right_format||'';st.rightBomSheets=d.right_bom_sheets||[];
    if(d.left_header_row) $(prefix+'LeftHdr').value=d.left_header_row;
    if(d.right_header_row) $(prefix+'RightHdr').value=d.right_header_row;
    if($(prefix+'LeftKey')) $(prefix+'LeftKey').innerHTML=cmpOptions(st.leftHeaders,d.detected_left_key||d.left_detected_key);
    if($(prefix+'RightKey')) $(prefix+'RightKey').innerHTML=cmpOptions(st.rightHeaders,d.detected_right_key||d.right_detected_key);
    cmpPopulateCustomerStandardMapping(prefix);
    cmpRenderKeyPairs(prefix);
    cmpRenderPairs(prefix);
    const rightFmt=st.rightFormat==='plm_full'?'PLM 全量 BOM':'标准 HQ BOM';
    const rightSheetHint=st.rightFormat==='plm_full'?`，HQ 已识别 ${rightFmt}（${(st.rightBomSheets||[]).join('、')}）`:`，HQ 已识别 ${rightFmt}`;
    setPlainStatus(prefix+'LoadStatus',`已加载：左侧 ${st.leftHeaders.length} 列，右侧 ${st.rightHeaders.length} 列${rightSheetHint}`);
  }catch(e){
    if($(prefix+'CommonPairs')) $(prefix+'CommonPairs').innerHTML='<span style="font-size:12px;color:#888">列读取失败</span>';
    $(prefix+'LoadStatus').textContent='';showInlineError(prefix+'Error',e.message,prefix+'Status');
  }
}
function cmpCollectPairs(prefix){
  const pairs=[];
  document.querySelectorAll('#'+prefix+'CommonPairs input[type="checkbox"]:checked').forEach(x=>pairs.push({left:x.value,right:x.value}));
  (cmpState(prefix).manualPairs||[]).forEach(p=>{if(p.left&&p.right)pairs.push({left:p.left,right:p.right});});
  const seen=new Set();
  return pairs.filter(p=>{const k=p.left+'\u0000'+p.right;if(seen.has(k))return false;seen.add(k);return true;});
}
function initGenericBomCompare(prefix, compareType, apiPrefix='/api/bom_compare'){
  cmpState(prefix);
  $(prefix+'Refresh').onclick=()=>cmpRefresh(prefix,compareType,apiPrefix);
  $(prefix+'LeftFile').onchange=()=>cmpRefresh(prefix,compareType,apiPrefix);
  $(prefix+'RightFile').onchange=()=>cmpRefresh(prefix,compareType,apiPrefix);
  $(prefix+'LeftSheet').onchange=()=>cmpRefresh(prefix,compareType,apiPrefix);
  $(prefix+'RightSheet').onchange=()=>cmpRefresh(prefix,compareType,apiPrefix);
  $(prefix+'LeftHdr').onchange=()=>cmpRefresh(prefix,compareType,apiPrefix);
  $(prefix+'RightHdr').onchange=()=>cmpRefresh(prefix,compareType,apiPrefix);
  if($(prefix+'LeftKey')) $(prefix+'LeftKey').onchange=()=>{cmpRenderKeyPairs(prefix);cmpRenderPairs(prefix);};
  if($(prefix+'RightKey')) $(prefix+'RightKey').onchange=()=>{cmpRenderKeyPairs(prefix);cmpRenderPairs(prefix);};
  if($(prefix+'SelectAll')) $(prefix+'SelectAll').onclick=()=>document.querySelectorAll('#'+prefix+'CommonPairs input[type="checkbox"]').forEach(x=>x.checked=true);
  if($(prefix+'SelectNone')) $(prefix+'SelectNone').onclick=()=>document.querySelectorAll('#'+prefix+'CommonPairs input[type="checkbox"]').forEach(x=>x.checked=false);
  if($(prefix+'AddPair')) $(prefix+'AddPair').onclick=()=>{const st=cmpState(prefix);st.manualPairs.push({left:'',right:''});cmpRenderManualPairs(prefix);};
  if($(prefix+'AddKey')) $(prefix+'AddKey').onclick=()=>{const st=cmpState(prefix);cmpEnsureKeyPairs(prefix);st.keyPairs.push({left:'',right:'',transform:''});cmpRenderKeyPairs(prefix);cmpRenderPairs(prefix);};
  $(prefix+'Run').onclick=async function(){
    const leftFile=$(prefix+'LeftFile').files[0], rightFile=$(prefix+'RightFile').files[0];
    if(!leftFile||!rightFile){showInlineError(prefix+'Error','请上传两份 BOM 文件',prefix+'Status');return;}
    let fieldPairs=cmpCollectPairs(prefix);
    const keyCfg=cmpCollectKeyConfig(prefix);
    if(!keyCfg.left_key_cols.length||!keyCfg.right_key_cols.length){showInlineError(prefix+'Error','请先加载列并选择两侧匹配键',prefix+'Status');return;}
    if(!fieldPairs.length){showInlineError(prefix+'Error','请至少选择一组需要比对的字段',prefix+'Status');return;}
    const btn=$(prefix+'Run');btn.disabled=true;$(prefix+'Status').textContent='正在比对...';hide($(prefix+'Result'));clearInlineError(prefix+'Error');
    const cfg={compare_type:compareType,left_sheet:$(prefix+'LeftSheet').value,right_sheet:$(prefix+'RightSheet').value,
      left_header_row:parseInt($(prefix+'LeftHdr').value)||1,right_header_row:parseInt($(prefix+'RightHdr').value)||1,
      left_key_col:keyCfg.left_key_cols[0]||'',right_key_col:keyCfg.right_key_cols[0]||'',
      left_key_cols:keyCfg.left_key_cols,right_key_cols:keyCfg.right_key_cols,left_key_transforms:keyCfg.left_key_transforms,
      field_pairs:fieldPairs};
    const fd=new FormData();fd.append('left_file',leftFile);fd.append('right_file',rightFile);fd.append('config',JSON.stringify(cfg));
    try{
      const d=await postFormJson(apiPrefix+'/generic',fd);
      if(!d.success) throw new Error(d.error||'比对失败');
      const leftName='Cadence BOM';
      const rightName='HQ BOM';
      const expandHint=d.expanded_refdes?'<br>已按位号展开逐点比对':'';
      const diffText=false
        ? `制造商差异 <b style="color:#c07000">${d.manufacturer_diff||0}</b> | 型号差异 <b style="color:#c00000">${d.model_diff||0}</b> | 二供差异 <b style="color:#c07000">${d.second_source_diff||0}</b>`
        : `字段变更 <b style="color:#c07000">${d.changed}</b>`;
      const leftOnlyLabel=false?'仅客户BOM存在':`${leftName} 独有`;
      const rightOnlyLabel=false?'仅HQ BOM存在':`${rightName} 独有`;
      $(prefix+'Stats').innerHTML=`${leftOnlyLabel} <b style="color:#c00000">${d.left_only}</b> | ${rightOnlyLabel} <b style="color:#2a8a2a">${d.right_only}</b> | ${diffText} | 完全一致 <b>${d.same}</b><br>${leftName} 共 ${d.left_total} 行，${rightName} 共 ${d.right_total} 行${expandHint}`;
      $(prefix+'Dl').href=d.download;show($(prefix+'Result'));$(prefix+'Status').textContent='完成';
    }catch(e){showInlineError(prefix+'Error',e.message,prefix+'Status');}
    btn.disabled=false;
  };
}


function custOptionHtml(headers, selected='', optionalText='不映射'){
  let html=`<option value="">${optionalText}</option>`;
  (headers||[]).forEach(h=>{if(h) html+=`<option value="${_escH(h)}"${h===selected?' selected':''}>${_escH(h)}</option>`;});
  return html;
}
function custState(){window._custPreviewState=window._custPreviewState||{leftHeaders:[],rightHeaders:[]};return window._custPreviewState;}
async function custLoadColumns(apiPrefix='/api/bom_compare'){
  const leftFile=$('custLeftFile').files[0], rightFile=$('custRightFile').files[0];
  hide($('custPreviewPanel'));clearInlineError('custError');
  if(!leftFile||!rightFile){$('custLoadStatus').textContent='请先上传客户 BOM 和 HQ BOM';return;}
  setLoadingStatus('custLoadStatus','正在读取列...');
  const fd=new FormData();
  fd.append('left_file',leftFile);fd.append('right_file',rightFile);fd.append('compare_type','customer_preview');
  fd.append('left_header_row',$('custLeftHdr').value||1);
  const ls=$('custLeftSheet').value, rs=$('custRightSheet').value;
  if(ls&&ls!=='先选择文件'&&ls!=='加载中...') fd.append('left_sheet',ls);
  if(rs&&rs!=='先选择文件'&&rs!=='加载中...') fd.append('right_sheet',rs);
  try{
    const d=await postFormJson(apiPrefix+'/generic_sheets',fd);
    if(!d.success) throw new Error(d.error||'读取列失败');
    $('custLeftSheet').innerHTML=(d.left_sheets||[]).map(s=>`<option${s===d.left_current_sheet?' selected':''}>${_escH(s)}</option>`).join('');
    $('custRightSheet').innerHTML=(d.right_sheets||[]).map(s=>`<option${s===d.right_current_sheet?' selected':''}>${_escH(s)}</option>`).join('');
    if(d.left_header_row) $('custLeftHdr').value=d.left_header_row;
    const st=custState();st.leftHeaders=d.left_headers||[];st.rightHeaders=d.right_headers||[];
    $('custMapModel').innerHTML=custOptionHtml(st.leftHeaders,'','请选择客户列');
    $('custMapManufacturer').innerHTML=custOptionHtml(st.leftHeaders,'','请选择客户列');
    $('custMapRefdes').innerHTML=custOptionHtml(st.leftHeaders,'','不使用位号');
    $('custMapQuantity').innerHTML=custOptionHtml(st.leftHeaders,'','不映射');
    setPlainStatus('custLoadStatus',`已加载：客户 ${st.leftHeaders.length} 列，HQ ${st.rightHeaders.length} 列`);
  }catch(e){$('custLoadStatus').textContent='';showInlineError('custError',e.message,'custStatus');}
}
function custRowsHtml(rows, side){
  return (rows||[]).map(r=>{
    const cells=side==='customer'
      ? [r.row,r.match_key,r.refdes,r.model,r.manufacturer_raw,r.manufacturer_mapped,r.quantity,r.status,r.issue]
      : [r.row,r.match_key,r.refdes,r.model,r.manufacturer_raw,r.manufacturer_mapped,r.quantity,r.part_no,r.alternate,r.status,r.issue];
    return '<tr>'+cells.map(v=>`<td>${_escH(v==null?'':String(v))}</td>`).join('')+'</tr>';
  }).join('');
}
async function custRunPreview(apiPrefix='/api/bom_compare'){
  const leftFile=$('custLeftFile').files[0], rightFile=$('custRightFile').files[0];
  if(!leftFile||!rightFile){showInlineError('custError','请先上传客户 BOM 和 HQ BOM','custStatus');return;}
  const mapping={model:$('custMapModel').value,manufacturer:$('custMapManufacturer').value,refdes:$('custMapRefdes').value,quantity:$('custMapQuantity').value};
  if(!mapping.model||!mapping.manufacturer){showInlineError('custError','请先映射规格型号和制造商','custStatus');return;}
  if($('custMatchMode').value==='refdes'&&!mapping.refdes){showInlineError('custError','按位号匹配时必须映射位号列','custStatus');return;}
  const btn=$('custPreview');btn.disabled=true;setLoadingStatus('custStatus','正在标准化预览...');clearInlineError('custError');
  const cfg={left_sheet:$('custLeftSheet').value,right_sheet:$('custRightSheet').value,left_header_row:parseInt($('custLeftHdr').value)||1,mapping,match_mode:$('custMatchMode').value};
  const fd=new FormData();fd.append('left_file',leftFile);fd.append('right_file',rightFile);fd.append('config',JSON.stringify(cfg));
  try{
    const d=await postFormJson(apiPrefix+'/customer_hq_preview',fd);
    if(!d.success) throw new Error(d.error||'预览失败');
    $('custPreviewStats').innerHTML=`匹配模式：<b>${d.match_mode==='refdes'?'按位号匹配':'按型号+制造商匹配'}</b> | 客户 ${d.customer_total} 行，异常 ${d.customer_invalid} 行 | HQ ${d.hq_total} 行，异常 ${d.hq_invalid} 行`;
    $('custCustomerPreviewRows').innerHTML=custRowsHtml(d.customer_preview,'customer')||'<tr><td colspan="10">无数据</td></tr>';
    $('custHqPreviewRows').innerHTML=custRowsHtml(d.hq_preview,'hq')||'<tr><td colspan="11">无数据</td></tr>';
    show($('custPreviewPanel'));setPlainStatus('custStatus','预览完成');
  }catch(e){showInlineError('custError',e.message,'custStatus');}
  btn.disabled=false;
}
async function custRunExport(apiPrefix='/api/bom_compare'){
  const leftFile=$('custLeftFile').files[0], rightFile=$('custRightFile').files[0];
  if(!leftFile||!rightFile){showInlineError('custError','请先上传客户 BOM 和 HQ BOM','custStatus');return;}
  const mapping={model:$('custMapModel').value,manufacturer:$('custMapManufacturer').value,refdes:$('custMapRefdes').value,quantity:$('custMapQuantity').value};
  if(!mapping.model||!mapping.manufacturer){showInlineError('custError','请先映射规格型号和制造商','custStatus');return;}
  if($('custMatchMode').value==='refdes'&&!mapping.refdes){showInlineError('custError','按位号匹配时必须映射位号列','custStatus');return;}
  const btn=$('custExport');btn.disabled=true;setLoadingStatus('custStatus','正在生成详细比对报告...');clearInlineError('custError');hide($('custExportResult'));
  const cfg={left_sheet:$('custLeftSheet').value,right_sheet:$('custRightSheet').value,left_header_row:parseInt($('custLeftHdr').value)||1,mapping,match_mode:$('custMatchMode').value};
  const fd=new FormData();fd.append('left_file',leftFile);fd.append('right_file',rightFile);fd.append('config',JSON.stringify(cfg));
  try{
    const d=await postFormJson(apiPrefix+'/customer_hq_export',fd);
    if(!d.success) throw new Error(d.error||'导出失败');
    $('custExportStats').innerHTML=`匹配成功 <b>${d.matched}</b> | 字段差异 <b style="color:#c07000">${d.changed}</b> | 仅客户存在 <b style="color:#c00000">${d.customer_only}</b> | 仅HQ存在 <b style="color:#2a8a2a">${d.hq_only}</b> | 异常行 <b>${(d.customer_invalid||0)+(d.hq_invalid||0)}</b>`;
    $('custExportDl').href=d.download;
    show($('custExportResult'));setPlainStatus('custStatus','报告已生成');
  }catch(e){showInlineError('custError',e.message,'custStatus');}
  btn.disabled=false;
}
function initCustomerHqPreview(apiPrefix='/api/bom_compare'){
  $('custRefresh').onclick=()=>custLoadColumns(apiPrefix);
  $('custLeftFile').onchange=()=>custLoadColumns(apiPrefix);
  $('custRightFile').onchange=()=>custLoadColumns(apiPrefix);
  $('custLeftSheet').onchange=()=>custLoadColumns(apiPrefix);
  $('custRightSheet').onchange=()=>custLoadColumns(apiPrefix);
  $('custLeftHdr').onchange=()=>custLoadColumns(apiPrefix);
  $('custPreview').onclick=()=>custRunPreview(apiPrefix);
  $('custExport').onclick=()=>custRunExport(apiPrefix);
}

function initBomCompare(apiPrefix='/api/bom_compare'){
  document.querySelectorAll('.bomcmp-tab-btn').forEach(btn=>{
    btn.onclick=function(){
      document.querySelectorAll('.bomcmp-tab-btn').forEach(b=>{
        b.classList.remove('active');
        b.style.color='#888';
        b.style.borderBottomColor='transparent';
      });
      this.classList.add('active');
      this.style.color='#1a5ad4';
      this.style.borderBottomColor='#1a5ad4';
      document.querySelectorAll('#bomcmp-tab-customer-hq,#bomcmp-tab-hq-version,#bomcmp-tab-machine-hq-version,#bomcmp-tab-cadence-hq').forEach(el=>el.style.display='none');
      $('bomcmp-tab-'+this.dataset.bomcmpTab).style.display='block';
    };
  });
  initCustomerHqPreview(apiPrefix);
  initVersionCompare('hqv',{sheetsApi:apiPrefix+'/local_sheets',compareApi:apiPrefix+'/hq_version',label:'HQ BOM'});
  initVersionCompare('machv',{sheetsApi:apiPrefix+'/machine_local_sheets',compareApi:apiPrefix+'/machine_hq_version',label:'整机 HQ BOM'});
  initGenericBomCompare('cad','cadence_hq',apiPrefix);
}

function vcState(prefix){
  const key = '_vc_' + prefix;
  window[key] = window[key] || {oldHeaders:[], newHeaders:[], oldFormat:'', newFormat:'', oldBomSheets:[], newBomSheets:[]};
  return window[key];
}

function vcGetCheckedCols(prefix){
  return [...document.querySelectorAll('#'+prefix+'CompareCols input[type="checkbox"]:checked')].map(x=>x.value);
}

function vcRenderCompareCols(prefix, opts={}){
  const st=vcState(prefix);
  const common=st.oldHeaders.filter(h=>h && st.newHeaders.includes(h));
  const key=$(prefix+'KeyCol')?.value||'';
  if($(prefix+'RuleSummary')) $(prefix+'RuleSummary').textContent=`匹配规则：按「${key||'料号'}」匹配同一物料`;
  let h='';
  common.forEach(col=>{
    const checked=col===key?'':' checked';
    h+=`<label><input type="checkbox" value="${_escH(col)}"${checked}>${_escH(col)}</label>`;
  });
  $(prefix+'CompareCols').innerHTML=h || '<span style="font-size:12px;color:#888">请先加载两份 BOM 的列</span>';
}

async function vcLoadOne(prefix, which, sheetsApi){
  const fileEl=$(which==='old'?prefix+'OldFile':prefix+'NewFile');
  const sheetEl=$(which==='old'?prefix+'OldSheet':prefix+'NewSheet');
  const f=fileEl.files[0]; if(!f) return null;
  const fd=new FormData();
  fd.append('file',f);
  fd.append('header_row',$(prefix+'Hdr').value||1);
  if(sheetEl.value && sheetEl.value!=='先选择文件' && sheetEl.value!=='加载中...') fd.append('sheet_name',sheetEl.value);
  const r=await fetch(sheetsApi,{method:'POST',body:fd});
  const d=await r.json();
  if(!d.success) throw new Error(d.error||'读取文件失败');
  sheetEl.innerHTML=(d.sheets||[]).map(s=>`<option${s===d.current_sheet?' selected':''}>${_escH(s)}</option>`).join('');
  const st=vcState(prefix);
  if(which==='old'){st.oldHeaders=d.headers||[];st.oldFormat=d.format||'';st.oldBomSheets=d.bom_sheets||[];}
  else {st.newHeaders=d.headers||[];st.newFormat=d.format||'';st.newBomSheets=d.bom_sheets||[];}
  if(d.header_row) $(prefix+'Hdr').value=d.header_row;
  return d;
}

async function vcRefreshColumns(prefix, opts){
  clearInlineError(prefix+'Error');hide($(prefix+'Result'));
  const st=vcState(prefix); st.oldHeaders=[];st.newHeaders=[];st.oldFormat='';st.newFormat='';st.oldBomSheets=[];st.newBomSheets=[];
  clearSelectOptions(prefix+'KeyCol','加载中...');
  $(prefix+'CompareCols').innerHTML=excelLoadingHtml('正在读取两份 HQ BOM 列...');
  setLoadingStatus(prefix+'LoadStatus','读取列中...');
  try{
    const oldData=await vcLoadOne(prefix,'old',opts.sheetsApi);
    const newData=await vcLoadOne(prefix,'new',opts.sheetsApi);
    const common=st.oldHeaders.filter(h=>h && st.newHeaders.includes(h));
    const keySel=$(prefix+'KeyCol');
    keySel.innerHTML=common.map(h=>`<option value="${_escH(h)}">${_escH(h)}</option>`).join('');
    const detected=(oldData&&newData&&oldData.detected_key===newData.detected_key)?oldData.detected_key:(oldData?.detected_key||newData?.detected_key||'');
    if(detected && common.includes(detected)) keySel.value=detected;
    vcRenderCompareCols(prefix, opts);
    const fmtLabel=st.oldFormat==='plm_full'?'PLM 全量 BOM':'标准 HQ BOM';
    const sheetHint=st.oldFormat==='plm_full'?`；将比对 ${((st.oldBomSheets||[]).filter(s=>(st.newBomSheets||[]).includes(s))).join('、')}`:'';
    setPlainStatus(prefix+'LoadStatus',`已识别：${fmtLabel}，已加载共同列 ${common.length} 个${sheetHint}`);
  }catch(e){
    $(prefix+'CompareCols').innerHTML='<span style="font-size:12px;color:#888">列读取失败</span>';
    $(prefix+'LoadStatus').textContent='';
    showInlineError(prefix+'Error',e.message,prefix+'Status');
  }
}

function initVersionCompare(prefix, opts){
  vcState(prefix);
  $(prefix+'OldFile').onchange=()=>vcRefreshColumns(prefix,opts);
  $(prefix+'NewFile').onchange=()=>vcRefreshColumns(prefix,opts);
  $(prefix+'OldSheet').onchange=()=>vcRefreshColumns(prefix,opts);
  $(prefix+'NewSheet').onchange=()=>vcRefreshColumns(prefix,opts);
  $(prefix+'Refresh').onclick=()=>vcRefreshColumns(prefix,opts);
  $(prefix+'KeyCol').onchange=()=>vcRenderCompareCols(prefix,opts);
  $(prefix+'SelectAll').onclick=()=>document.querySelectorAll('#'+prefix+'CompareCols input[type="checkbox"]').forEach(x=>x.checked=true);
  $(prefix+'SelectNone').onclick=()=>document.querySelectorAll('#'+prefix+'CompareCols input[type="checkbox"]').forEach(x=>x.checked=false);
  $(prefix+'Run').onclick=async function(){
    const oldFile=$(prefix+'OldFile').files[0], newFile=$(prefix+'NewFile').files[0];
    if(!oldFile||!newFile){showInlineError(prefix+'Error',`请上传基准版本和对比版本 ${opts.label}` ,prefix+'Status');return;}
    const compareCols=vcGetCheckedCols(prefix);
    if(!$(prefix+'KeyCol').value){showInlineError(prefix+'Error','请选择匹配键列',prefix+'Status');return;}
    if(!compareCols.length){showInlineError(prefix+'Error','请至少选择一个比对字段',prefix+'Status');return;}
    const btn=$(prefix+'Run');btn.disabled=true;$(prefix+'Status').textContent='比对中...';
    hide($(prefix+'Result'));clearInlineError(prefix+'Error');
    const cfg={header_row:parseInt($(prefix+'Hdr').value)||1, old_sheet:$(prefix+'OldSheet').value,
      new_sheet:$(prefix+'NewSheet').value, key_col:$(prefix+'KeyCol').value, compare_cols:compareCols};
    const fd=new FormData();
    fd.append('old_file',oldFile);fd.append('new_file',newFile);fd.append('config',JSON.stringify(cfg));
    try{
      const d=await postFormJson(opts.compareApi,fd);
      if(!d.success) throw new Error(d.error||'比对失败');
      $(prefix+'Stats').innerHTML=`新增 <b style="color:#2a8a2a">${d.added}</b> | 删除 <b style="color:#c00000">${d.removed}</b> | 变更 <b style="color:#c07000">${d.changed}</b> | 未变更 <b>${d.unchanged}</b><br>基准版本 ${d.old_total} 项，对比版本 ${d.new_total} 项`;
      $(prefix+'Dl').href=d.download;
      show($(prefix+'Result'));$(prefix+'Status').textContent='完成！';
    }catch(e){showInlineError(prefix+'Error',e.message,prefix+'Status');}
    btn.disabled=false;
  };
}
// ─── BOM ──────────────────────────────────────────────
function updateBomPreview(headers, previewRows){
  if(!headers||!headers.length){$('bomPreview').innerHTML='';return;}
  let h='<table>';
  h+='<tr style="background:#e8f0fe">';
  headers.forEach((_,ci)=>h+=`<th style="color:#1a5ad4;font-size:11px;font-weight:700;padding:2px 6px">${String.fromCharCode(65+ci)}</th>`);
  h+='</tr><tr>';
  headers.forEach(hd=>h+=`<th>${_escH(hd||'')}</th>`);
  h+='</tr>';
  (previewRows||[]).forEach(r=>{
    h+='<tr>';
    r.forEach(c=>h+=`<td>${_escH(c||'')}</td>`);
    h+='</tr>';
  });
  h+='</table>';
  $('bomPreview').innerHTML=h;
}

function bomPrepareLoad(clearSheet){
  hide($('bomResult'));clearInlineError('bomError');
  $('bomPreview').innerHTML=excelLoadingHtml('\u6b63\u5728\u8bfb\u53d6 BOM \u9884\u89c8...');
  $('bomDetected').textContent='';
  if(clearSheet) setSelectPlaceholder('bomSheet','\u52a0\u8f7d\u4e2d...');
  setLoadingStatus('bomStatus','\u6b63\u5728\u8bfb\u53d6 Excel...');
}

function updateBomFromApi(d){
  if(!d.success) throw new Error(d.error||'\u8bfb\u53d6 Excel \u5931\u8d25');
  let opts='';
  (d.sheets||[]).forEach(s=>opts+=`<option${s===d.current_sheet?' selected':''}>${_escH(s)}</option>`);
  $('bomSheet').innerHTML=opts;
  if(d.detected){
    if(d.detected.name)  $('bomColName').value=d.detected.name;
    if(d.detected.qty)   $('bomColQty').value=d.detected.qty;
    if(d.detected.brand) $('bomColBrand').value=d.detected.brand;
    if(d.detected.model) $('bomColModel').value=d.detected.model;
  }
  if(d.fmt_guess){
    const sel=$('bomFmt');
    for(let i=0;i<sel.options.length;i++){
      if(sel.options[i].value===d.fmt_guess){sel.selectedIndex=i;break;}
    }
  }
  $('bomDetected').textContent='\u81ea\u52a8\u8bc6\u522b\uff1a'+JSON.stringify(d.detected||{});
  updateBomPreview(d.headers, d.preview);
  setPlainStatus('bomStatus',`\u5df2\u52a0\u8f7d ${(d.headers||[]).length} \u5217`);
}

async function bomLoadDetect(opts={}){
  const f=$('bomFile').files[0];
  if(!f){
    $('bomPreview').innerHTML='';$('bomDetected').textContent='';
    setSelectPlaceholder('bomSheet','\u5148\u9009\u62e9\u6587\u4ef6');
    setPlainStatus('bomStatus','');
    return;
  }
  bomPrepareLoad(!!opts.clearSheet);
  const fd=new FormData();
  fd.append('file',f);
  fd.append('header_row',$('bomHdr').value||1);
  if(!opts.clearSheet && $('bomSheet').value && $('bomSheet').value!=='\u5148\u9009\u62e9\u6587\u4ef6' && $('bomSheet').value!=='\u52a0\u8f7d\u4e2d...') fd.append('sheet_name',$('bomSheet').value);
  try{
    const r=await fetch('/api/bom/detect',{method:'POST',body:fd});
    const d=await r.json();
    updateBomFromApi(d);
  }catch(e){
    $('bomPreview').innerHTML='';
    showInlineError('bomError',e.message,'bomStatus');
  }
}

function initBom(){
  $('bomSheet').onchange=()=>bomLoadDetect({clearSheet:false});
  $('bomFile').onchange=()=>bomLoadDetect({clearSheet:true});
  $('bomRefresh').onclick=()=>bomLoadDetect({clearSheet:false});
  $('bomRun').onclick=async function(){
    const f=$('bomFile').files[0];if(!f){showInlineError('bomError','\u8bf7\u9009\u62e9\u6587\u4ef6','bomStatus');return;}
    const btn=$('bomRun');btn.disabled=true;$('bomStatus').textContent='\u5904\u7406\u4e2d...';
    hide($('bomResult'));clearInlineError('bomError');
    const fd=new FormData();fd.append('file',f);
    fd.append('fmt',$('bomFmt').value);
    fd.append('sheet',$('bomSheet').value);
    fd.append('header_row',$('bomHdr').value);
    fd.append('col_name',$('bomColName').value);
    fd.append('col_qty',$('bomColQty').value);
    fd.append('col_brand',$('bomColBrand').value);
    fd.append('col_model',$('bomColModel').value);
    fd.append('output_mode',$('bomOutputMode').value);
    try{
      const r=await fetch('/api/bom/convert',{method:'POST',body:fd});
      const d=await r.json();
      if(d.success){
        let stats=`\u5171 <b>${d.total}</b> \u884c`;
        if(d.skipped!=null) stats+=`\uff0c\u8df3\u8fc7\u7a7a\u884c ${d.skipped} \u884c`;
        $('bomStats').innerHTML=stats;
        $('bomDl').href=d.download;
        show($('bomResult'));$('bomStatus').textContent='\u5b8c\u6210\uff01';
      } else {
        $('bomError').textContent=d.error||'\u8f6c\u6362\u5931\u8d25';show($('bomError'));$('bomStatus').textContent='';
      }
    }catch(e){$('bomError').textContent=e.message;show($('bomError'));$('bomStatus').textContent='';}
    btn.disabled=false;
  };
}

function buildColMapTable(headers){
  if(!headers||!headers.length){$('plmColMap').innerHTML='';return;}
  let h='<table><tr style="background:#e8f0fe"><th style="width:50px">列</th><th>列名</th></tr>';
  headers.forEach(e=>{
    const m=e.match(/^([A-Z]+):(.+)/);
    if(m) h+=`<tr><td style="font-weight:700;color:#1a5ad4">${_escH(m[1])}</td><td>${_escH(m[2])}</td></tr>`;
    else  h+=`<tr><td></td><td>${_escH(e)}</td></tr>`;
  });
  h+='</table>';
  $('plmColMap').innerHTML=h;
}

function updatePlmFromApi(d){
  if(!d.success) return;
  window._plmPreviewHeaders = d.preview_headers || [];
  let opts='';
  d.sheets.forEach(s=>opts+=`<option${s===d.current_sheet?' selected':''}>${s}</option>`);
  $('plmSheet').innerHTML=opts;
  if(d.detected){
    if(d.detected.seq)         $('plmColSeq').value=d.detected.seq;
    if(d.detected.hq_pn)       $('plmColHqpn').value=d.detected.hq_pn;
    if(d.detected.supply_type) $('plmColStype').value=d.detected.supply_type;
    if(d.detected.qty && !window._plmQtyConfigs?.length) plmSetQtyConfigs([{qty_col:d.detected.qty}]);
  }
  $('plmDetectLog').textContent='自动识别：'+(d.headers||[]).join(' | ');
  buildColMapTable(d.headers);
  plmRefreshQtyConfigNames(false);
}

function plmPrepareLoad(clearSheet, clearConfigs){
  hide($('plmResult'));clearInlineError('plmError');
  window._plmPreviewHeaders=[];
  buildColMapTable([]);
  $('plmDetectLog').innerHTML=excelLoadingHtml('\u6b63\u5728\u8bfb\u53d6 PLM \u5217...');
  if(clearSheet) setSelectPlaceholder('plmSheet','\u52a0\u8f7d\u4e2d...');
  if(clearConfigs){window._plmQtyConfigs=[];plmRenderQtyConfigs();}
  setLoadingStatus('plmStatus','\u6b63\u5728\u8bfb\u53d6 Excel...');
}
async function plmLoadDetect(opts={}){
  const f=$('plmFile').files[0];
  if(!f){
    buildColMapTable([]);$('plmDetectLog').textContent='';window._plmPreviewHeaders=[];
    setSelectPlaceholder('plmSheet','\u5148\u9009\u62e9\u6587\u4ef6');setPlainStatus('plmStatus','');
    return;
  }
  plmPrepareLoad(!!opts.clearSheet, !!opts.clearConfigs);
  const fd=new FormData();fd.append('file',f);fd.append('header_row',$('plmHdr').value||4);
  if(!opts.clearSheet && $('plmSheet').value && $('plmSheet').value!=='\u5148\u9009\u62e9\u6587\u4ef6' && $('plmSheet').value!=='\u52a0\u8f7d\u4e2d...') fd.append('sheet_name',$('plmSheet').value);
  try{
    const r=await fetch('/api/plm/detect',{method:'POST',body:fd});
    const d=await r.json();
    if(!d.success) throw new Error(d.error||'\u8bfb\u53d6 Excel \u5931\u8d25');
    updatePlmFromApi(d);
    setPlainStatus('plmStatus',`\u5df2\u52a0\u8f7d ${(d.headers||[]).length} \u5217`);
  }catch(e){
    buildColMapTable([]);$('plmDetectLog').textContent='';showInlineError('plmError',e.message,'plmStatus');
  }
}

function plmColToIndex(col){
  const s=String(col||'').trim().toUpperCase();
  if(!s) return null;
  if(/^\d+$/.test(s)) return parseInt(s,10);
  let n=0;
  for(const ch of s){
    if(ch<'A'||ch>'Z') return null;
    n=n*26+(ch.charCodeAt(0)-64);
  }
  return n;
}

function plmIndexToCol(n){
  n=parseInt(n,10);
  let s='';
  while(n>0){
    const m=(n-1)%26;
    s=String.fromCharCode(65+m)+s;
    n=Math.floor((n-1)/26);
  }
  return s;
}

function plmGetCellText(row, col){
  const idx=plmColToIndex(col);
  if(!idx || !window._plmPreviewHeaders) return '';
  return window._plmPreviewHeaders[idx-1] || '';
}

function plmDefaultNameRow(){
  const hdr=parseInt($('plmHdr')?.value)||4;
  return Math.max(1, hdr-1);
}

function plmSetQtyConfigs(configs){
  window._plmQtyConfigs=(configs||[]).map(cfg=>({
    name_row: cfg.name_row || plmDefaultNameRow(),
    name: cfg.name || '',
    qty_col: cfg.qty_col || '',
  }));
  plmRenderQtyConfigs();
  plmRefreshQtyConfigNames(false);
}

function plmRenderQtyConfigs(){
  const body=$('plmQtyCfgBody');
  if(!body) return;
  const cfgs=window._plmQtyConfigs||[];
  body.innerHTML=cfgs.map((cfg,i)=>`
    <tr>
      <td><input type="number" min="1" value="${cfg.name_row||plmDefaultNameRow()}" style="width:80px" onchange="plmUpdateQtyCfg(${i},'name_row',this.value);plmRefreshOneQtyName(${i},true)"></td>
      <td><input type="text" value="${_escH(cfg.name||'')}" style="width:100%" placeholder="自动读取，可手动修改" onchange="plmUpdateQtyCfg(${i},'name',this.value)"></td>
      <td><input type="text" value="${_escH(cfg.qty_col||'')}" style="width:64px" placeholder="如 K" onchange="plmUpdateQtyCfg(${i},'qty_col',this.value);plmRefreshOneQtyName(${i},true)"></td>
      <td><button class="btn btn-sm btn-gray" type="button" onclick="plmRemoveQtyCfg(${i})">删除</button></td>
    </tr>`).join('');
}

function plmUpdateQtyCfg(i,key,val){
  window._plmQtyConfigs=window._plmQtyConfigs||[];
  if(!window._plmQtyConfigs[i]) return;
  window._plmQtyConfigs[i][key]=key==='qty_col'?String(val||'').trim().toUpperCase():val;
}

function plmRefreshOneQtyName(i,force){
  const cfg=(window._plmQtyConfigs||[])[i];
  if(!cfg || (!force && cfg.name)) return;
  const row=parseInt(cfg.name_row)||plmDefaultNameRow();
  const col=cfg.qty_col;
  const f=$('plmFile')?.files?.[0];
  if(!f || !col) {
    cfg.name = cfg.name || (col ? `用量${col}` : '');
    plmRenderQtyConfigs();
    return;
  }
  const fd=new FormData();
  fd.append('file',f);
  fd.append('sheet_name',$('plmSheet').value);
  fd.append('header_row',row);
  fetch('/api/plm/detect',{method:'POST',body:fd}).then(r=>r.json()).then(d=>{
    if(d.success){
      const idx=plmColToIndex(col);
      const name=(d.preview_headers||[])[idx-1] || '';
      cfg.name = name || `用量${col}`;
      plmRenderQtyConfigs();
    }
  });
}

function plmRefreshQtyConfigNames(force){
  (window._plmQtyConfigs||[]).forEach((_,i)=>plmRefreshOneQtyName(i,force));
}

function plmAddQtyCfg(){
  window._plmQtyConfigs=window._plmQtyConfigs||[];
  const last=window._plmQtyConfigs[window._plmQtyConfigs.length-1];
  const lastIdx=plmColToIndex(last?.qty_col);
  const nextCol=lastIdx?plmIndexToCol(lastIdx+1):'';
  window._plmQtyConfigs.push({name_row:plmDefaultNameRow(),name:'',qty_col:nextCol});
  plmRenderQtyConfigs();
  plmRefreshOneQtyName(window._plmQtyConfigs.length-1,true);
}

function plmRemoveQtyCfg(i){
  window._plmQtyConfigs=(window._plmQtyConfigs||[]).filter((_,idx)=>idx!==i);
  plmRenderQtyConfigs();
}

function plmSetQtyNameRow(row){
  window._plmQtyConfigs=window._plmQtyConfigs||[];
  window._plmQtyConfigs.forEach(cfg=>{cfg.name_row=row; cfg.name='';});
  plmRenderQtyConfigs();
  plmRefreshQtyConfigNames(true);
}

function plmCollectQtyConfigs(){
  return (window._plmQtyConfigs||[])
    .map(cfg=>({name_row:parseInt(cfg.name_row)||plmDefaultNameRow(), name:String(cfg.name||'').trim(), qty_col:String(cfg.qty_col||'').trim().toUpperCase()}))
    .filter(cfg=>cfg.qty_col);
}

function initPlm(){
  // ── PLM sub-tab switching ──
  document.querySelectorAll('.plm-tab-btn').forEach(btn=>{
    btn.onclick=function(){
      document.querySelectorAll('.plm-tab-btn').forEach(b=>{
        b.style.color='#888'; b.style.borderBottomColor='transparent';
      });
      this.style.color='#1a5ad4'; this.style.borderBottomColor='#1a5ad4';
      document.querySelectorAll('#plm-tab-bom-cfg,#plm-tab-spec-extract').forEach(el=>el.style.display='none');
      $('plm-tab-'+this.dataset.plmTab).style.display='block';
    };
  });

  // ── 规格型号提取 ──
  $('seFile').onchange = seLoadFile;
  $('seRefresh').onclick = ()=>{ if($('seFile').files[0]) seLoadFile(); };
  $('seRun').onclick = seRun;

  window._plmQtyConfigs=[];
  window._plmPreviewHeaders=[];
  plmRenderQtyConfigs();
  $('plmAddQtyCfg').onclick=plmAddQtyCfg;
  $('plmUsePrevRow').onclick=()=>plmSetQtyNameRow(plmDefaultNameRow());
  $('plmUseHeaderRow').onclick=()=>plmSetQtyNameRow(parseInt($('plmHdr').value)||4);
  $('plmFile').onchange=()=>plmLoadDetect({clearSheet:true, clearConfigs:true});
  $('plmSheet').onchange=()=>plmLoadDetect({clearSheet:false, clearConfigs:false});
  $('plmRefresh').onclick=()=>plmLoadDetect({clearSheet:false, clearConfigs:false});
  $('plmRun').onclick=async function(){
    const f=$('plmFile').files[0];if(!f){showInlineError('plmError','请选择文件','plmStatus');return;}
    const btn=$('plmRun');btn.disabled=true;$('plmStatus').textContent='处理中...';
    hide($('plmResult'));clearInlineError('plmError');hide($('plmLogBox'));
    const fd=new FormData();fd.append('file',f);
    fd.append('sheet',$('plmSheet').value);
    fd.append('header_row',$('plmHdr').value);
    fd.append('col_seq',$('plmColSeq').value);
    fd.append('col_hqpn',$('plmColHqpn').value);
    fd.append('col_stype',$('plmColStype').value);
    fd.append('qty_configs',JSON.stringify(plmCollectQtyConfigs()));
    try{
      const r=await fetch('/api/plm/convert',{method:'POST',body:fd});
      const d=await r.json();
      if(d.success){
        let stats=`写入 <b>${d.total}</b> 行${d.skipped?`，跳过 ${d.skipped} 行`:''}`;
        if(d.files&&d.files.length){
          stats += '<div style="margin-top:6px;font-size:12px;line-height:1.8">';
          stats += d.files.map(f=>`${f.project_name}（${f.qty_col}列）：${f.total} 行${f.skipped?`，跳过 ${f.skipped} 行`:''}`).join('<br>');
          stats += '</div>';
        }
        $('plmStats').innerHTML=stats;
        $('plmDl').href=d.download;
        $('plmDl').textContent=d.is_zip?'📦 下载 PLM 导入文件包':'📥 下载 PLM 导入文件';
        show($('plmResult'));$('plmStatus').textContent='完成！';
        if(d.skip_logs&&d.skip_logs.length){$('plmLog').textContent=d.skip_logs.join('\n');show($('plmLogBox'));}else hide($('plmLogBox'));
      } else {
        $('plmError').textContent=d.error||'转换失败';show($('plmError'));$('plmStatus').textContent='';
      }
    }catch(e){$('plmError').textContent=e.message;show($('plmError'));}
    btn.disabled=false;
  };
}

// ─── 规格型号提取 ─────────────────────────────────────


function initPlmAuto(){
  $('paRun').onclick=plmAutoRun;
  $('paAttRun').onclick=plmAutoAttachmentRun;
  $('paAttHqpn').addEventListener('input',()=>{
    const cleaned=($('paAttHqpn').value||'').replace(/\s+/g,'');
    if($('paAttHqpn').value!==cleaned) $('paAttHqpn').value=cleaned;
  });
  document.querySelectorAll('.plm-auto-tab-btn').forEach(btn=>{
    btn.onclick=()=>{
      const tab=btn.dataset.paTab;
      document.querySelectorAll('.plm-auto-tab-btn').forEach(b=>{
        const active=b.dataset.paTab===tab;
        b.style.color=active?'#1a5ad4':'#888';
        b.style.borderBottomColor=active?'#1a5ad4':'transparent';
        b.classList.toggle('active',active);
      });
      $('paTabSpec').style.display=tab==='spec'?'flex':'none';
      $('paTabAttach').style.display=tab==='attach'?'flex':'none';
    };
  });
}

async function plmAutoRun(){
  const f=$('paFile').files[0];
  const user=($('paUser').value||'').trim();
  const pass=$('paPass').value||'';
  if(!user){showInlineError('paError','请输入账号','paStatus');return;}
  if(!pass){showInlineError('paError','请输入密码','paStatus');return;}
  if(!f){showInlineError('paError','请选择需要上传的 Excel 文件','paStatus');return;}
  const btn=$('paRun');
  btn.disabled=true;
  $('paStatus').textContent='正在运行，浏览器会自动打开并操作 PLM...';
  hide($('paResult'));hide($('paLogBox'));clearInlineError('paError');
  const fd=new FormData();
  fd.append('username',user);
  fd.append('password',pass);
  fd.append('file',f);
  try{
    const r=await fetch('/api/plm/auto_spec_reverse',{method:'POST',body:fd});
    const d=await r.json();
    if(d.success){
      $('paDl').href=d.download;
      $('paStats').textContent=d.filename||'PLM 导出文件已生成';
      $('paStatus').textContent='完成！';
      show($('paResult'));
      if(d.log){$('paLog').textContent=d.log;show($('paLogBox'));}
    }else{
      $('paError').textContent=d.error||'自动化执行失败';
      show($('paError'));
      $('paStatus').textContent='';
      if(d.log){$('paLog').textContent=d.log;show($('paLogBox'));}
    }
  }catch(e){
    $('paError').textContent=e.message;
    show($('paError'));
    $('paStatus').textContent='';
  }
  btn.disabled=false;
}

function plmAttSetProgress(stage,pct,note){
  const panel=$('paAttProgressPanel');
  const bar=$('paAttProgressBar');
  const pctEl=$('paAttProgressPct');
  const stageEl=$('paAttProgressStage');
  const noteEl=$('paAttProgressNote');
  const value=Math.max(0,Math.min(100,parseInt(pct)||0));
  if(panel) show(panel);
  if(bar) bar.style.width=value+'%';
  if(pctEl) pctEl.textContent=value+'%';
  if(stageEl) stageEl.textContent=stage||'\u5904\u7406\u4e2d';
  if(noteEl) noteEl.textContent=note||'PLM \u81ea\u52a8\u5316\u6b63\u5728\u6267\u884c\uff0c\u8bf7\u52ff\u5173\u95ed\u672c\u9875\u9762';
}

async function plmAutoAttachmentRun(){
  const user=($('paAttUser').value||'').trim();
  const pass=$('paAttPass').value||'';
  const hqpn=($('paAttHqpn').value||'').replace(/\s+/g,'');
  if(!user){showInlineError('paAttError','\u8bf7\u8f93\u5165\u8d26\u53f7','paAttStatus');return;}
  if(!pass){showInlineError('paAttError','\u8bf7\u8f93\u5165\u5bc6\u7801','paAttStatus');return;}
  if(!hqpn){showInlineError('paAttError','\u8bf7\u8f93\u5165 HQ \u6599\u53f7','paAttStatus');return;}
  const btn=$('paAttRun');
  btn.disabled=true;
  $('paAttStatus').textContent='\u6b63\u5728\u521b\u5efa\u4e0b\u8f7d\u4efb\u52a1...';
  hide($('paAttResult'));hide($('paAttLogBox'));clearInlineError('paAttError');
  plmAttSetProgress('\u51c6\u5907\u542f\u52a8',3,'\u6b63\u5728\u63d0\u4ea4\u4efb\u52a1');
  const fd=new FormData();
  fd.append('username',user);
  fd.append('password',pass);
  fd.append('hqpn',hqpn);
  let pollTimer=null;
  try{
    const r=await fetch('/api/plm/auto_hq_attachments',{method:'POST',body:fd});
    const d=await r.json();
    if(!d.success) throw new Error(d.error||'\u81ea\u52a8\u5316\u4efb\u52a1\u521b\u5efa\u5931\u8d25');
    const statusUrl=d.status_url||('/api/plm/auto_hq_attachments/status/'+d.job_id);
    $('paAttStatus').textContent='\u4efb\u52a1\u5df2\u542f\u52a8\uff0c\u6b63\u5728\u6267\u884c PLM \u81ea\u52a8\u5316...';
    const poll=async()=>{
      const sr=await fetch(statusUrl);
      const s=await sr.json();
      if(!s.success) throw new Error(s.error||'\u8bfb\u53d6\u4efb\u52a1\u72b6\u6001\u5931\u8d25');
      plmAttSetProgress(s.stage||'\u5904\u7406\u4e2d',s.progress||0,s.status==='done'?'\u5904\u7406\u5b8c\u6210':'\u6b63\u5728\u6267\u884c\uff1a'+(s.hqpn||hqpn));
      if(s.log){$('paAttLog').textContent=s.log;show($('paAttLogBox'));}
      if(s.status==='done'){
        if(pollTimer) clearInterval(pollTimer);
        $('paAttDl').href=s.download;
        $('paAttStats').textContent=s.filename||'PLM \u9644\u4ef6\u5df2\u4e0b\u8f7d';
        $('paAttStatus').textContent='\u5b8c\u6210\uff01';
        plmAttSetProgress('\u4e0b\u8f7d\u5b8c\u6210',100,'\u53ef\u4ee5\u70b9\u51fb\u4e0b\u65b9\u94fe\u63a5\u4e0b\u8f7d\u9644\u4ef6');
        show($('paAttResult'));
        btn.disabled=false;
      }else if(s.status==='error'){
        if(pollTimer) clearInterval(pollTimer);
        $('paAttError').textContent=s.error||'\u81ea\u52a8\u5316\u6267\u884c\u5931\u8d25';
        show($('paAttError'));
        $('paAttStatus').textContent='';
        plmAttSetProgress('\u6267\u884c\u5931\u8d25',100,'\u8bf7\u67e5\u770b\u6267\u884c\u65e5\u5fd7');
        btn.disabled=false;
      }
    };
    await poll();
    pollTimer=setInterval(()=>{poll().catch(e=>{
      if(pollTimer) clearInterval(pollTimer);
      $('paAttError').textContent=e.message;
      show($('paAttError'));
      $('paAttStatus').textContent='';
      btn.disabled=false;
    });},1000);
  }catch(e){
    if(pollTimer) clearInterval(pollTimer);
    $('paAttError').textContent=e.message;
    show($('paAttError'));
    $('paAttStatus').textContent='';
    plmAttSetProgress('\u6267\u884c\u5931\u8d25',100,'\u4efb\u52a1\u672a\u80fd\u542f\u52a8\u6216\u72b6\u6001\u8bfb\u53d6\u5931\u8d25');
    btn.disabled=false;
  }
}
async function seLoadFile(){
  const f=$('seFile').files[0];
  clearInlineError('seError');hide($('seResult'));
  clearSelectOptions('seCol','\u52a0\u8f7d\u4e2d...');
  clearSelectOptions('seExcludeCol','\u52a0\u8f7d\u4e2d...');
  if(!f){setSelectPlaceholder('seSheet','\u5148\u9009\u62e9\u6587\u4ef6');setPlainStatus('seStatus','');return;}
  setLoadingStatus('seStatus','\u6b63\u5728\u8bfb\u53d6 Excel \u5217...');
  const fd=new FormData();
  fd.append('file',f);
  fd.append('header_row',$('seHdr').value||1);
  const cur=$('seSheet').value;
  if(cur&&cur!=='\u5148\u9009\u62e9\u6587\u4ef6'&&cur!=='\u52a0\u8f7d\u4e2d...') fd.append('sheet_name',cur);
  try{
    const r=await fetch('/api/feishu/local_sheets',{method:'POST',body:fd});
    const d=await r.json();
    if(!d.success) throw new Error(d.error||'\u52a0\u8f7d\u5931\u8d25');
    const ss=$('seSheet');
    ss.innerHTML=d.sheets.map(s=>`<option${s===d.current_sheet?' selected':''}>${_escH(s)}</option>`).join('');
    ss.onchange=seLoadFile;
    const kc=$('seCol');
    kc.innerHTML=d.headers.map(h=>`<option value="${_escH(h)}">${_escH(h)}</option>`).join('');
    const auto=d.headers.find(h=>h.includes('\u89c4\u683c\u578b\u53f7'));
    if(auto) kc.value=auto;
    const ex=$('seExcludeCol');
    ex.innerHTML='<option value="">\u4e0d\u542f\u7528\u5254\u9664\u89c4\u5219</option>'+d.headers.map(h=>`<option value="${_escH(h)}">${_escH(h)}</option>`).join('');
    const autoEx=d.headers.find(h=>h.includes('HQ\u6599\u53f7')||h==='HQ\u6599\u53f7'||h.includes('\u6599\u53f7'));
    if(autoEx) ex.value=autoEx;
    setPlainStatus('seStatus',`\u5df2\u52a0\u8f7d ${d.headers.length} \u5217`);
  }catch(e){
    clearSelectOptions('seCol','\u52a0\u8f7d\u5931\u8d25');clearSelectOptions('seExcludeCol','\u52a0\u8f7d\u5931\u8d25');
    setPlainStatus('seStatus','\u52a0\u8f7d\u5931\u8d25\uff1a'+e.message);
  }
}

async function seRun(){
  const f=$('seFile').files[0];
  if(!f){showInlineError('seError','请先上传文件','seStatus');return;}
  const col=$('seCol').value;
  if(!col||col==='先刷新列表'){showInlineError('seError','请先刷新列表并选择提取列','seStatus');return;}
  const excludeCol=$('seExcludeCol').value;
  $('seRun').disabled=true; $('seStatus').style.color='#1a8a1a'; $('seStatus').textContent='提取中...';
  $('seResult').style.display='none'; clearInlineError('seError');
  const fd=new FormData();
  fd.append('file',f);
  fd.append('config',JSON.stringify({
    header_row:parseInt($('seHdr').value)||1,
    sheet_name:$('seSheet').value,
    col_name:col,
    exclude_col_name:excludeCol,
  }));
  try{
    const r=await fetch('/api/plm/spec_extract',{method:'POST',body:fd});
    const d=await r.json();
    if(d.success){
      $('seStats').textContent=`共提取 ${d.count} 条规格型号${d.skipped_excluded?`，按剔除列跳过 ${d.skipped_excluded} 行`:''}`;
      $('seDl').href=d.download;
      $('seStatus').textContent='';
      $('seResult').style.display='block';
    }else{
      $('seError').textContent='错误：'+d.error;
      $('seError').style.display='block';
      $('seStatus').textContent='';
    }
  }catch(e){$('seError').textContent=e.message;$('seError').style.display='block';$('seStatus').textContent='';}
  $('seRun').disabled=false;
}

// ─── 飞书匹配 ─────────────────────────────────────────
const FS_TABLES = (window.BOM_TOOLS_BOOTSTRAP && window.BOM_TOOLS_BOOTSTRAP.presetTables) || [];
const FS_DEFAULT_CONFIG = (window.BOM_TOOLS_BOOTSTRAP && window.BOM_TOOLS_BOOTSTRAP.defaultConfig) || {};

// ── SheetConfig 默认值工厂 ──────────────────────────────────
function _mkSheetCfg(){
  const fkInit=_fsGlobalLocalKeys.map(()=>'');
  return {enabled:false, local_key_names:_fsGlobalLocalKeys.slice(), feishu_key_names:fkInit,
          fetch_col_names:[], fetch_col_aliases:{}, _headers:[], _expanded:false,
          cache_key:'', cache_row_count:0, cache_fetched_at:0,
          row_count_at_cache:0, _cache_stale:false};
}

let fsTables = FS_TABLES.map((t,i)=>({
  ...t, idx:i, _sheets:[], _connected:false,
  sheet_configs:{},
}));
let fsCurIdx = null;
let _fsGlobalFetchMap = [  // [{std_name, default_alias}]  pre-populated defaults
  {std_name:'HQ料号',    default_alias:'HQ料号'},
  {std_name:'HQ规格型号', default_alias:'HQ规格型号'},
  {std_name:'HQ制造商',  default_alias:'HQ制造商'},
  {std_name:'优选等级',  default_alias:'优选等级'},
  {std_name:'HQ描述',    default_alias:'HQ描述'},
];
let _fsGlobalLocalKeys = [''];  // dynamic local-side match keys
let _fsGlobalLocalKeyTransforms = [''];  // '' | 'manufacturer_alias'

function fsRenderGLKSection(){
  const cont=$('fsGLkContainer'); if(!cont) return;
  const hdrs=window._fsLocalHeaders||[];
  while(_fsGlobalLocalKeyTransforms.length<_fsGlobalLocalKeys.length) _fsGlobalLocalKeyTransforms.push('');
  if(_fsGlobalLocalKeyTransforms.length>_fsGlobalLocalKeys.length) _fsGlobalLocalKeyTransforms.splice(_fsGlobalLocalKeys.length);
  let h='';
  _fsGlobalLocalKeys.forEach((val,i)=>{
    const req=i===0;
    let opts=`<option value="">${req?'— 选择列 —':'（不使用）'}</option>`;
    hdrs.forEach(h2=>{if(h2) opts+=`<option value="${_escH(h2)}"${h2===val?' selected':''}>${_escH(h2)}</option>`;});
    const mfgChecked=_fsGlobalLocalKeyTransforms[i]==='manufacturer_alias'?' checked':'';
    const showMfgMap=/厂商|厂家|制造商|供应商|brand|manufacturer/i.test(val||'');
    h+=`<div style="display:flex;align-items:center;gap:4px">
      <span style="min-width:16px;font-size:12px;${req?'color:#c00000':'color:#888'}">${i+1}</span>
      <select id="fsGlk_${i}" style="width:180px;font-size:12px" onchange="fsGLKChange()">${opts}</select>
      <label title="勾选后，本地该列会先通过厂商命名映射表转换为 HQ 规范厂商名，再参与匹配" style="font-size:12px;color:#555;display:${showMfgMap?'flex':'none'};align-items:center;gap:2px;white-space:nowrap">
        <input type="checkbox" id="fsGlkMfg_${i}" onchange="fsGLKChange()"${mfgChecked}>厂商映射
      </label>
      ${_fsGlobalLocalKeys.length>1?`<button class="btn btn-sm btn-gray" style="padding:0 6px;font-size:12px;line-height:20px" onclick="fsRemoveGLK(${i})">×</button>`:''}
    </div>`;
  });
  h+=`<button class="btn btn-sm btn-gray" style="padding:0 8px;font-size:12px;line-height:20px;margin-top:2px" onclick="fsAddGLK()">+</button>`;
  cont.innerHTML=h;
}

function fsGLKChange(){
  _fsGlobalLocalKeys=_fsGlobalLocalKeys.map((_,i)=>($(`fsGlk_${i}`)||{value:''}).value.trim());
  _fsGlobalLocalKeyTransforms=_fsGlobalLocalKeys.map((k,i)=>
    (/厂商|厂家|制造商|供应商|brand|manufacturer/i.test(k||'') && ($(`fsGlkMfg_${i}`)||{checked:false}).checked)
      ? 'manufacturer_alias' : ''
  );
  fsSaveConfig();
  if(fsCurIdx!==null) fsSelectTable(fsCurIdx);
}

function fsAddGLK(){
  _fsGlobalLocalKeys=_fsGlobalLocalKeys.map((_,i)=>($(`fsGlk_${i}`)||{value:''}).value.trim());
  _fsGlobalLocalKeyTransforms=_fsGlobalLocalKeys.map((k,i)=>
    (/厂商|厂家|制造商|供应商|brand|manufacturer/i.test(k||'') && ($(`fsGlkMfg_${i}`)||{checked:false}).checked)
      ? 'manufacturer_alias' : ''
  );
  _fsGlobalLocalKeys.push('');
  _fsGlobalLocalKeyTransforms.push('');
  fsTables.forEach(t=>Object.values(t.sheet_configs).forEach(sc=>sc.feishu_key_names.push('')));
  fsRenderGLKSection(); fsSaveConfig();
  if(fsCurIdx!==null) fsSelectTable(fsCurIdx);
}

function fsRemoveGLK(i){
  _fsGlobalLocalKeys=_fsGlobalLocalKeys.map((_,j)=>($(`fsGlk_${j}`)||{value:''}).value.trim());
  _fsGlobalLocalKeyTransforms=_fsGlobalLocalKeys.map((k,j)=>
    (/厂商|厂家|制造商|供应商|brand|manufacturer/i.test(k||'') && ($(`fsGlkMfg_${j}`)||{checked:false}).checked)
      ? 'manufacturer_alias' : ''
  );
  _fsGlobalLocalKeys.splice(i,1);
  _fsGlobalLocalKeyTransforms.splice(i,1);
  if(!_fsGlobalLocalKeys.length) _fsGlobalLocalKeys=[''];
  if(!_fsGlobalLocalKeyTransforms.length) _fsGlobalLocalKeyTransforms=[''];
  fsTables.forEach(t=>Object.values(t.sheet_configs).forEach(sc=>{
    if(sc.feishu_key_names.length>i) sc.feishu_key_names.splice(i,1);
    if(!sc.feishu_key_names.length) sc.feishu_key_names=[''];
  }));
  fsRenderGLKSection(); fsSaveConfig();
  if(fsCurIdx!==null) fsSelectTable(fsCurIdx);
}

// ─── localStorage 持久化 ──────────────────────────────
const _FS_KEY = 'bom-tools-feishu-v2';  // v2: per-sheet config

function fsSaveConfig(){
  const bu=$('fsBaseUrl'), or_=$('fsOrigin'), ui=$('fsUserId');
  const prev = window._fsSavedCfg || {};
  const cfg = {
    base_url: bu ? bu.value : prev.base_url||'',
    origin:  or_ ? or_.value : prev.origin||'',
    user_id:  ui ? ui.value  : prev.user_id||'',
    global_local_keys: _fsGlobalLocalKeys.slice(),
    global_local_key_transforms: _fsGlobalLocalKeyTransforms.slice(),
    global_fetch_map: _fsGlobalFetchMap.map(r=>({std_name:r.std_name||'',default_alias:r.default_alias||''})),
    tables: fsTables.map(t=>({
      idx: t.idx, token: t.token||'',
      _sheets: t._sheets||[], _connected: !!t._connected,
      sheet_configs: Object.fromEntries(
        Object.entries(t.sheet_configs).map(([sid, sc])=>[sid,{
          enabled: !!sc.enabled,
          local_key_names: sc.local_key_names||[''],
          feishu_key_names: sc.feishu_key_names||[''],
          fetch_col_names: sc.fetch_col_names||[],
          fetch_col_aliases: sc.fetch_col_aliases||{},
          _headers: sc._headers||[],
          cache_key: sc.cache_key||'',
          cache_row_count: sc.cache_row_count||0,
          cache_fetched_at: sc.cache_fetched_at||0,
          row_count_at_cache: sc.row_count_at_cache||0,
        }])
      ),
    })),
  };
  try{ localStorage.setItem(_FS_KEY, JSON.stringify(cfg)); window._fsSavedCfg=cfg; }catch(e){}
}

function fsExportConfig(){
  // 按最新架构导出：全局本地键 + 全局提取列 + 每个表格有效 sheet 配置
  const out = {
    global_local_keys: _fsGlobalLocalKeys.filter(k=>k),
    global_local_key_transforms: _fsGlobalLocalKeyTransforms.slice(0, _fsGlobalLocalKeys.length),
    global_fetch_map: _fsGlobalFetchMap.filter(r=>r.std_name).map(r=>r.std_name),
    tables: fsTables.map(t=>{
      // 找出有意义的 sheet 配置：enabled 或填了至少一个飞书键或有提取列覆盖
      const sheetList = (t._sheets||[]).map(s=>{
        const sid=s.sheetId;
        const sc=t.sheet_configs[sid];
        if(!sc) return null;
        const fk=(sc.feishu_key_names||[]).filter(k=>k);
        const aliases=sc.fetch_col_aliases||{};
        if(!sc.enabled && !fk.length && !Object.keys(aliases).length) return null;
        return {
          sheet_id: sid,
          sheet_name: s.title||sid,
          enabled: !!sc.enabled,
          feishu_key_names: fk,
          fetch_col_aliases: aliases,
        };
      }).filter(Boolean);
      if(!sheetList.length) return null;
      return {name:t.name, token:t.token||'', sheets:sheetList};
    }).filter(Boolean),
  };

  const json = JSON.stringify(out, null, 2);
  const blob = new Blob([json], {type:'application/json'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'bom_config.json';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function fsClearConfig(){
  if(!confirm('确认清除所有飞书本地配置？（含 Token、列映射、全局列映射、缓存记录）')) return;
  try{ localStorage.removeItem(_FS_KEY); }catch(e){}
  window._fsSavedCfg = null;
  _fsGlobalFetchMap = [];
  _fsGlobalLocalKeys = [''];
  _fsGlobalLocalKeyTransforms = [''];
  fsTables = FS_TABLES.map((t,i)=>({...t, idx:i, _sheets:[], _connected:false, sheet_configs:{}}));
  renderFsTrees();
  fsRenderGFMSection();
  fsRenderGLKSection();
  $('fsCfgStatus').textContent='配置已清除';
}

function fsRestoreDefault(){
  if(!FS_DEFAULT_CONFIG || !FS_DEFAULT_CONFIG.tables) { setStatus('fsCfgStatus','没有内置默认配置','#c00000'); return; }
  if(!confirm('将用内置默认配置覆盖当前所有设置（本地键、提取列、所有 Sheet 映射），确定？')) return;
  // 恢复全局本地键
  _fsGlobalLocalKeys = (FS_DEFAULT_CONFIG.global_local_keys||[]).length
    ? FS_DEFAULT_CONFIG.global_local_keys.slice() : [''];
  _fsGlobalLocalKeyTransforms = (FS_DEFAULT_CONFIG.global_local_key_transforms||[]).slice(0, _fsGlobalLocalKeys.length);
  while(_fsGlobalLocalKeyTransforms.length<_fsGlobalLocalKeys.length) _fsGlobalLocalKeyTransforms.push('');
  // 恢复全局提取列
  _fsGlobalFetchMap = (FS_DEFAULT_CONFIG.global_fetch_map||[]).map(n=>({std_name:n,default_alias:n}));
  if(!_fsGlobalFetchMap.length) _fsGlobalFetchMap=[{std_name:'HQ料号',default_alias:'HQ料号'},{std_name:'HQ规格型号',default_alias:'HQ规格型号'},{std_name:'HQ制造商',default_alias:'HQ制造商'},{std_name:'优选等级',default_alias:'优选等级'},{std_name:'HQ描述',default_alias:'HQ描述'}];
  // 重建 fsTables，注入默认 sheet configs
  const defByToken = {};
  (FS_DEFAULT_CONFIG.tables||[]).forEach(dt=>{ defByToken[dt.token] = dt; });
  fsTables = FS_TABLES.map((t,i)=>{
    const dt = defByToken[t.token]||{};
    const sheet_configs = {};
    (dt.sheets||[]).forEach(ds=>{
      sheet_configs[ds.sheet_id] = {
        enabled: !!ds.enabled,
        feishu_key_names: ds.feishu_key_names||[],
        fetch_col_aliases: ds.fetch_col_aliases||{},
        fetch_col_names: [],
        _headers: [], _expanded: false,
        cache_key:'', cache_row_count:0, cache_fetched_at:0,
        row_count_at_cache:0, _cache_stale:false,
      };
    });
    return {...t, idx:i, _sheets:[], _connected:false, sheet_configs};
  });
  fsCurIdx = null;
  renderFsTrees();
  fsRenderGFMSection();
  fsRenderGLKSection();
  fsSaveConfig();
  $('fsCfgStatus').textContent='默认配置已恢复';
}

// 页面加载时预读配置并合并到 fsTables
(function(){
  try{
    const raw = localStorage.getItem(_FS_KEY);
    if(!raw) return;
    const cfg = JSON.parse(raw);
    window._fsSavedCfg = cfg;
    // Only restore if saved map has at least one non-empty entry
    if(cfg.global_fetch_map && cfg.global_fetch_map.some(r=>r.std_name))
      _fsGlobalFetchMap = cfg.global_fetch_map.filter(r=>r.std_name);
    if(cfg.tables){
      cfg.tables.forEach(st=>{
        const t = fsTables[st.idx]; if(!t) return;
        t.token      = st.token || t.token || '';
        t._sheets    = st._sheets || [];
        t._connected = !!st._connected;
        const saved_sc = st.sheet_configs || {};
        Object.entries(saved_sc).forEach(([sid, sc])=>{
          t.sheet_configs[sid] = Object.assign(_mkSheetCfg(), sc, {_expanded:false, _cache_stale:false});
        });
      });
    }
  }catch(e){}
}());

// ─── 全局输出列映射管理 ────────────────────────────────
function fsRenderGFMSection(){
  const el=$('fsGFMSection'); if(!el) return;
  let h='';
  _fsGlobalFetchMap.forEach((r,i)=>{
    h+=`<div style="display:flex;align-items:center;gap:4px">
      <span style="min-width:16px;font-size:12px;color:#888">${i+1}</span>
      <input type="text" id="gfm_std_${i}" value="${_escH(r.std_name)}"
        style="width:180px;font-size:12px;padding:1px 5px;border:1px solid #aaa;border-radius:3px;height:24px"
        placeholder="列名" oninput="fsUpdateGFMRow(${i})">
      ${_fsGlobalFetchMap.length>1?`<button class="btn btn-sm btn-gray" style="padding:0 6px;font-size:12px;line-height:20px" onclick="fsRemoveGFMRow(${i})">×</button>`:''}
    </div>`;
  });
  h+=`<button class="btn btn-sm btn-gray" style="padding:0 8px;font-size:12px;line-height:20px;margin-top:2px" onclick="fsAddGFMRow()">+</button>`;
  el.innerHTML=h;
}

function fsAddGFMRow(){
  _fsGlobalFetchMap.push({std_name:'',default_alias:''});
  fsRenderGFMSection();  fsSaveConfig();
}

function fsRemoveGFMRow(i){
  _fsGlobalFetchMap.splice(i,1);
  fsRenderGFMSection();  fsSaveConfig();
  if(fsCurIdx!==null) fsSelectTable(fsCurIdx);
}

function fsUpdateGFMRow(i){
  const s=$(`gfm_std_${i}`);
  if(s){
    _fsGlobalFetchMap[i].std_name=s.value;
    _fsGlobalFetchMap[i].default_alias=s.value;
  }
  fsSaveConfig();
  if(fsCurIdx!==null) fsSelectTable(fsCurIdx);
}


function initFeishu(){
  const saved = window._fsSavedCfg;
  if(saved){
    if(saved.base_url) $('fsBaseUrl').value = saved.base_url;
    if(saved.origin)   $('fsOrigin').value  = saved.origin;
    if(saved.user_id)  $('fsUserId').value  = saved.user_id;
    if(saved.global_local_keys && saved.global_local_keys.length){
      _fsGlobalLocalKeys=saved.global_local_keys.slice();
      if(!_fsGlobalLocalKeys.length) _fsGlobalLocalKeys=[''];
    }
    if(saved.global_local_key_transforms && saved.global_local_key_transforms.length){
      _fsGlobalLocalKeyTransforms=saved.global_local_key_transforms.slice(0,_fsGlobalLocalKeys.length);
    }else{
      _fsGlobalLocalKeyTransforms=_fsGlobalLocalKeys.map(()=>'');
    }
    while(_fsGlobalLocalKeyTransforms.length<_fsGlobalLocalKeys.length) _fsGlobalLocalKeyTransforms.push('');
    if(_fsGlobalLocalKeys.length===1 && !_fsGlobalLocalKeys.includes('厂商')){
      _fsGlobalLocalKeys.push('厂商');
      _fsGlobalLocalKeyTransforms.push('manufacturer_alias');
      fsTables.forEach(t=>Object.values(t.sheet_configs).forEach(sc=>{
        const makerAlias=(sc.fetch_col_aliases||{})['HQ制造商']||'制造商';
        if((sc.feishu_key_names||[]).length===1) sc.feishu_key_names.push(makerAlias);
      }));
      fsSaveConfig();
      $('fsCfgStatus').textContent = '已从旧配置自动补齐厂商匹配键';
    }
    if(!$('fsCfgStatus').textContent) $('fsCfgStatus').textContent = '已从本地存储恢复配置';
  }
  ['fsBaseUrl','fsOrigin','fsUserId'].forEach(id=>{
    const el=$(id); if(el) el.oninput = fsSaveConfig;
  });
  fsRenderGFMSection();
  fsRenderGLKSection();
  // 上传文件后更新本地键 selects
  const _updateLkList=()=>{ fsRenderGLKSection(); };
  window._updateLkList=_updateLkList;
  renderFsTrees();
  $('fsBatchUpdate').onclick=fsBatchUpdate;
  $('fsExportCfg').onclick=fsExportConfig;
  $('fsRestoreDefault').onclick=fsRestoreDefault;
  $('fsClearConfig').onclick=fsClearConfig;
  $('fsRun').onclick=fsRunMatch;
  $('fsFile').onchange=async function(){
    const f=this.files[0];
    window._fsLocalHeaders=[];
    setSelectPlaceholder('fsSheet','\u52a0\u8f7d\u4e2d...');
    fsRenderGLKSection();
    hide($('fsResult'));clearInlineError('fsError');
    if(!f){setSelectPlaceholder('fsSheet','\u5148\u9009\u62e9\u6587\u4ef6');setPlainStatus('fsRunStatus2','');return;}
    setLoadingStatus('fsRunStatus2','\u6b63\u5728\u8bfb\u53d6\u672c\u5730 BOM \u5217...');
    const fd=new FormData();fd.append('file',f);fd.append('header_row',$('fsHdr').value||1);
    try{
      const r=await fetch('/api/feishu/local_sheets',{method:'POST',body:fd});
      const d=await r.json();
      if(!d.success) throw new Error(d.error||'\u8bfb\u53d6\u6587\u4ef6\u5931\u8d25');
      window._fsLocalHeaders=d.headers||[];
      const sel=$('fsSheet');sel.innerHTML='';
      (d.sheets||[]).forEach(s=>{const o=document.createElement('option');o.value=s;o.textContent=s;sel.appendChild(o);});
      if(window._updateLkList) window._updateLkList();
      setPlainStatus('fsRunStatus2',`\u5df2\u52a0\u8f7d ${(d.headers||[]).length} \u5217`);
    }catch(e){
      setSelectPlaceholder('fsSheet','\u52a0\u8f7d\u5931\u8d25');
      showInlineError('fsError',e.message,'fsRunStatus2');
    }
  };
}

function _fmtTime(ts){
  if(!ts) return '';
  const d=new Date(ts*1000);
  return `${d.getMonth()+1}/${d.getDate()} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
}

function _sheetSummary(t){
  // returns {enabledCount, cachedCount, staleCount, totalEnabled}
  let enabled=0, cached=0, stale=0;
  Object.values(t.sheet_configs).forEach(sc=>{
    if(sc.enabled){ enabled++;
      if(sc.cache_key){ cached++; if(sc._cache_stale) stale++; }
    }
  });
  return {enabled, cached, stale};
}

function renderFsTrees(){
  const cats={'优选库':[],'对应关系库':[]};
  fsTables.forEach(t=>{const c=t.category||'优选库';(cats[c]=cats[c]||[]).push(t);});
  let h='';
  for(const [cat,ts] of Object.entries(cats)){
    if(!ts.length)continue;
    h+=`<div class="tree-cat${cat==='对应关系库'?' alt':''}">▸ ${cat}</div>`;
    ts.forEach(t=>{
      const {enabled, cached, stale} = _sheetSummary(t);
      // ld badge: ✓=全部启用sheet已缓存  ⚠=有过期缓存  ◎=已连接  —=未连接
      const ld = cached>0 ? (stale>0?'⚠':'✓') : (t._connected?'◎':'—');
      const ldCls = cached>0 ? (stale>0?'badge-orange':'badge-green') : (t._connected?'badge-blue':'badge-gray');
      // mt badge: ●=有启用且有缓存  ！=有启用无缓存  ○=无启用
      const mt = enabled>0&&cached>0?'●':enabled>0?'！':'○';
      const mtCls = enabled>0&&cached>0?'badge-green':enabled>0?'badge-orange':'badge-gray';
      const hint = enabled>0 ? ` <span style="font-size:10px;color:#888">${enabled}Sheet</span>` : '';
      h+=`<div class="tree-item" data-idx="${t.idx}" onclick="fsSelectTable(${t.idx})">
        <span class="indent"></span><span style="flex:1">${t.name}${hint}</span>
        <span class="badge-sm ${ldCls}">${ld}</span>
        <span class="badge-sm ${mtCls}">${mt}</span>
      </div>`;
    });
  }
  $('fsTreeList').innerHTML=h;
}

function _escH(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/"/g,'&quot;'); }

function fsSelectTable(idx){
  fsCurIdx=idx;
  const t=fsTables[idx];
  document.querySelectorAll('#fsTreeList .tree-item').forEach(el=>el.classList.toggle('selected',parseInt(el.dataset.idx)===idx));

  let h=`<h4>${_escH(t.name)} <span style="font-size:11px;color:#888;font-weight:400">(${_escH(t.category||'优选库')})</span></h4>`;
  h+=`<div class="row"><label>Token：</label><input type="text" id="fsDetToken" value="${_escH(t.token||'')}" style="width:300px" placeholder="飞书表格 Token"></div>`;
  h+=`<div class="row">
    <button class="btn btn-sm btn-primary" onclick="fsConnectTable(${idx})">连接并获取 Sheet 列表</button>
  </div>`;

  if(t._sheets.length){
    h+=`<div style="margin-top:10px;margin-bottom:4px;font-weight:600">Sheets 配置：</div>`;
    t._sheets.forEach(s=>{
      const sid=s.sheetId;
      const sc = t.sheet_configs[sid] || _mkSheetCfg();
      const headers = sc._headers||[];
      const chk = sc.enabled?'checked':'';
      let cacheTag;
      if(sc.cache_key){
        cacheTag = sc._cache_stale
          ? `<span style="color:#c07000;font-size:11px">⚠过期(${_fmtTime(sc.cache_fetched_at)})</span>`
          : `<span style="color:#2a8a2a;font-size:11px">✓${sc.cache_row_count}行(${_fmtTime(sc.cache_fetched_at)})</span>`;
      } else {
        cacheTag = `<span style="color:#e07000;font-size:11px">无缓存</span>`;
      }
      const hdrTag = headers.length ? `<span style="color:#888;font-size:11px">${headers.length}列</span>` : '';
      const exp = sc._expanded;
      h+=`<div style="border:1px solid #dde;border-radius:4px;margin-bottom:6px;background:${sc.enabled?'#f8fff8':'#fafafa'}">`;
      // Sheet header row
      h+=`<div style="display:flex;align-items:center;gap:8px;padding:6px 10px;cursor:pointer" onclick="fsToggleSheetExpand(${idx},'${sid}')">
        <input type="checkbox" ${chk} onclick="event.stopPropagation();fsToggleSheetEnable(${idx},'${sid}',this.checked)" title="参与匹配">
        <span style="flex:1;font-size:13px">${_escH(s.title)}</span>
        ${hdrTag} ${cacheTag}
        <button class="btn btn-sm btn-green" style="padding:1px 7px;font-size:12px" onclick="event.stopPropagation();fsCacheSheet(${idx},'${sid}')">⬇ 缓存</button>
        ${sc.cache_key?`<button class="btn btn-sm btn-gray" style="padding:1px 7px;font-size:12px;color:#c00000" onclick="event.stopPropagation();fsClearSheetCache(${idx},'${sid}')">✕ 清除</button>`:''}
        <span style="font-size:12px;color:#666">${exp?'▲':'▼'}</span>
      </div>`;
      // Expanded config section
      if(exp){
        // local key datalist
        const lkDl = `fslk_${idx}_${sid}`;
        const fkDl = `fsfk_${idx}_${sid}`;
        let lkOpts='';(window._fsLocalHeaders||[]).forEach(h2=>lkOpts+=`<option value="${_escH(h2)}">`);
        let fkOpts='';headers.forEach(h2=>fkOpts+=`<option value="${_escH(h2)}">`);
        h+=`<datalist id="${lkDl}">${lkOpts}</datalist><datalist id="${fkDl}">${fkOpts}</datalist>`;
        h+=`<div style="padding:8px 12px;border-top:1px solid #eee;display:flex;gap:24px;align-items:flex-start">`;
        // ── 左列：匹配键 ──
        h+=`<div style="flex:0 0 auto;min-width:260px">`;
        h+=`<div style="font-size:12px;font-weight:600;margin-bottom:6px;color:#333">匹配键（飞书侧）</div>`;
        // feishu key count mirrors local key count — persist any resize immediately
        const glkCount=_fsGlobalLocalKeys.length;
        let _fkResized=false;
        while(sc.feishu_key_names.length<glkCount){sc.feishu_key_names.push('');_fkResized=true;}
        if(sc.feishu_key_names.length>glkCount){sc.feishu_key_names.splice(glkCount);_fkResized=true;}
        if(_fkResized) fsSaveConfig();
        sc.feishu_key_names.forEach((fv, ki)=>{
          const lkName=_fsGlobalLocalKeys[ki]||`键${ki+1}`;
          const req=ki===0?'color:#c00000':'color:#555';
          let opts=`<option value="">${ki===0?'— 选择列 —':'（不使用）'}</option>`;
          headers.forEach(h2=>{if(h2) opts+=`<option value="${_escH(h2)}"${h2===fv?' selected':''}>${_escH(h2)}</option>`;});
          h+=`<div style="display:flex;align-items:center;gap:6px;margin-bottom:4px">
            <label style="min-width:65px;font-size:12px;${req};text-align:right">${_escH(lkName)}：</label>
            <select id="fsfk_${idx}_${sid}_${ki}" style="width:190px;font-size:12px" onchange="fsSaveSheetCfg(${idx},'${sid}')">${opts}</select>
          </div>`;
        });
        h+=`</div>`;
        // ── 右列：提取列映射 ──
        const validGFM = _fsGlobalFetchMap.filter(r=>r.std_name);
        h+=`<div style="flex:1">`;
        if(validGFM.length){
          h+=`<div style="font-size:12px;font-weight:600;margin-bottom:6px;color:#333">提取列映射</div>`;
          validGFM.forEach((gm, gi)=>{
            const overrideVal=(sc.fetch_col_aliases||{})[gm.std_name]||'';
            const defAlias=gm.default_alias||gm.std_name;
            let opts=`<option value="">— 待映射 —</option>`;
            headers.forEach(h2=>{if(h2) opts+=`<option value="${_escH(h2)}"${h2===overrideVal?' selected':''}>${_escH(h2)}</option>`;});
            h+=`<div style="display:flex;align-items:center;gap:6px;margin-bottom:4px">
              <label style="min-width:75px;font-size:12px;color:#444;text-align:left">${_escH(gm.std_name)}：</label>
              <select id="fsgfa_${idx}_${sid}_${gi}" style="width:190px;font-size:12px" onchange="fsSaveSheetCfg(${idx},'${sid}')">${opts}</select>
            </div>`;
          });
        } else {
          h+=`<div style="color:#aaa;font-size:12px">在上方"提取列"中添加标准列后显示</div>`;
        }
        h+=`</div>`;
        h+=`</div>`;  // end flex row
      }
      h+=`</div>`;
    });
  } else if(t._connected){
    h+=`<div style="color:#888;font-size:12px;margin-top:10px">该在线表格无 Sheet 或 Token 失效</div>`;
  } else {
    h+=`<div style="color:#888;font-size:12px;margin-top:10px">填写 Token 后点击「连接」获取 Sheet 列表</div>`;
  }

  $('fsDetail').innerHTML=h;
}

function fsToggleSheetEnable(tIdx, sid, val){
  const t=fsTables[tIdx];
  if(!t.sheet_configs[sid]) t.sheet_configs[sid]=_mkSheetCfg();
  t.sheet_configs[sid].enabled=val;
  renderFsTrees();fsSaveConfig();
  // re-render without collapsing current expand state
  fsSelectTable(tIdx);
}

function fsToggleSheetExpand(tIdx, sid){
  const t=fsTables[tIdx];
  if(!t.sheet_configs[sid]) t.sheet_configs[sid]=_mkSheetCfg();
  t.sheet_configs[sid]._expanded = !t.sheet_configs[sid]._expanded;
  fsSelectTable(tIdx);
}


function fsSaveSheetCfg(tIdx, sid){
  const t=fsTables[tIdx];
  if(!t.sheet_configs[sid]) t.sheet_configs[sid]=_mkSheetCfg();
  const sc=t.sheet_configs[sid];
  const _fkCount=sc.feishu_key_names.length||1;
  sc.feishu_key_names=Array.from({length:_fkCount},(_,ki)=>($(`fsfk_${tIdx}_${sid}_${ki}`)||{value:''}).value.trim());
  // Save per-sheet alias overrides for global fetch map
  sc.fetch_col_aliases={};
  _fsGlobalFetchMap.filter(r=>r.std_name).forEach((gm,gi)=>{
    const el=$(`fsgfa_${tIdx}_${sid}_${gi}`);
    if(el && el.value.trim()) sc.fetch_col_aliases[gm.std_name]=el.value.trim();
  });
  fsSaveConfig();
}


async function fsConnectTable(idx){
  const t=fsTables[idx];
  const tokenEl=$('fsDetToken');
  if(tokenEl)t.token=tokenEl.value.trim();
  if(!t.token){showInlineError('fsError','请先填写 Token','fsRunStatus2');return;}
  clearInlineError('fsError');
  const btns=$('fsDetail').querySelectorAll('button');btns.forEach(b=>b.disabled=true);
  // 加载提示
  let hint=$('fsConnectHint_'+idx);
  if(!hint){
    const btn=$('fsDetail').querySelector('button.btn-primary');
    if(btn){hint=document.createElement('span');hint.id='fsConnectHint_'+idx;hint.className='hint';btn.after(hint);}
  }
  if(hint) hint.textContent='⏳ 连接中，正在获取 Sheet 列表及表头...';
  try{
    const r=await fetch('/api/feishu/sheets',{method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({base_url:$('fsBaseUrl').value,origin:$('fsOrigin').value,user_id:$('fsUserId').value,token:t.token})});
    const d=await r.json();
    if(d.success){
      t._sheets=d.sheets;t._connected=true;
      const sheetHeaders=d.sheet_headers||{};
      // 更新每个 sheet 的表头，并检测缓存是否过期
      d.sheets.forEach(s=>{
        const sid=s.sheetId;
        if(!t.sheet_configs[sid]) t.sheet_configs[sid]=_mkSheetCfg();
        const sc=t.sheet_configs[sid];
        const hdrs=(sheetHeaders[sid]||[]).filter(h=>h);
        if(hdrs.length) sc._headers=hdrs;
        // 检测缓存是否过期（行数比对）
        if(sc.cache_key && sc.row_count_at_cache>0){
          sc._cache_stale = (s.rowCount||0) !== sc.row_count_at_cache;
        }
      });
      fsSelectTable(idx);renderFsTrees();fsSaveConfig();
    } else { if(hint)hint.textContent=''; showInlineError('fsError',d.error||'连接失败','fsRunStatus2'); }
  }catch(e){ if(hint)hint.textContent=''; showInlineError('fsError',e.message,'fsRunStatus2'); }
  btns.forEach(b=>b.disabled=false);
}

// 缓存单个 Sheet 数据到服务端
async function fsCacheSheet(tIdx, sid){
  const t=fsTables[tIdx];
  if(!t.token){showInlineError('fsError','请先填写并连接 Token','fsRunStatus2');return;}
  clearInlineError('fsError');
  const btns=$('fsDetail').querySelectorAll('button');btns.forEach(b=>b.disabled=true);
  // 找到该 sheet 的缓存按钮，在旁边插入提示
  let cHint=document.getElementById(`fsCacheHint_${tIdx}_${sid}`);
  if(!cHint){
    const allBtns=[...$('fsDetail').querySelectorAll('button')];
    const cBtn=allBtns.find(b=>b.getAttribute('onclick')||'');  // 先占位，fsSelectTable 会重绘
    cHint=document.createElement('span');cHint.id=`fsCacheHint_${tIdx}_${sid}`;cHint.className='hint';
    $('fsDetail').appendChild(cHint);
  }
  cHint.textContent='⏳ 正在拉取全部数据，请稍候...';
  try{
    const r=await fetch('/api/feishu/load',{method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({base_url:$('fsBaseUrl').value,origin:$('fsOrigin').value,user_id:$('fsUserId').value,
        token:t.token, sheet_id:sid})});
    const d=await r.json();
    if(d.success){
      if(!t.sheet_configs[sid]) t.sheet_configs[sid]=_mkSheetCfg();
      const sc=t.sheet_configs[sid];
      sc.cache_key          = d.cache_key;
      sc.cache_row_count    = d.row_count;
      sc.cache_fetched_at   = d.fetched_at;
      sc.row_count_at_cache = d.row_count_at_cache||0;
      sc._cache_stale       = false;
      if(d.headers&&d.headers.length) sc._headers=d.headers.filter(h=>h);
      fsSelectTable(tIdx);renderFsTrees();fsSaveConfig();
    } else { showInlineError('fsError','缓存失败：'+(d.error||'未知错误'),'fsRunStatus2'); }
  }catch(e){showInlineError('fsError',e.message,'fsRunStatus2');}
  btns.forEach(b=>b.disabled=false);
}

function fsClearSheetCache(tIdx, sid){
  const sc=fsTables[tIdx].sheet_configs[sid];
  if(!sc) return;
  sc.cache_key=''; sc.cache_row_count=0; sc.cache_fetched_at=0;
  sc.row_count_at_cache=0; sc._cache_stale=false;
  fsSaveConfig();
  fsSelectTable(tIdx); renderFsTrees();
}

// 一键缓存所有启用的 Sheet
async function fsBatchUpdate(){
  if(!confirm('即将开始缓存所有启用的 Sheet。\n\n缓存过程中请勿离开或刷新当前页面，否则缓存任务会停止，需要重新开始。\n\n确认开始缓存？')) return;
  $('fsBatchUpdate').disabled=true;
  const _bh=$('fsBatchHint'); _bh.style.color='#1a8a1a'; _bh.textContent='缓存中...';
  let count=0;
  try{
    for(let i=0;i<fsTables.length;i++){
      const t=fsTables[i];
      if(!t.token) continue;
      // 缓存所有已启用的 Sheet（含从未缓存过的）
      const enabledSids=Object.entries(t.sheet_configs).filter(([,sc])=>sc.enabled).map(([sid])=>sid);
      if(!enabledSids.length) continue;
      if(!t._connected){
        _bh.textContent=`连接中：${t.name}`;
        // silently connect to get sheets+headers
        try{
          const r=await fetch('/api/feishu/sheets',{method:'POST',headers:{'Content-Type':'application/json'},
            body:JSON.stringify({base_url:$('fsBaseUrl').value,origin:$('fsOrigin').value,user_id:$('fsUserId').value,token:t.token})});
          const d=await r.json();
          if(d.success){ t._sheets=d.sheets;t._connected=true;
            const sh=d.sheet_headers||{};
            d.sheets.forEach(s=>{if(!t.sheet_configs[s.sheetId])t.sheet_configs[s.sheetId]=_mkSheetCfg();
              const hdrs=(sh[s.sheetId]||[]).filter(h=>h);if(hdrs.length)t.sheet_configs[s.sheetId]._headers=hdrs;});
          }
        }catch(e){console.error('连接失败:',t.name,e);}
      }
      if(!t._sheets) continue; // 连接失败则跳过此库
      for(const sid of enabledSids){
        const sheetTitle=(t._sheets||[]).find(s=>s.sheetId===sid)?.title||sid;
        _bh.textContent=`缓存 ${t.name}/${sheetTitle}`;
        try{
          const r=await fetch('/api/feishu/load',{method:'POST',headers:{'Content-Type':'application/json'},
            body:JSON.stringify({base_url:$('fsBaseUrl').value,origin:$('fsOrigin').value,user_id:$('fsUserId').value,token:t.token,sheet_id:sid})});
          const d=await r.json();
          if(d.success){
            if(!t.sheet_configs[sid])t.sheet_configs[sid]=_mkSheetCfg();
            const sc=t.sheet_configs[sid];
            sc.cache_key=d.cache_key;sc.cache_row_count=d.row_count;
            sc.cache_fetched_at=d.fetched_at;sc.row_count_at_cache=d.row_count_at_cache||0;
            sc._cache_stale=false;if(d.headers&&d.headers.length)sc._headers=d.headers.filter(h=>h);
            count++;
          }
        }catch(e){console.error('缓存失败:',t.name,sid,e);}
      }
      renderFsTrees(); // 每完成一个在线库立即更新左侧状态
    }
  }finally{
    $('fsBatchUpdate').disabled=false;
  }
  _bh.style.color='#1a8a1a'; _bh.textContent=`✓ 完成 ${count} 个 Sheet`;
  renderFsTrees();fsSaveConfig();
  if(fsCurIdx!==null)fsSelectTable(fsCurIdx);
}

async function fsRunMatch(){
  const f=$('fsFile').files[0];if(!f){showInlineError('fsError','请先上传 BOM 文件','fsRunStatus2');return;}

  // 全局本地键
  const gLocalKeyPairs=_fsGlobalLocalKeys.map((k,i)=>({
    name:(k||'').trim(),
    transform:_fsGlobalLocalKeyTransforms[i]||'',
  })).filter(k=>k.name);
  const gLocalKeys=gLocalKeyPairs.map(k=>k.name);
  const gLocalKeyTransforms=gLocalKeyPairs.map(k=>k.transform);
  if(!gLocalKeys.length){showInlineError('fsError','请先填写本地匹配键（本地 BOM 文件区域）','fsRunStatus2');return;}

  // 构建 per-sheet 配置
  const tables=[];
  fsTables.forEach(t=>{
    if(!t.token) return;
    const enabledSheets=[];
    (t._sheets||[]).forEach(s=>{
      const sc=t.sheet_configs[s.sheetId];
      if(!sc||!sc.enabled) return;
      const fk=(sc.feishu_key_names||[]).filter(k=>k);
      if(!fk.length) return;
      // 本地键与飞书键取相同数量
      const lk=gLocalKeys.slice(0,fk.length);
      const lkt=gLocalKeyTransforms.slice(0,fk.length);
      // 构建全局列映射（含per-sheet覆盖别名）
      const gfm=_fsGlobalFetchMap.filter(r=>r.std_name);
      const overrides=sc.fetch_col_aliases||{};
      const fetch_col_map=gfm.map(r=>({
        output:r.std_name,
        alias:(overrides[r.std_name]||'').trim()||(r.default_alias||r.std_name),
      }));
      enabledSheets.push({
        sheet_id:s.sheetId, sheet_name:s.title,
        enabled:true,
        local_key_names:lk, feishu_key_names:fk,
        local_key_transforms:lkt,
        fetch_col_names:sc.fetch_col_names||[],
        fetch_col_map: fetch_col_map,
        cache_key:sc.cache_key||'',
      });
    });
    if(enabledSheets.length) tables.push({name:t.name, token:t.token, sheets:enabledSheets});
  });

  if(!tables.length){showInlineError('fsError','没有启用的 Sheet，请先在表格库中勾选并配置列映射','fsRunStatus2');return;}

  const runBtn=$('fsRun');
  const oldRunText=runBtn.innerHTML;
  runBtn.disabled=true;$('fsBatchUpdate').disabled=true;
  runBtn.innerHTML='<span class="spinner">&#8635;</span> 匹配中';
  $('fsRunStatus2').textContent='正在处理，完成后会自动生成结果文件';
  show($('fsMatchWait'));
  hide($('fsResult'));clearInlineError('fsError');hide($('fsLogBox'));

  const config={
    base_url:$('fsBaseUrl').value, origin:$('fsOrigin').value, user_id:$('fsUserId').value,
    sheet_name:$('fsSheet').value, header_row:parseInt($('fsHdr').value)||1,
    tables,
  };
  const fd=new FormData();fd.append('file',f);fd.append('config',JSON.stringify(config));
  try{
    const r=await fetch('/api/feishu/match',{method:'POST',body:fd});
    const d=await r.json();
    if(d.logs&&d.logs.length){$('fsLog').textContent=d.logs.join('\n');show($('fsLogBox'));}
    if(d.success){
      $('fsStats').innerHTML=`共 <b>${d.total}</b> 行 | 命中 <b>${d.matched}</b> | 未匹配 <b>${d.unmatched}</b>`;
      $('fsDl').href=d.download;show($('fsResult'));$('fsRunStatus2').textContent='完成！';
    } else {
      $('fsError').textContent=d.error||'匹配失败';show($('fsError'));$('fsRunStatus2').textContent='';
    }
  }catch(e){$('fsError').textContent=e.message;show($('fsError'));$('fsRunStatus2').textContent='';}
  hide($('fsMatchWait'));
  runBtn.innerHTML=oldRunText;
  runBtn.disabled=false;$('fsBatchUpdate').disabled=false;
}

// ═══════════════════ 查询BOM优选率 ═══════════════════
(function(){
  // 工具初始化（每次切换到此工具时调用）
  function prInit(){
    $('prFile').onchange = prLoadFile;
    $('prRefresh').onclick = ()=>{ if($('prFile').files[0]) prLoadFile(); };
    $('prRun').onclick = prRun;
    hide($('prResult')); hide($('prError'));
  }

  async function prLoadFile(){
    const f = $('prFile').files[0];
    clearInlineError('prError');hide($('prResult'));
    clearSelectOptions('prKeyCol','\u52a0\u8f7d\u4e2d...');
    if(!f){setSelectPlaceholder('prSheet','\u5148\u9009\u62e9\u6587\u4ef6');setPlainStatus('prStatus','');return;}
    setLoadingStatus('prStatus','\u6b63\u5728\u8bfb\u53d6 Excel \u5217...');
    const hdr = parseInt($('prHdr').value)||1;
    const fd = new FormData();
    fd.append('file', f);
    fd.append('header_row', hdr);
    const cur = $('prSheet').value;
    if(cur && cur !== '\u5148\u9009\u62e9\u6587\u4ef6' && cur !== '\u52a0\u8f7d\u4e2d...') fd.append('sheet_name', cur);
    try{
      const r = await fetch('/api/feishu/local_sheets', {method:'POST', body:fd});
      const d = await r.json();
      if(!d.success) throw new Error(d.error||'\u52a0\u8f7d\u5931\u8d25');
      const ss = $('prSheet');
      ss.innerHTML = d.sheets.map(s=>`<option${s===d.current_sheet?' selected':''}>${_escH(s)}</option>`).join('');
      ss.onchange = prLoadFile;
      const kc = $('prKeyCol');
      kc.innerHTML = d.headers.map(h=>`<option value="${_escH(h)}">${_escH(h)}</option>`).join('');
      const auto = d.headers.find(h=>h.includes('HQ\u6599\u53f7')||h==='HQ\u6599\u53f7');
      if(auto) kc.value = auto;
      setPlainStatus('prStatus',`\u5df2\u52a0\u8f7d ${d.headers.length} \u5217`);
    }catch(e){
      clearSelectOptions('prKeyCol','\u52a0\u8f7d\u5931\u8d25');
      setPlainStatus('prStatus','\u52a0\u8f7d\u5931\u8d25\uff1a'+e.message);
    }
  }

  async function prRun(){
    const f=$('prFile').files[0];
    if(!f){showInlineError('prError','请先上传 BOM 文件','prStatus');return;}
    const keyCol=$('prKeyCol').value;
    if(!keyCol||keyCol==='先刷新列表'){showInlineError('prError','请先刷新列表并选择 HQ料号 列','prStatus');return;}

    // Build tables config: 优选库 only, enabled sheets with cache
    const tables=[];
    fsTables.forEach(t=>{
      // filter by category from FS_TABLES (preset)
      const preset=FS_TABLES.find(p=>p.token===t.token);
      if(!preset||preset.category!=='优选库') return;
      const sheets=[];
      Object.entries(t.sheet_configs).forEach(([sid,sc])=>{
        if(!sc.enabled||!sc.cache_key) return;
        sheets.push({sid, name:(t._sheets||[]).find(s=>s.sheetId===sid)?.title||sid,
                     cache_key:sc.cache_key, fetch_col_aliases:sc.fetch_col_aliases||{}});
      });
      if(sheets.length) tables.push({name:t.name, token:t.token, sheets});
    });

    if(!tables.length){
      showInlineError('prError','没有找到已缓存的优选库 sheet。请先在「飞书优选库+关系库匹配」中缓存相关 sheet。','prStatus');
      return;
    }

    $('prRun').disabled=true; $('prStatus').style.color='#1a8a1a';
    $('prStatus').textContent='查询中...';
    hide($('prResult')); clearInlineError('prError');

    const cfg={
      header_row: parseInt($('prHdr').value)||1,
      sheet_name: $('prSheet').value,
      local_key_col: keyCol,
      tables,
    };
    const fd=new FormData();
    fd.append('file',f);
    fd.append('config',JSON.stringify(cfg));
    try{
      const r=await fetch('/api/feishu/pref_rate',{method:'POST',body:fd});
      const d=await r.json();
      if(d.success){
        $('prStats').innerHTML=
          `共 <b>${d.total}</b> 行 &nbsp;|&nbsp; `+
          `匹配到HQ料号 <b style="color:#2a8a2a">${d.matched}</b> 行 &nbsp;|&nbsp; `+
          `未匹配 <b style="color:#888">${d.unmatched}</b> 行（不参与计算）<br>`+
          `已匹配中：优选料 <b style="color:#2a8a2a">${d.preferred}</b> 行 &nbsp;|&nbsp; `+
          `非优选料 <b style="color:#c07000">${d.non_preferred}</b> 行<br>`+
          `<span style="font-size:13px;color:#555">优选率（优选料/已匹配）= </span>`+
          `<b style="font-size:20px;color:#1a5ad4">${d.rate}</b>`;
        $('prDl').href=d.download;
        $('prStatus').textContent='';
        show($('prResult'));
      } else {
        $('prError').textContent='错误：'+d.error;
        show($('prError'));
        $('prStatus').textContent='';
      }
    }catch(e){$('prError').textContent=e.message;show($('prError'));$('prStatus').textContent='';}
    $('prRun').disabled=false;
  }

  // Register init hook
  window._toolInits = window._toolInits||{};
  window._toolInits['pref-rate'] = prInit;
})();

// ═══════════════════ 页面加载 ═══════════════════
const initialTool = decodeURIComponent((location.hash || '').replace(/^#/, ''));
if(initialTool && TOOLS[initialTool]) switchTool(initialTool);
else showOverview();
