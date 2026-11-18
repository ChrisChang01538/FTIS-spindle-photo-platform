/**
 * 主軸元件照片蒐集平台 — 後端（Google Apps Script）
 * 功能：接收前端（GitHub Pages）上傳的照片 → 依「機構/廢物品」分資料夾存入 Google Drive，
 *       紀錄寫入 Google 試算表；提供設定、清單、刪除 API。
 *
 * 部署步驟請見 README.md。
 */

/* ======================= 請依需求修改 ======================= */
const ADMIN_CODE  = 'change-me';   // 管理密碼：修改設定、刪除照片時需要（請務必改掉）
const UPLOAD_CODE = '';            // 上傳通行碼：留空＝任何人開網址都能上傳；填值＝現場人員需輸入一次
const ROOT_FOLDER_NAME = '主軸元件照片蒐集';   // Drive 根資料夾名稱（第一次執行自動建立）
const SHARE_PUBLIC = true;         // true：照片設為「知道連結者可檢視」，前端才能顯示縮圖；false：僅你本人可看
/* ============================================================ */

const SHEET_PHOTOS = '照片';
const SHEET_CONFIG = '設定';
const HEADER = ['id','檔名','機構代碼','機構名稱','類別','廢物品','主軸元件','階段','序號/備註','上傳者','上傳時間','大小(bytes)','Drive檔案ID','Drive連結'];

/** 第一次部署前請在編輯器手動執行一次此函式，完成授權並建立資料夾／試算表。 */
function setup(){
  const root = getRoot();
  const ss = getSpreadsheet_();
  Logger.log('Drive 資料夾：' + root.getUrl());
  Logger.log('紀錄試算表：' + ss.getUrl());
  return { folder: root.getUrl(), sheet: ss.getUrl() };
}

function doGet(e){ return handle_(e && e.parameter ? e.parameter : {}); }
function doPost(e){
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch(err) { return json_({ ok:false, error:'JSON 解析失敗' }); }
  return handle_(body);
}

function handle_(req){
  try{
    switch(req.action){
      case 'ping':       return json_({ ok:true, time:new Date().toISOString(), needUploadCode: !!UPLOAD_CODE });
      case 'config':     return json_({ ok:true, config:getConfig_() });
      case 'saveConfig': requireAdmin_(req); setConfig_(req.config); return json_({ ok:true });
      case 'upload':     return json_(upload_(req));
      case 'list':       return json_({ ok:true, photos:list_(), folderUrl:getRoot().getUrl(), sheetUrl:getSpreadsheet_().getUrl() });
      case 'delete':     requireAdmin_(req); return json_(remove_(req.id));
      case 'checkAdmin': requireAdmin_(req); return json_({ ok:true });
      default:           return json_({ ok:false, error:'未知的 action' });
    }
  }catch(err){
    return json_({ ok:false, error: String(err && err.message || err) });
  }
}

function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function requireAdmin_(req){ if(!ADMIN_CODE || req.admin !== ADMIN_CODE) throw new Error('管理密碼錯誤'); }

/* ---------- Drive / Sheet ---------- */
function getRoot(){
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ROOT_FOLDER_ID');
  if(id){ try{ return DriveApp.getFolderById(id); }catch(e){} }
  const it = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(ROOT_FOLDER_NAME);
  props.setProperty('ROOT_FOLDER_ID', folder.getId());
  return folder;
}
function subFolder_(parent, name){
  name = sanitize_(name) || '未分類';
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function getSpreadsheet_(){
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if(id){ try{ return SpreadsheetApp.openById(id); }catch(e){} }
  const ss = SpreadsheetApp.create('主軸元件照片紀錄');
  const file = DriveApp.getFileById(ss.getId());
  getRoot().addFile(file); DriveApp.getRootFolder().removeFile(file);
  const sh = ss.getSheets()[0]; sh.setName(SHEET_PHOTOS); sh.appendRow(HEADER); sh.setFrozenRows(1);
  ss.insertSheet(SHEET_CONFIG).getRange('A1').setValue('');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}
function photosSheet_(){ const ss=getSpreadsheet_(); return ss.getSheetByName(SHEET_PHOTOS) || ss.insertSheet(SHEET_PHOTOS); }
function configSheet_(){ const ss=getSpreadsheet_(); return ss.getSheetByName(SHEET_CONFIG) || ss.insertSheet(SHEET_CONFIG); }

function getConfig_(){
  const v = configSheet_().getRange('A1').getValue();
  if(!v) return null;
  try{ return JSON.parse(v); }catch(e){ return null; }
}
function setConfig_(cfg){
  if(!cfg || !Array.isArray(cfg.plants) || !Array.isArray(cfg.categories)) throw new Error('設定格式不正確');
  cfg.updatedAt = new Date().toISOString();
  configSheet_().getRange('A1').setValue(JSON.stringify(cfg));
}

/* ---------- 上傳 ---------- */
function upload_(req){
  if(UPLOAD_CODE && req.code !== UPLOAD_CODE) throw new Error('上傳通行碼錯誤');
  const m = req.meta || {};
  if(!req.data) throw new Error('沒有照片資料');
  if(!m.fileName) throw new Error('缺少檔名');
  const bytes = Utilities.base64Decode(req.data);
  if(bytes.length > 25*1024*1024) throw new Error('檔案過大');

  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const root = getRoot();
    const plantFolder = subFolder_(root, (m.plant||'') + (m.plantName ? '_' + m.plantName : ''));
    const itemFolder = subFolder_(plantFolder, m.item || '未分類');
    let name = sanitize_(m.fileName); if(!/\.jpe?g$/i.test(name)) name += '.jpg';
    name = uniqueName_(itemFolder, name);
    const blob = Utilities.newBlob(bytes, 'image/jpeg', name);
    const file = itemFolder.createFile(blob);
    if(SHARE_PUBLIC){ try{ file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(e){} }
    const id = Utilities.getUuid();
    const now = new Date().toISOString();
    photosSheet_().appendRow([id, name, m.plant||'', m.plantName||'', m.cat||'', m.item||'', m.comp||'', m.stage||'', m.serial||'', m.uploader||'', now, bytes.length, file.getId(), file.getUrl()]);
    return { ok:true, id:id, fileName:name, fileId:file.getId(), url:file.getUrl(), uploadedAt:now, size:bytes.length };
  } finally { lock.releaseLock(); }
}
function uniqueName_(folder, name){
  if(!folder.getFilesByName(name).hasNext()) return name;
  const base = name.replace(/\.jpg$/i,''); let i=2;
  while(folder.getFilesByName(base+'_'+i+'.jpg').hasNext()) i++;
  return base+'_'+i+'.jpg';
}
function sanitize_(s){ return String(s==null?'':s).replace(/[\\/:*?"<>|\s]+/g,'_').replace(/_+/g,'_').replace(/^_+|_+$/g,''); }

/* ---------- 清單 / 刪除 ---------- */
function list_(){
  const sh = photosSheet_(); const last = sh.getLastRow(); if(last < 2) return [];
  const rows = sh.getRange(2, 1, last-1, HEADER.length).getValues();
  const out = [];
  rows.forEach(r=>{ if(!r[0]) return; out.push({ id:r[0], fileName:r[1], plant:r[2], plantName:r[3], cat:r[4], item:r[5], comp:r[6], stage:r[7], serial:r[8], uploader:r[9], uploadedAt: (r[10] instanceof Date ? r[10].toISOString() : String(r[10])), size:r[11], fileId:r[12], url:r[13] }); });
  return out.reverse();
}
function remove_(id){
  const sh = photosSheet_(); const last = sh.getLastRow(); if(last<2) return { ok:false, error:'找不到紀錄' };
  const ids = sh.getRange(2,1,last-1,1).getValues();
  for(let i=0;i<ids.length;i++){
    if(ids[i][0] === id){
      const fileId = sh.getRange(i+2, 13).getValue();
      try{ if(fileId) DriveApp.getFileById(fileId).setTrashed(true); }catch(e){}
      sh.deleteRow(i+2);
      return { ok:true };
    }
  }
  return { ok:false, error:'找不到紀錄' };
}
