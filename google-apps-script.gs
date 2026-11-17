/**
 * SPC Inventory -> Google Sheets archive (Google Apps Script).
 * Deploy as a Web App: Execute as "Me"; access "Anyone".
 * Create Script Property SPC_SYNC_KEY with a long random secret.
 * Never commit the secret to GitHub.
 */
const SPC_SHEETS = Object.freeze({
  "Jesup":"1g0uOU0kVxzIGZJLX_9nCQE6NVLOx7RTGq8qY2gC7ykQ",
  "Richmond Hill":"1zNDwUNFh-u6r7uplolt295j2ARCdmLKcobbrwhtjSTQ",
  "Ogeechee":"1G_-Zjr6ITDZdrnAQ2QE0w8nIqZa6L5RjYO150eKEGcs",
  "Claxton":"1J7tw4_XpltoDRaIMAGG_7ZXt24Envd8JeRc86GxMih0",
  "Blackshear":"1-mYDzbtseQikCCqkme4NByKoKeZACw8a6anp3WKxPEs"
});
function asText(v) {
  const t=String(v == null ? "" : v).slice(0,500);
  return /^[=+@-]/.test(t) ? "'"+t : t; // prevent accidental sheet formulas
}
function response(result){
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
function doGet(){ return response({app:"SPC Inventory",status:"ready"}); }
function doPost(e){
  const lock=LockService.getScriptLock();
  try {
    const data=JSON.parse(e && e.postData && e.postData.contents || "{}");
    const expected=PropertiesService.getScriptProperties().getProperty("SPC_SYNC_KEY");
    if (!expected || expected.length<24) throw Error("SYNC_KEY_NOT_CONFIGURED");
    if (String(data.key||"") !== expected) throw Error("UNAUTHORIZED");
    const loc=String(data.location||"");
    if (!Object.prototype.hasOwnProperty.call(SPC_SHEETS,loc)) throw Error("UNKNOWN_LOCATION");
    const countId=String(data.countId||"");
    if(!/^\d{10,20}$/.test(countId)) throw Error("INVALID_COUNT_ID");
    const localDate=String(data.localDate||"");
    if(!/^\d{4}-\d{2}-\d{2}$/.test(localDate)) throw Error("INVALID_DATE");
    const lines=Array.isArray(data.lines)?data.lines:[];
    const catalog=Array.isArray(data.catalog)?data.catalog:[];
    if(!lines.length || lines.length>1000 || catalog.length>1000)throw Error("INVALID_ITEMS");
    if(!lines.every(x=>Number.isFinite(Number(x.quantity)) && Number(x.quantity)>=0))throw Error("INVALID_QUANTITY");
    lock.waitLock(30000);
    const ss=SpreadsheetApp.openById(SPC_SHEETS[loc]);
    for (const old of ss.getSheets()) {
      if(old.getName()!=="START HERE" && String(old.getRange("B2").getDisplayValue())===countId) {
        return response({ok:true,duplicate:true,sheet:old.getName()});
      }
    }
    const earlier=ss.getSheets().find(x=>x.getName()!=="START HERE");
    const oldPars={};
    if(earlier && earlier.getLastRow()>5){
      const existing=earlier.getRange(6,1,earlier.getLastRow()-5,9).getValues();
      existing.forEach(r=>{if(r[8] && r[4]!=="" && r[4]!==null)oldPars[String(r[8])]=r[4];});
    }
    const nameBase=localDate.slice(5,7)+"/"+localDate.slice(8,10);
    let tabName=nameBase, suffix=2;
    while(ss.getSheetByName(tabName))tabName=nameBase+" ("+(suffix++)+")";
    const tab=ss.insertSheet(tabName,0);
    tab.setHiddenGridlines(true);
    tab.setFrozenRows(5);
    tab.setTabColor("#BD2633");
    tab.getRange("A1:I1").merge().setValue(loc.toUpperCase()+" — INVENTORY COUNT");
    tab.getRange("A2:B2").setValues([["COUNT ID",countId]]);
    tab.getRange("D2:E2").setValues([["COMPLETED",localDate]]);
    tab.getRange("A3:I3").merge().setValue("Savannah Pizza Company  |  Counts recorded in full or decimal cases");
    tab.getRange("A5:I5").setValues([[
       "SUPPLIER SKU","INVENTORY ITEM","CASE PACK","COUNT (CASES)","PAR (CASES)",
       "SUGGESTED ORDER","CATEGORY","COUNT ID","ITEM ID"
    ]]);
    const byId={};
    lines.forEach(x=>{if(x && x.itemId)byId[String(x.itemId)]=x;});
    const base=catalog.length?catalog:lines;
    const rows=[], categories=[];
    let category="";
    for(const item of base){
      if(!item)continue;
      const id=String(item.id||item.itemId||"");
      const line=byId[id] || (catalog.length ? null : item);
      const nextCat=String(item.category||"Other");
      if(nextCat!==category){
        category=nextCat;
        categories.push(rows.length+6);
        rows.push(["",asText(category.toUpperCase()),"","","","", "", "", ""]);
      }
      const pack=item.packCount && item.eachSize
        ? String(item.packCount)+" x "+String(item.eachSize)+(item.sizeUnit?" "+item.sizeUnit:"")
        : String(item.originalName||"").replace(String(item.name||""),"").trim();
      const count=line && line.quantity!=="" && line.quantity!=null ? Number(line.quantity) : "";
      const par=Object.prototype.hasOwnProperty.call(oldPars,id)?oldPars[id]:"";
      const row=rows.length+6;
      rows.push([asText(item.supplierSku||""),asText(item.name||item.originalName||""),
       asText(pack),count,par,
       '=IF(OR(D'+row+'="",E'+row+'=""),"",MAX(0,ROUNDUP(E'+row+'-D'+row+',0)))',
       asText(category),asText(countId),asText(id)]);
    }
    if(!rows.length)throw Error("NO_ITEMS");
    tab.getRange(6,1,rows.length,9).setValues(rows);
    tab.getRange("A1:I1").setBackground("#BD2633").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(16);
    tab.getRange("A3:I3").setBackground("#F8F5F2").setFontColor("#6B7280").setFontSize(10);
    tab.getRange("A5:I5").setBackground("#EAE6E2").setFontWeight("bold").setWrap(true);
    tab.getRange(6,4,rows.length,3).setNumberFormat("0.###");
    for(const n of categories){
      tab.getRange(n,1,1,9).setBackground("#F2EFEB").setFontColor("#9E2029").setFontWeight("bold");
    }
    const widths=[130,310,150,135,125,140,135,145,120];
    widths.forEach((width,i)=>tab.setColumnWidth(i+1,width));
    tab.setRowHeight(1,38);tab.setRowHeight(5,34);
    tab.hideColumns(8,2);
    const starter=ss.getSheetByName("START HERE");
    if(starter && ss.getSheets().length>1)ss.deleteSheet(starter);
    SpreadsheetApp.flush();
    return response({ok:true,sheet:tabName,location:loc,rows:rows.length});
  }catch(err){
    return response({ok:false,error:String(err.message||err)});
  }finally{
    try{lock.releaseLock();}catch(_){}
  }
}
