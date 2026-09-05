<!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<title>Tra cứu giá khuôn</title>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
*{box-sizing:border-box}
body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;background:#f4f6f8;color:#18212b}
.wrap{max-width:1180px;margin:auto;padding:24px 18px 40px}
.status{display:inline-block;background:#e8f7ed;color:#20833f;border-radius:999px;padding:9px 14px;font-weight:700;margin-bottom:18px}
h1{font-size:32px;line-height:1.15;margin:4px 0}
.sub{color:#68727d;font-size:18px;margin:0 0 22px}
.layout{display:grid;grid-template-columns:minmax(0,1.35fr) minmax(320px,.65fr);gap:18px;align-items:start}
.panel{background:#fff;border-radius:20px;padding:18px;box-shadow:0 3px 14px rgba(0,0,0,.07)}
.panel-title{font-size:21px;font-weight:850;margin-bottom:6px}
.panel-help{font-size:13px;color:#68727d;line-height:1.45;margin-bottom:14px}
.search{display:flex;gap:10px;align-items:stretch}
input{min-width:0;flex:1;padding:15px;border:2px solid #ccd4dc;border-radius:15px;font-size:18px;outline:none;background:#fff;-webkit-appearance:none}
input:focus{border-color:#526171}
button{border:0;border-radius:14px;font-size:17px;font-weight:800;cursor:pointer;-webkit-tap-highlight-color:transparent}
#searchBtn{padding:0 22px;background:#1f2937;color:white;min-width:90px}
.toolbar{display:flex;gap:10px;flex-wrap:wrap;margin-top:12px}
.import{padding:12px 18px;background:#dbeafe;color:#1d4ed8}
.clear{padding:12px 18px;background:#e9eef3;color:#354152}
.file-name{font-size:13px;color:#68727d;margin-top:8px;word-break:break-word}
#result{margin-top:18px}
.empty{text-align:center;color:#68727d;padding:28px 10px;font-size:17px}
.count{font-weight:700;color:#4d5966;margin:12px 2px}
.card{background:#fff;border:1px solid #e5e9ee;border-radius:18px;margin-top:12px;padding:18px;box-shadow:0 2px 10px rgba(0,0,0,.05)}
.label{font-size:12px;color:#7a8490;text-transform:uppercase;margin-top:10px}
.value{font-size:17px;font-weight:600;margin-top:4px;word-break:break-word}
.price{font-size:30px;font-weight:800;margin-top:6px;color:#111827}
.status-badge{display:inline-flex;align-items:center;gap:6px;padding:7px 11px;border-radius:999px;font-size:13px;font-weight:800;margin-top:10px}.status-ok{background:#e8f7ed;color:#20833f}.status-repair{background:#fff4e5;color:#b45309}.status-dot{font-size:16px;line-height:1}
.online{background:#f8fafc;border:1px solid #e4e8ed;border-radius:16px;padding:15px;margin:0}
.online-title{font-weight:800;font-size:17px;margin-bottom:7px}
.online-help{font-size:13px;color:#68727d;line-height:1.45;margin-bottom:10px}
.online-row{display:flex;gap:8px;flex-wrap:wrap}
#sheetUrl{flex:1 1 100%;min-width:0;padding:12px 13px;border:2px solid #ccd4dc;border-radius:12px;font-size:14px;outline:none;background:#fff}
#sheetUrl:focus{border-color:#526171}
#saveUrlBtn,#refreshBtn{padding:11px 14px;color:#fff;flex:1}
#saveUrlBtn{background:#374151}
#refreshBtn{background:#0f766e}
.online-meta{font-size:12px;color:#68727d;margin-top:10px;display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap}
.online-note{font-size:13px;color:#68727d;margin:12px 0 0;line-height:1.45}
.function-block{margin-top:16px}
.function-title{font-size:15px;font-weight:800;margin-bottom:8px}
@media(max-width:820px){
  .layout{grid-template-columns:1fr}
}
@media(max-width:430px){
  .wrap{padding:18px 12px 30px}
  h1{font-size:28px}
  .sub{font-size:17px}
  .panel{padding:14px}
  .search{gap:8px}
  #searchBtn{min-width:78px;padding:0 15px}
  .online-row{flex-direction:column}
  #saveUrlBtn,#refreshBtn{width:100%}
}
</style>
<style>
.tabs{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:18px}
.tab{padding:17px 14px;border-radius:18px;background:#e9eef6;color:#344154;font-size:18px;font-weight:800}
.tab.active{background:#1f2937;color:#fff}
.section-panel{background:#fff;border-radius:20px;padding:18px;box-shadow:0 3px 14px rgba(0,0,0,.07)}
.hidden{display:none}
@media(max-width:430px){
  .tabs{gap:10px}
  .tab{font-size:17px;padding:16px 8px}
  .section-panel{padding:14px}
  .search{gap:8px}
  #searchBtn{min-width:78px;padding:0 15px}
}
</style>

</head>
<body>
<div class="wrap">
  <div class="status" id="status">✓ Sẵn sàng tra cứu</div>
  <h1>TRA CỨU GIÁ KHUÔN</h1>
  <div class="sub">Tra cứu nhanh theo Item Code, Item Name hoặc Process Note</div>

  <div class="tabs" role="tablist" aria-label="Chọn chức năng">
    <button class="tab active" id="tabSearch" type="button" role="tab" aria-selected="true">🔎 Tra cứu</button>
    <button class="tab" id="tabFunctions" type="button" role="tab" aria-selected="false">⚙️ Chức năng khác</button>
  </div>

  <div id="searchSection" class="section-panel">
    <div class="panel-title">🔎 TRA CỨU</div>
    <div class="panel-help">Nhập một phần hoặc toàn bộ Item Code, Item Name hoặc Process Note để tìm thông tin khuôn.</div>

    <form class="search" id="searchForm">
      <input id="q" type="search" placeholder="Ví dụ: 032E 63220, Item Name hoặc A11445"
             autocomplete="off" autocapitalize="characters" spellcheck="false">
      <button id="searchBtn" type="submit">Tìm</button>
    </form>

    <div class="toolbar">
      <button class="clear" id="clearBtn" type="button">Xóa</button>
    </div>

    <div id="result" class="empty">Nhập Item Code, Item Name hoặc Process Note để tra giá khuôn.</div>
  </div>

  <div id="functionsSection" class="section-panel hidden">
    <div class="panel-title">⚙️ CHỨC NĂNG KHÁC</div>
    <div class="panel-help">Nhập dữ liệu và đồng bộ dữ liệu online.</div>

    <div class="function-block">
      <div class="function-title">📥 Nhập dữ liệu</div>
      <button class="import" id="importBtn" type="button">📥 Import Excel/CSV</button>
      <button id="uploadToSheetBtn" type="button" style="padding:12px 18px;background:#dcfce7;color:#166534;">☁️ Đẩy dữ liệu lên Google Sheets</button>
      <input id="fileInput" type="file" accept=".xlsx,.xls,.csv,.json" hidden>
      <div class="file-name" id="fileName"></div>
    </div>

    <div class="function-block">
      <div class="function-title">🌐 Dữ liệu online</div>
        <div class="online-help"><b>Một lần thiết lập:</b> nhập Web App URL của Google Apps Script bên dưới. Sau đó máy chính có thể Import Excel và gửi dữ liệu lên Google Sheets; các máy khác tự đọc dữ liệu chung.</div>
        <input id="appsScriptUrl" type="url" placeholder="Dán Web App URL của Google Apps Script" autocomplete="off" style="width:100%;padding:12px 13px;border:2px solid #ccd4dc;border-radius:12px;font-size:14px;margin-bottom:9px;">
        <button id="saveAppsScriptBtn" type="button" style="width:100%;padding:12px;background:#374151;color:#fff;margin-bottom:10px;">Lưu Web App URL</button>
      <div class="online">
        <div class="online-help">Dán link Google Sheets để các máy dùng chung một nguồn dữ liệu. Máy khác có thể tự cập nhật dữ liệu mới.</div>
        <div class="online-row">
          <input id="sheetUrl" type="url" placeholder="Dán link Google Sheets tại đây" autocomplete="off">
          <button id="saveUrlBtn" type="button">Lưu link</button>
          <button id="refreshBtn" type="button">↻ Cập nhật</button>
        </div>
        <div class="online-meta">
          <span id="syncText">Chưa kết nối online — đang dùng dữ liệu có sẵn.</span>
          <span id="autoText">Tự cập nhật: 15 giây</span>
        </div>
      </div>
      <div class="online-note"><b>Máy chính:</b> cập nhật Google Sheets. <b>Máy khác:</b> mở web và tra cứu.</div>
    </div>
  </div>
</div>
<script>
let DATA = [{"Item Code":"130B-002V001","Item Name":"BATTERY TAPE","Process Note":"R-A12618- Chung @N-13F1-E6U0200-E0C0-1Q","Gía khuôn":"1A1","Trạng thái":"OK"},{"Item Code":"2PN00001","Item Name":"RC5-0042-000(Tape 29*70)","Process Note":"R/6CAV:A8811+A8812","Gía khuôn":"1A2+1A3","Trạng thái":"OK"},{"Item Code":"2PN800051","Item Name":"RC5-4056-000(Tape 29*70)","Process Note":"R-A8812+A9561","Gía khuôn":"1A3+1A4","Trạng thái":"OK"},{"Item Code":"2PN600002","Item Name":"302LV11751(TAPE)","Process Note":"Rotary: A9562+A9563","Gía khuôn":"1A5+1A6","Trạng thái":"OK"},{"Item Code":"2PN600000","Item Name":"302LV11741(TAPE)","Process Note":"Rotary: A9562+A9801","Gía khuôn":"1A5+1A7","Trạng thái":"OK"},{"Item Code":"30C0D14270","Item Name":"SHEET COVER MAG B","Process Note":"R-A9960+A10657","Gía khuôn":"1A8+1A9","Trạng thái":"OK"},{"Item Code":"5-030-001-02","Item Name":"TAPE BLT","Process Note":"Q-A12685; A12841; A12832","Gía khuôn":"2A1+2A2+2A3","Trạng thái":"OK"},{"Item Code":"D02CJ7001","Item Name":"BLADE FILM FCL","Process Note":"A11999+A12000+A12001","Gía khuôn":"2A4+2A5+2A6","Trạng thái":"OK"},{"Item Code":"D02CJ7001","Item Name":"BLADE FILM FCL","Process Note":"A12500+A12501+A12502","Gía khuôn":"2A7+2A8+2A9","Trạng thái":"OK"},{"Item Code":"5-054-633-01","Item Name":"BATTERY TAPE","Process Note":"rotary/A12618","Gía khuôn":"3A1","Trạng thái":"OK"},{"Item Code":"9PN134000","Item Name":"QC2-6384-000 (TAPE)","Process Note":"PC1-A10033+PC2-A11058","Gía khuôn":"3A2+3A3","Trạng thái":"OK"},{"Item Code":"D01B6A001","Item Name":"TAPE 374X5","Process Note":"R-A9357(khuôn D00416001)","Gía khuôn":"3A4","Trạng thái":"OK"},{"Item Code":"D01KRL001","Item Name":"LOWER FILM FCL","Process Note":"R-A12467+A11239+A11240","Gía khuôn":"3A5+3A6+3A7","Trạng thái":"OK"},{"Item Code":"FEH1001-A","Item Name":"DS TAPE","Process Note":"R-A12768+A12769","Gía khuôn":"3A8+3A9","Trạng thái":"OK"},{"Item Code":"D00AD7001","Item Name":"LOWER FILM ELL","Process Note":"R-A12617+A12054","Gía khuôn":"4A1+4A2","Trạng thái":"OK"},{"Item Code":"D03CWF001","Item Name":"LOWER FILM ELL B","Process Note":"NC Rotary-A12071","Gía khuôn":"4A3","Trạng thái":"OK"},{"Item Code":"LY2486001","Item Name":"LOWER FILM BLL","Process Note":"R-A11445","Gía khuôn":"4A4","Trạng thái":"OK"},{"Item Code":"LY8261001","Item Name":"LOWER FILM DSL","Process Note":"NC ROTARY-A8884+A13232","Gía khuôn":"4A5+4A6","Trạng thái":"OK"},{"Item Code":"LY9203001","Item Name":"LOWER FILM DLL","Process Note":"R-A10046+A13077","Gía khuôn":"4A7+4A8","Trạng thái":"OK"},{"Item Code":"D00360001","Item Name":"LOWER FILM DL","Process Note":"R-A9093+A9094","Gía khuôn":"4A9","Trạng thái":"OK"},{"Item Code":"032E 3552-T2","Item Name":"DOUBLE SIDE TAPE","Process Note":"R-A10419+A9630","Gía khuôn":"5A1+5A2","Trạng thái":"OK"},{"Item Code":"035K 11070","Item Name":"SEAL ASSY-PLATEN, LEFT","Process Note":"ROTARY-A12375","Gía khuôn":"5A3","Trạng thái":"OK"},{"Item Code":"035K 11080","Item Name":"SEAL ASSY-PLATEN, RIGHT","Process Note":"ROTARY-A12378","Gía khuôn":"5A4","Trạng thái":"OK"},{"Item Code":"032K 15060-AB3832","Item Name":"","Process Note":"R-A12301+A12302","Gía khuôn":"5A5+5A6","Trạng thái":"OK"},{"Item Code":"032K 15067-AB3832","Item Name":"","Process Note":"R-12301+A13106","Gía khuôn":"5A6+5A7","Trạng thái":"OK"},{"Item Code":"QC5-8994-000","Item Name":"ADF GLASS SHEET G181","Process Note":"R-A12897/Dap tu mat PET","Gía khuôn":"5A8","Trạng thái":"OK"},{"Item Code":"UL9159001","Item Name":"LOWER FILM","Process Note":"R-A9847+JIG tròn A11135","Gía khuôn":"5A9","Trạng thái":"OK"},{"Item Code":"308000477630","Item Name":"TAPE HARNESS(10x15)","Process Note":"R/50cav A12460","Gía khuôn":"1B1","Trạng thái":"OK"},{"Item Code":"032E 63220","Item Name":"GUIDE-LH HCF","Process Note":"R-A12059+JIG","Gía khuôn":"1B2","Trạng thái":"OK"},{"Item Code":"063E 15750","Item Name":"TAPE-CVT, 2","Process Note":"R-A11316 khuôn chung","Gía khuôn":"1B3","Trạng thái":"OK"},{"Item Code":"063E 15780","Item Name":"TAPE-PLTN, 2","Process Note":"R-A10374","Gía khuôn":"1B4","Trạng thái":"OK"},{"Item Code":"5058332-11","Item Name":"CUSHION, ENCLOSURE (UM)","Process Note":"A10374 cắt10W; A8926 khuôn 1 lưỡi cắt chiều 60W","Gía khuôn":"1B4","Trạng thái":"OK"},{"Item Code":"063E 15790","Item Name":"TAPE-PLTN, 3","Process Note":"R-A8845","Gía khuôn":"1B5","Trạng thái":"OK"},{"Item Code":"2PN00004","Item Name":"PET16-SN(67L*27W)","Process Note":"R-A13039  khuôn 303PH02040","Gía khuôn":"1B6","Trạng thái":"OK"},{"Item Code":"835E 44100","Item Name":"SEAL-UPPER","Process Note":"R-A10622+JIG tròn A11136","Gía khuôn":"1B7","Trạng thái":"OK"},{"Item Code":"2PN800049","Item Name":"QC4-9578-000(Tape 18*18)","Process Note":"A11756-CHUNG QC4-9783-000","Gía khuôn":"1B8","Trạng thái":"OK"},{"Item Code":"2PN800110","Item Name":"QC4-9783+QC7-1521(Tape 18*18)","Process Note":"NC+R-A11756","Gía khuôn":"1B8","Trạng thái":"OK"},{"Item Code":"2PN800141","Item Name":"302NT03020-semi 18W (5000NS+#20W-0.2T)","Process Note":"R-A12123","Gía khuôn":"1B9","Trạng thái":"OK"},{"Item Code":"2PN900113","Item Name":"#8100CH-2-0.1T+LMRS-50#","Process Note":"A11673+ NL T5005-EX-0.05T để lót","Gía khuôn":"1B10","Trạng thái":"OK"},{"Item Code":"302NR12320","Item Name":"SHEET DRUM","Process Note":"Rotary-A12403 + JIG K08","Gía khuôn":"1B11","Trạng thái":"OK"},{"Item Code":"3V2LV24220","Item Name":"SHEET REGIST","Process Note":"R-A11887","Gía khuôn":"2B1","Trạng thái":"OK"},{"Item Code":"3V2NM17460","Item Name":"TAPE CONTACT GLASS C","Process Note":"R-A9196 (NC ROTARY)","Gía khuôn":"2B2","Trạng thái":"OK"},{"Item Code":"3V2P711120","Item Name":"SHEET MC CLN","Process Note":"Rotary-A12310","Gía khuôn":"2B3","Trạng thái":"OK"},{"Item Code":"5-020-279-02","Item Name":"CUSHION LOWER","Process Note":"R-A13019","Gía khuôn":"2B4","Trạng thái":"OK"},{"Item Code":"5-030-000-02","Item Name":"CUSHION BLT","Process Note":"R-A13024","Gía khuôn":"2B5","Trạng thái":"OK"},{"Item Code":"5-030-002-01","Item Name":"CUSHION SPEAKER(BACK)","Process Note":"R-A12817","Gía khuôn":"2B6","Trạng thái":"OK"},{"Item Code":"5-051-681-01","Item Name":"TAPE HARNESS(10x15)","Process Note":"R/50cav A12460 BÓC bỏ cả 1 hàng khi có khuyết thiế","Gía khuôn":"2B7","Trạng thái":"OK"},{"Item Code":"5058332-01","Item Name":"CUSHION, ENCLOSURE (UM)","Process Note":"NC+R-A10374+A8926","Gía khuôn":"2B8+2B9","Trạng thái":"OK"},{"Item Code":"D02A79001","Item Name":"SCANNER BASE FILM ELLEFB","Process Note":"ROTARY+NC A12503+ A13201","Gía khuôn":"2B10+2B11","Trạng thái":"OK"},{"Item Code":"302YJ21880","Item Name":"","Process Note":"R-A12494","Gía khuôn":"3B1","Trạng thái":"OK"},{"Item Code":"D01RA5001","Item Name":"","Process Note":"R-A11968+A12045","Gía khuôn":"3B2+3B3","Trạng thái":"OK"},{"Item Code":"9PN700022","Item Name":"3V2LV02600( LMRS-75#+ #8800CH-99.1W*443L)","Process Note":"A9925+ JIG K77+ CẮT 443+-2","Gía khuôn":"3B4","Trạng thái":"OK"},{"Item Code":"D000PT001","Item Name":"WHITE FILM DLFB","Process Note":"NC+R-A9597","Gía khuôn":"3B5","Trạng thái":"OK"},{"Item Code":"D00416001","Item Name":"SU SUPPORT SEAL DL","Process Note":"R-A11121.có chiều keo, chiều dài sp keo ko co giãn","Gía khuôn":"3B6","Trạng thái":"OK"},{"Item Code":"D01WDY001","Item Name":"CLN LOWER FILM EL","Process Note":"A11840","Gía khuôn":"3B7","Trạng thái":"OK"},{"Item Code":"FE8-5771-000","Item Name":"TAPE_HOLDER_TOP_CIS-2","Process Note":"NC R-A9439 Chung 3V2NM17440","Gía khuôn":"3B8","Trạng thái":"OK"},{"Item Code":"FL0-0852-000","Item Name":"SHEET-FLOWSCAN-GLASS1","Process Note":"R-A9823","Gía khuôn":"3B9","Trạng thái":"OK"},{"Item Code":"K530022500A","Item Name":"Seal","Process Note":"Q-A11938","Gía khuôn":"3B10","Trạng thái":"OK"},{"Item Code":"LJA427001","Item Name":"WLAN TAPE DSLE2","Process Note":"R-A11175","Gía khuôn":"3B11","Trạng thái":"OK"},{"Item Code":"LP2593001","Item Name":"WHITE FILM","Process Note":"R-A10144(Rotary pina tròn)","Gía khuôn":"4B1","Trạng thái":"OK"},{"Item Code":"LX4226001","Item Name":"TAPE 359X5","Process Note":"R-A9357(khuôn D00416001)","Gía khuôn":"4B2","Trạng thái":"OK"},{"Item Code":"QC3-6113-000","Item Name":"DOUBLE SIDE TAPE L11 (5X218)","Process Note":"chung khuôn A9357","Gía khuôn":"4B2","Trạng thái":"OK"},{"Item Code":"LX5210001","Item Name":"SCANNER BASE FILM BLL","Process Note":"ROTARY+NC  A5120","Gía khuôn":"4B3","Trạng thái":"OK"},{"Item Code":"QC2-8345-000","Item Name":"CR FFC SHEET L81","Process Note":"PRESS/R-A9752","Gía khuôn":"4B4","Trạng thái":"OK"},{"Item Code":"QC5-1208-000","Item Name":"LCD BLIND SHEET G56","Process Note":"R-A11721","Gía khuôn":"4B5","Trạng thái":"OK"},{"Item Code":"QC7-9167-000","Item Name":"FABRIC SHEET A L202","Process Note":"R-A12934","Gía khuôn":"4B6","Trạng thái":"OK"},{"Item Code":"QC7-9168-000","Item Name":"FABRIC SHEET B L202","Process Note":"R-A11283","Gía khuôn":"4B7","Trạng thái":"OK"},{"Item Code":"QC7-9169-000","Item Name":"FABRIC SHEET C L202","Process Note":"R-A11284","Gía khuôn":"4B8","Trạng thái":"OK"},{"Item Code":"QX1-3132-000","Item Name":"Z60_MERCHANDISE CODE LABEL","Process Note":"R-13234","Gía khuôn":"4B9","Trạng thái":"OK"},{"Item Code":"RC4-3609-000","Item Name":"SCRAPER","Process Note":"R-A10590","Gía khuôn":"4B10","Trạng thái":"OK"},{"Item Code":"V9-08047F-00","Item Name":"LABEL","Process Note":"NC Rotary- A11640","Gía khuôn":"4B11","Trạng thái":"OK"},{"Item Code":"D00AD7001 - Cutter 74T","Item Name":"","Process Note":"R-A12977","Gía khuôn":"5B1","Trạng thái":"OK"},{"Item Code":"D00AD7001 - Cutter 74T","Item Name":"","Process Note":"R-A12912","Gía khuôn":"5B2","Trạng thái":"OK"},{"Item Code":"D00AD7001 - Cutter 74T","Item Name":"","Process Note":"R-A11807","Gía khuôn":"5B3","Trạng thái":"OK"},{"Item Code":"D00AD7001 - Cutter 74T","Item Name":"","Process Note":"R-A8818","Gía khuôn":"5B4","Trạng thái":"OK"},{"Item Code":"Rotary 10 cutter 80T","Item Name":"","Process Note":"R-A13027","Gía khuôn":"5B5","Trạng thái":"OK"},{"Item Code":"Rotary 10 cutter 80T","Item Name":"","Process Note":"R-A12930","Gía khuôn":"5B6","Trạng thái":"OK"},{"Item Code":"Rotary cutter 50T","Item Name":"","Process Note":"R-A12418","Gía khuôn":"5B7","Trạng thái":"OK"},{"Item Code":"Rotary cutter 60T","Item Name":"","Process Note":"R- A12518","Gía khuôn":"5B8","Trạng thái":"OK"},{"Item Code":"Rotary cutter 74T - 250L","Item Name":"","Process Note":"R-A13148","Gía khuôn":"5B9","Trạng thái":"OK"},{"Item Code":"Rotary 15 cutter 72T","Item Name":"","Process Note":"R-A13233","Gía khuôn":"5B10","Trạng thái":"OK"},{"Item Code":"3V2WL17030 - EOC 1P04 (48T)","Item Name":"","Process Note":"R-A9934","Gía khuôn":"5B11","Trạng thái":"OK"},{"Item Code":"GH02-28072A","Item Name":"TAPE DOUBLE FACE BATTERY","Process Note":"R-A12848+A12849+A12850+A12851","Gía khuôn":"1C1+1C2+1C3+1C4","Trạng thái":"OK"},{"Item Code":"25B0805-C53r0","Item Name":"","Process Note":"R-A12742+A12743+A12744+A12745","Gía khuôn":"1C5+1C6+1C7+1C8","Trạng thái":"OK"},{"Item Code":"Rotary slit tooling","Item Name":"","Process Note":"R-A12168","Gía khuôn":"1C9","Trạng thái":"OK"},{"Item Code":"GH02-28072A","Item Name":"TAPE DOUBLE FACE BATTERY","Process Note":"R-A12969+A12970+A12971+A12972","Gía khuôn":"2C1+2C2+2C3+2C4","Trạng thái":"OK"},{"Item Code":"25B0805-C55r0","Item Name":"","Process Note":"R-A13035+A12747+A12748+A12749","Gía khuôn":"2C5+2C6+2C7+2C8","Trạng thái":"OK"},{"Item Code":"2PN200046","Item Name":"302LV04490( 75-C/PET-C1S-R2+VNS-T7083)","Process Note":"R-A9118(Kyoyo-Rotary slit)","Gía khuôn":"2C9","Trạng thái":"OK"},{"Item Code":"Rotary slit tooling","Item Name":"","Process Note":"R-A9118","Gía khuôn":"2C9","Trạng thái":"OK"},{"Item Code":"GH02-27982A","Item Name":"TAPE DOUBLE FACE-MAIN WINDOW BTM","Process Note":"R-A13226+A13227","Gía khuôn":"3C1+3C2","Trạng thái":"OK"},{"Item Code":"GH02-27982A","Item Name":"TAPE DOUBLE FACE-MAIN WINDOW BTM","Process Note":"R-A13017+A13018","Gía khuôn":"3C3+3C4","Trạng thái":"OK"},{"Item Code":"Magnet roll 70T","Item Name":"","Process Note":"R-A9253","Gía khuôn":"3C5","Trạng thái":"OK"},{"Item Code":"FL0-0711-000","Item Name":"SHEET_INSULATING","Process Note":"ROTARY A8921+A8922","Gía khuôn":"3C6+3C7","Trạng thái":"OK"},{"Item Code":"GH02-26755A","Item Name":"TAPE BATTERY","Process Note":"ROTARY-A12134","Gía khuôn":"3C8","Trạng thái":"OK"},{"Item Code":"NRAW-00223","Item Name":"BOTTOM PROTECTION FILM","Process Note":"R-A12018- chung NC04-00008","Gía khuôn":"3C9","Trạng thái":"OK"},{"Item Code":"QX1-3132-000","Item Name":"Z60_MERCHANDISE CODE LABEL","Process Note":"R-A11176  hoặc  A11179","Gía khuôn":"4C1+4C2","Trạng thái":"OK"},{"Item Code":"Y0000004422","Item Name":"TOP SHEET","Process Note":"R- A11814","Gía khuôn":"4C3","Trạng thái":"OK"}];
let DATA_SOURCE = "embedded";
let LOCAL_IMPORT_TIME = 0;

function normSearch(s) {
  return String(s ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"']/g, function(m) {
    return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[m];
  });
}

function headerKind(key) {
  var nk = normSearch(key);

  // Nhận diện các cột chính, không phụ thuộc dấu/ký tự đặc biệt.
  if (nk === "itemcode" || nk.includes("itemcode")) return "Item Code";
  if (nk === "itemname" || nk.includes("itemname") || nk.includes("tenitem")) return "Item Name";
  if (nk === "processnote" || nk.includes("processnote")) return "Process Note";
  if (nk === "trangthai" || nk.includes("trangthai") || nk === "status" || nk.includes("status") || nk === "tinhtrang") return "Trạng thái";

  // Giá khuôn / Vị trí khuôn.
  // Quan trọng: phải dùng đúng "giakhuon" (có chữ u), vì normSearch("Giá khuôn") = "giakhuon".
  if (
    nk === "giakhuon" ||
    nk.startsWith("giakhuon") ||
    nk.includes("giakhuonhanggiacot") ||
    nk.includes("vitrikhuon") ||
    nk.includes("vitrigiakhuon") ||
    nk.includes("vitrigiakhuon") ||
    nk.includes("moldlocation")
  ) return "Gía khuôn";

  return "";
}

function normalizeRow(row) {
  var out = {
    "Item Code": "",
    "Item Name": "",
    "Process Note": "",
    "Gía khuôn": "",
    "Trạng thái": ""
  };

  Object.keys(row || {}).forEach(function(k) {
    var kind = headerKind(k);
    if (!kind) return;

    if (kind === "Gía khuôn") {
      var nk = normSearch(k);

      // Prefer the actual "Giá khuôn (hàng-giá-cột)" / "Vị trí khuôn"
      // and never confuse it with "Khuôn dự phòng".
      var strongPrice =
        nk.startsWith("giakhon") ||
        nk.includes("giakhonhanggiacot") ||
        nk.includes("vitrikhon") ||
        nk.includes("vitrigiakhon") ||
        nk.includes("vitrikhuon") ||
        nk.includes("vitrigiakhuon") ||
        nk.includes("moldlocation");

      if (strongPrice || !String(out["Gía khuôn"] || "").trim()) {
        out["Gía khuôn"] = row[k];
      }
    } else {
      out[kind] = row[k];
    }
  });

  return {
    "Item Code": String(out["Item Code"] ?? "").trim(),
    "Item Name": String(out["Item Name"] ?? "").trim(),
    "Process Note": String(out["Process Note"] ?? "").trim(),
    "Gía khuôn": String(out["Gía khuôn"] ?? "").trim(),
    "Trạng thái": String(out["Trạng thái"] ?? "").trim()
  };
}

function importRows(rows, fileName) {
  var cleaned = rows.map(normalizeRow).filter(function(r) {
    return r["Item Code"] || r["Item Name"] || r["Process Note"] || r["Gía khuôn"] || r["Trạng thái"];
  });

  if (!cleaned.length) throw new Error("File không có dữ liệu hợp lệ.");

  var hasCore = cleaned.some(function(r) {
    return r["Item Code"] || r["Item Name"] || r["Process Note"] || r["Gía khuôn"];
  });

  if (!hasCore) throw new Error("Không tìm thấy các cột dữ liệu cần thiết.");

  DATA = cleaned;
  DATA_SOURCE = "import";
  LOCAL_IMPORT_TIME = Date.now();
  document.getElementById("fileName").textContent =
    "Đã import: " + fileName + " — " + DATA.length + " dòng";
  document.getElementById("status").textContent = "✓ Đã nạp " + DATA.length + " dòng";
  document.getElementById("result").className = "empty";
  document.getElementById("result").innerHTML =
    "Dữ liệu đã được import. Nhập Item Code, Item Name hoặc Process Note để tra cứu.";
}

document.getElementById("importBtn").addEventListener("click", function() {
  document.getElementById("fileInput").click();
});

document.getElementById("fileInput").addEventListener("change", function(e) {
  var file = e.target.files && e.target.files[0];
  if (!file) return;
  var reader = new FileReader();
  reader.onload = function(evt) {
    try {
      var data = evt.target.result;
      var wb;
      if (/\.json$/i.test(file.name)) {
        var json = JSON.parse(typeof data === "string" ? data : new TextDecoder().decode(data));
        var rows = Array.isArray(json) ? json : (Array.isArray(json.data) ? json.data : []);
        importRows(rows, file.name);
        return;
      }
      wb = XLSX.read(data, { type: "array", cellDates: false, raw: true });
      var firstSheet = wb.Sheets[wb.SheetNames[0]];
      var matrix = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "", raw: true });
      if (!matrix.length) throw new Error("File Excel không có dữ liệu.");

      // Đọc theo VỊ TRÍ CỘT thay vì phụ thuộc tên key của sheet_to_json.
      // Cách này tránh lỗi khi Excel có ký tự đặc biệt, BOM hoặc tên cột hơi khác.
      var headerRow = matrix[0].map(function(h){ return String(h ?? "").replace(/^\uFEFF/, "").trim(); });
      var headerIndex = {};
      headerRow.forEach(function(h, i){
        var kind = headerKind(h);
        if (kind && headerIndex[kind] === undefined) headerIndex[kind] = i;
      });

      // Nếu không nhận ra tên cột Giá khuôn, tìm đúng cột theo tên gốc phổ biến.
      var priceCol = headerRow.findIndex(function(h){
        return headerKind(h) === "Gía khuôn";
      });
      if (priceCol >= 0) headerIndex["Gía khuôn"] = priceCol;

      var rows = matrix.slice(1).map(function(cols){
        function get(kind){
          var i = headerIndex[kind];
          return i === undefined ? "" : (cols[i] ?? "");
        }
        return {
          "Item Code": get("Item Code"),
          "Item Name": get("Item Name"),
          "Process Note": get("Process Note"),
          "Gía khuôn": get("Gía khuôn"),
          "Trạng thái": get("Trạng thái")
        };
      });
      importRows(rows, file.name);
    } catch (err) {
      document.getElementById("status").textContent = "Lỗi import";
      document.getElementById("result").className = "empty";
      document.getElementById("result").innerHTML = "Không thể import file: " + escapeHtml(err.message || String(err));
    } finally {
      e.target.value = "";
    }
  };
  if (/\.json$/i.test(file.name)) reader.readAsText(file);
  else reader.readAsArrayBuffer(file);
});


const DEFAULT_SHEET_URL = localStorage.getItem("moldSheetUrl") || "";
const AUTO_REFRESH_MS = 15 * 1000;
let onlineUrl = DEFAULT_SHEET_URL;
if (onlineUrl) document.getElementById("sheetUrl").value = onlineUrl;

function csvParse(text) {
  const rows = [];
  let row = [], cell = "", quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i], next = text[i + 1];
    if (ch === '"') {
      if (quoted && next === '"') { cell += '"'; i++; }
      else quoted = !quoted;
    } else if (ch === ',' && !quoted) {
      row.push(cell); cell = "";
    } else if ((ch === '\n' || ch === '\r') && !quoted) {
      if (ch === '\r' && next === '\n') i++;
      row.push(cell); cell = "";
      if (row.some(v => String(v).trim() !== "")) rows.push(row);
      row = [];
    } else cell += ch;
  }
  if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
  if (!rows.length) return [];

  const headers = rows[0].map(h => String(h ?? "").replace(/^\uFEFF/, "").trim());
  const idx = {};
  headers.forEach((h, i) => {
    const kind = headerKind(h);
    if (kind && idx[kind] === undefined) idx[kind] = i;
  });
  const hasPriceColumn = idx["Gía khuôn"] !== undefined;
  const data = rows.slice(1).map(cols => ({
    "Item Code": idx["Item Code"] === undefined ? "" : (cols[idx["Item Code"]] ?? ""),
    "Item Name": idx["Item Name"] === undefined ? "" : (cols[idx["Item Name"]] ?? ""),
    "Process Note": idx["Process Note"] === undefined ? "" : (cols[idx["Process Note"]] ?? ""),
    "Gía khuôn": idx["Gía khuôn"] === undefined ? "" : (cols[idx["Gía khuôn"]] ?? ""),
    "Trạng thái": idx["Trạng thái"] === undefined ? "" : (cols[idx["Trạng thái"]] ?? "")
  })).map(normalizeRow).filter(r =>
    r["Item Code"] || r["Item Name"] || r["Process Note"] || r["Gía khuôn"] || r["Trạng thái"]
  );
  data._hasPriceColumn = hasPriceColumn;
  return data;
}
function setSyncText(text, ok=false) {
  const el = document.getElementById("syncText");
  el.textContent = text;
  el.style.color = ok ? "#20833f" : "#68727d";
}

function toCsvUrl(input) {
  var url = String(input || "").trim();
  if (!url) return "";
  // Accept a Google Sheets edit URL and convert it to a published CSV URL.
  var m = url.match(/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (m) {
    var gidMatch = url.match(/[?#&]gid=(\d+)/);
    var gid = gidMatch ? gidMatch[1] : "0";
    return "https://docs.google.com/spreadsheets/d/" + m[1] + "/export?format=csv&gid=" + gid;
  }
  return url;
}

async function refreshOnline(showError=true) {
  const rawUrl = document.getElementById("sheetUrl").value.trim();
  const url = toCsvUrl(rawUrl);
  if (!url) {
    setSyncText("Chưa nhập Google Sheets — đang dùng dữ liệu có sẵn.");
    return false;
  }
  const status = document.getElementById("status");
  status.textContent = "⟳ Đang đồng bộ...";
  try {
    const response = await fetch(url + (url.indexOf("?") >= 0 ? "&" : "?") + "_ts=" + Date.now(), { cache: "no-store" });
    if (!response.ok) throw new Error("HTTP " + response.status);
    const text = await response.text();
    const rows = csvParse(text);
    const hasPriceColumn = rows._hasPriceColumn === true;
    if (!rows.length) throw new Error("Google Sheets không có dữ liệu");
    if (!hasPriceColumn) {
      setSyncText("⚠ Google Sheets chưa có cột Giá khuôn — giữ dữ liệu hiện tại.");
      status.textContent = "⚠ Thiếu cột Giá khuôn";
      return false;
    }
    DATA = rows;
    DATA_SOURCE = "online";
    const now = new Date();
    const time = now.toLocaleString("vi-VN", { hour: "2-digit", minute: "2-digit", second: "2-digit", day: "2-digit", month: "2-digit", year: "numeric" });
    setSyncText("✓ Đã đồng bộ " + rows.length + " dòng • " + time, true);
    status.textContent = "✓ Đang dùng dữ liệu online";
    doSearch();
    return true;
  } catch (err) {
    status.textContent = "⚠️ Chưa đồng bộ được";
    setSyncText("Không lấy được Google Sheets. Máy này đang dùng dữ liệu gần nhất.");
    if (showError) alert("Không thể kết nối Google Sheets.\n\nHãy kiểm tra Sheet đã được chia sẻ/công khai để web có thể đọc dữ liệu và đúng tên các cột: Item Code, Process Note, Gía khuôn (hoặc Vị trí khuôn/Vị trí giá khuôn), Trạng thái.");
    return false;
  }
}

document.getElementById("saveUrlBtn").addEventListener("click", async function() {
  onlineUrl = document.getElementById("sheetUrl").value.trim();
  localStorage.setItem("moldSheetUrl", onlineUrl);
  if (onlineUrl) await refreshOnline(true);
  else setSyncText("Đã xóa link online — đang dùng dữ liệu có sẵn.");
});

document.getElementById("refreshBtn").addEventListener("click", function() {
  refreshOnline(true);
});

setInterval(function() {
  if (document.getElementById("sheetUrl").value.trim()) refreshOnline(false);
}, AUTO_REFRESH_MS);

if (onlineUrl) refreshOnline(false);

function statusHtml(value) {
  var v = String(value || "").trim().toLowerCase();
  if (v === "ok") return '<div class="status-badge status-ok"><span class="status-dot">✓</span> OK</div>';
  if (v === "sửa chữa" || v === "sua chua" || v === "repair") {
    return '<div class="status-badge status-repair"><span class="status-dot">⚠</span> Sửa chữa</div>';
  }
  return '<div class="status-badge"><span class="status-dot">•</span> Chưa cập nhật</div>';
}

function doSearch() {
  var input = document.getElementById("q");
  var result = document.getElementById("result");
  var status = document.getElementById("status");
  var q = normSearch(input.value);

  if (!q) {
    result.className = "empty";
    result.innerHTML = "Nhập Item Code, Item Name hoặc Process Note để tra giá.";
    status.textContent = "✓ Sẵn sàng tra cứu";
    return;
  }

  var rows = DATA.filter(function(r) {
    return normSearch(r["Item Code"]).indexOf(q) !== -1 ||
           normSearch(r["Item Name"]).indexOf(q) !== -1 ||
           normSearch(r["Process Note"]).indexOf(q) !== -1;
  });

  if (!rows.length) {
    result.className = "empty";
    result.innerHTML = "Không tìm thấy Item Code hoặc Process Note phù hợp.";
    status.textContent = "0 kết quả";
    return;
  }

  status.textContent = rows.length + " kết quả";
  result.className = "";
  result.innerHTML = '<div class="count">Tìm thấy ' + rows.length + ' kết quả</div>' +
    rows.map(function(r) {
      var st = String(r["Trạng thái"] || "").trim().toLowerCase();
      var isRepair = st.includes("sửa") || st.includes("sua") || st.includes("repair") || st.includes("sửa chữa") || st.includes("sua chua");
      var isOk = !isRepair && (st === "ok" || st === "✓" || st === "đã ok" || st === "da ok" || st === "hoàn thành" || st === "hoan thanh");
      return '<div class="card">' +
        '<div class="label">Item Code</div>' +
        '<div class="value">' + escapeHtml(r["Item Code"]) + '</div>' +
        '<div class="label">Item Name</div>' +
        '<div class="value">' + escapeHtml(r["Item Name"]) + '</div>' +
        '<div class="label">Process Note</div>' +
        '<div class="value">' + escapeHtml(r["Process Note"]) + '</div>' +
        '<div class="label">Giá khuôn</div>' +
        '<div class="price">' + escapeHtml(r["Gía khuôn"]) + '</div>' +
        '<div class="label">Trạng thái</div>' +
        statusHtml(r["Trạng thái"]) +
      '</div>';
    }).join("");
}

document.getElementById("searchForm").addEventListener("submit", function(e) {
  e.preventDefault();
  doSearch();
});

document.getElementById("clearBtn").addEventListener("click", function() {
  document.getElementById("q").value = "";
  document.getElementById("result").className = "empty";
  document.getElementById("result").innerHTML = "Nhập Item Code, Item Name hoặc Process Note để tra giá khuôn.";
  document.getElementById("status").textContent = "✓ Sẵn sàng tra cứu";
  document.getElementById("q").focus();
});

// Tab switching
(function(){
  const tabSearch = document.getElementById("tabSearch");
  const tabFunctions = document.getElementById("tabFunctions");
  const searchSection = document.getElementById("searchSection");
  const functionsSection = document.getElementById("functionsSection");

  function showTab(name){
    const isSearch = name === "search";
    tabSearch.classList.toggle("active", isSearch);
    tabFunctions.classList.toggle("active", !isSearch);
    tabSearch.setAttribute("aria-selected", String(isSearch));
    tabFunctions.setAttribute("aria-selected", String(!isSearch));
    searchSection.classList.toggle("hidden", !isSearch);
    functionsSection.classList.toggle("hidden", isSearch);
  }

  tabSearch.addEventListener("click", () => showTab("search"));
  tabFunctions.addEventListener("click", () => showTab("functions"));
})();

// Google Sheets central sync
(function(){
  const urlInput = document.getElementById("appsScriptUrl");
  const saveBtn = document.getElementById("saveAppsScriptBtn");
  const uploadBtn = document.getElementById("uploadToSheetBtn");
  if(!urlInput || !saveBtn || !uploadBtn) return;

  const saved = localStorage.getItem("moldAppsScriptUrl") || "";
  urlInput.value = saved;

  saveBtn.addEventListener("click", function(){
    const u = urlInput.value.trim();
    if(!u) { alert("Vui lòng nhập Web App URL."); return; }
    localStorage.setItem("moldAppsScriptUrl", u);
    alert("Đã lưu Web App URL trên máy này.");
    refreshFromCentralSheet();
  });

  function setSync(msg){
    const el = document.getElementById("syncText");
    if(el) el.textContent = msg;
  }

  async function refreshFromCentralSheet(){
    const u = (localStorage.getItem("moldAppsScriptUrl") || urlInput.value || "").trim();
    if(!u) return;
    try{
      setSync("Đang cập nhật dữ liệu chung...");
      const res = await fetch(u, {cache:"no-store"});
      if(!res.ok) throw new Error("HTTP " + res.status);
      const payload = await res.json();
      if(!Array.isArray(payload.data)) throw new Error("Dữ liệu trả về không đúng định dạng.");
      if(payload.data.length){
        DATA.length = 0;
        payload.data.forEach(x => DATA.push(x));
        setSync("✓ Đã cập nhật dữ liệu chung lúc " + new Date().toLocaleTimeString("vi-VN"));
      } else {
        setSync("Google Sheets đang chưa có dữ liệu.");
      }
    }catch(err){
      setSync("Không cập nhật được Google Sheets: " + err.message);
    }
  }

  uploadBtn.addEventListener("click", async function(){
    const u = (localStorage.getItem("moldAppsScriptUrl") || urlInput.value || "").trim();
    if(!u) { alert("Hãy nhập và lưu Web App URL trước."); return; }
    if(!Array.isArray(DATA) || !DATA.length) { alert("Chưa có dữ liệu để gửi."); return; }

    const old = uploadBtn.textContent;
    uploadBtn.disabled = true;
    uploadBtn.textContent = "⏳ Đang gửi...";
    setSync("Đang gửi dữ liệu lên Google Sheets...");

    try{
      const res = await fetch(u, {
        method:"POST",
        headers:{"Content-Type":"text/plain;charset=utf-8"},
        body:JSON.stringify({data:DATA})
      });
      const text = await res.text();
      let payload = {};
      try { payload = JSON.parse(text); } catch(e) {}
      if(!res.ok || payload.ok === false) throw new Error(payload.error || ("HTTP " + res.status));
      setSync("✓ Đã đẩy " + DATA.length + " dòng lên Google Sheets lúc " + new Date().toLocaleTimeString("vi-VN"));
      alert("Đã đẩy dữ liệu lên Google Sheets thành công.");
    }catch(err){
      setSync("Lỗi đồng bộ: " + err.message);
      alert("Không thể đẩy dữ liệu lên Google Sheets.\n\n" + err.message);
    }finally{
      uploadBtn.disabled = false;
      uploadBtn.textContent = old;
    }
  });

  // Try central data once when the page opens, then every 15 seconds.
  refreshFromCentralSheet();
  setInterval(refreshFromCentralSheet, 15000);
})();
</script>
</body>
</html>
