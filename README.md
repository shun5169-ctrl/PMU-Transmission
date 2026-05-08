
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Malaysian Substation Map — Excel Import</title>

  <!-- Icons -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">

  <!-- Fonts -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,300..700;1,9..40,300..700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">

  <!-- XLSX (SheetJS) for Excel/CSV parsing -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

  <!-- JSZip for KMZ (zipped KML) import/export -->
  <script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>

  <style>
    :root{
      /* Core brand — richer greens */
      --brand-deep:  #0d4a1a;
      --brand-main:  #1a7a2e;
      --brand-accent:#4ade80;
      --brand-muted: #bbf7d0;
    
      /* UI neutrals — warmer tones */
      --ink:   #1a1a2e;
      --muted: #6b7280;
      --bg:    #f0f4f0;
      --card:  #ffffff;
      --line:  #e2e8f0;
    
      /* Alerts / toasts */
      --danger-deep:#b91c1c;
      --danger-soft:#fecaca;

      /* Toast success */
      --success-bg: #ecfdf5;
      --success-border: #6ee7b7;
    
      /* Map elements */
      --pmu: #D81B60;
      --route: #0DFF7A;   /* driving-route line colour */
    
      /* Transmission line colours */
      --kv-500: #C97125;
      --kv-275: #00A8FF;
      --kv-132: #FF1744;
      --kv-33:  #FF9800;
      --kv-11:  #9C27B0;

      /* Radius system */
      --radius-sm: 8px;
      --radius-md: 12px;
      --radius-lg: 16px;

      /* Shadows */
      --shadow-sm: 0 1px 3px rgba(0,0,0,.06), 0 1px 2px rgba(0,0,0,.04);
      --shadow-md: 0 4px 12px rgba(0,0,0,.07), 0 2px 4px rgba(0,0,0,.04);
      --shadow-lg: 0 10px 30px rgba(0,0,0,.1), 0 4px 8px rgba(0,0,0,.04);
    }

    *{box-sizing:border-box;}

    body{
      background:var(--bg);
      color:var(--ink);
      line-height:1.6;
      font-family:'DM Sans', system-ui, -apple-system, sans-serif;
      -webkit-font-smoothing: antialiased;
    }
    header{
      background: linear-gradient(135deg, #0d4a1a 0%, #1a7a2e 50%, #15803d 100%);
      color:#fff;
      padding: 5px 16px;
      border-bottom: 1px solid rgba(255,255,255,.1);
    }
    .logo{ color: var(--brand-muted); }
    .header-content{ display:flex; justify-content:space-between; align-items:center; }
    header h1{
      font-size: 1.1rem;
      margin-bottom: 0;
      line-height: 1.2;
    }
    header p{
      opacity:.9;
      font-size:.75rem;
      line-height:1.2;
      margin:0;
    }
    .logo{ font-size:1.3rem; color:#a5d6a7; }


    .filter-section{
      margin-bottom:0.6rem;
      background:#fff;
      border-radius:var(--radius-md);
      padding:1rem 1rem;
      box-shadow:var(--shadow-sm);
      font-size:0.9rem;
      border:1px solid var(--line);
      transition: box-shadow .2s;
    }
    .filter-section:hover{
      box-shadow:var(--shadow-md);
    }
    .filter-section h3{
      margin-bottom:.9rem;
      color:var(--brand-deep);
      font-size:0.95rem;
      font-weight:600;
      display:flex;
      align-items:center;
      gap:8px;
      letter-spacing:-0.01em;
    }
    .filter-section h3 i{
      color:var(--brand-main);
      font-size:0.9em;
    }

    .state-filter{display:grid;grid-template-columns:1fr 1fr;gap:.6rem}
    .state-checkbox{display:flex;align-items:center;background:#fafffe;padding:.5rem .7rem;border-radius:var(--radius-sm);border:1px solid var(--line);transition:all .2s ease}
    .state-checkbox:hover{background:#ecfdf5;transform:translateY(-1px);box-shadow:var(--shadow-sm);border-color:var(--brand-muted)}
    .state-checkbox input{margin-right:.7rem;accent-color:var(--brand-main);width:16px;height:16px}
    .filter-actions{display:flex;gap:10px;margin-top:.7rem}
    .filter-btn{
      flex:1;
      padding:.55rem;
      border:0;
      border-radius:var(--radius-sm);
      background: linear-gradient(135deg, var(--brand-deep), var(--brand-main));
      color:#fff;
      font-weight:600;
      cursor:pointer;
      transition:all .2s ease;
      display:flex;
      align-items:center;
      justify-content:center;
      gap:6px;
      font-size:0.85rem;
      font-family:inherit;
      letter-spacing:0.01em;
    }
    .filter-btn:hover{
      background: linear-gradient(135deg, #092e10, var(--brand-deep));
      transform:translateY(-1px);
      box-shadow: 0 4px 12px rgba(13,74,26,.25);
    }
    .filter-btn:active{
      transform:translateY(0);
    }
    .filter-btn.secondary{
      background: var(--brand-deep);
      opacity:0.9;
    }
    .filter-btn.secondary:hover{
      background: #092e10;
      opacity:1;
    }

    .search-box{margin-bottom:1rem}
    .search-box input{width:100%;padding:.7rem 1rem;border:1px solid var(--line);border-radius:var(--radius-sm);background:#fafffe;font-size:.9rem;transition:all .2s;font-family:inherit}
    .search-box input:focus{outline:none;border-color:var(--brand-accent);box-shadow:0 0 0 3px rgba(74,222,128,.15)}

    .uploader{border:2px dashed var(--brand-muted);background:#fafffe;border-radius:var(--radius-md);padding:.9rem;display:flex;flex-direction:column;gap:.6rem;align-items:flex-start;transition:all .2s}
    .uploader.drag{background:#ecfdf5;border-color:var(--brand-accent)}
    .uploader small{color:var(--muted)}
    .uploader .row{display:flex;gap:.6rem;width:100%;align-items:center}
    .uploader input[type=file]{flex:1}
    .import-status{font-size:.85rem;color:#2e7d32}

    .substation-list{
      max-height:380px;
      overflow-y:auto;
      border:1px solid var(--line);
      border-radius:var(--radius-sm);
      background:#fff;
      padding:.35rem;
      scrollbar-width:thin;
      scrollbar-color:var(--brand-muted) transparent;
    }
    .substation-item{
      padding:.55rem .6rem;
      border-bottom:1px solid #f8fafc;
      cursor:pointer;
      border-radius:var(--radius-sm);
      margin-bottom:2px;
      transition:all .15s ease;
    }
    .substation-item:hover{
      background:#ecfdf5;
      transform:translateX(2px);
    }
    .substation-name{
      font-weight:600;
      color:var(--brand-deep);
      margin-bottom:2px;
      font-size:.85rem;
      line-height:1.25;
    }
    .substation-location{
      font-size:.75rem;
      color:var(--muted);
      display:flex;
      align-items:center;
      gap:4px;
      line-height:1.2;
    }

    /* (stats moved to header — see .header-stat rules) */

    #msm-app{
      width: min(100vw, 1600px);
      margin: 0 auto;
      padding-left: 12px;
      padding-right: 12px;
    }
    @media (min-width: 1440px){
      #msm-app{ width:100vw; margin:0; padding-left:8px; padding-right:8px; }
    }
    .full-bleed{
      width: 100vw;
      position: relative;
      left: 50%;
      right: 50%;
      margin-left: -50vw;
      margin-right: -50vw;
    }
    .content{
  display:flex;
  gap:12px;
  overflow:hidden;

  /* KEY: make the sidebar + map row fill the screen */
  height: calc(100vh - 42px);
}

    .sidebar{ width:350px; background:#fff; padding:1.2rem 1.5rem; overflow-y:auto; box-shadow:var(--shadow-lg); z-index:500; border-right:1px solid var(--line) }

.map-container{
  flex:1;
  position:relative;
  min-width:720px;
  height: 100%;
}
#map{
  height: 100%;
  width: 100%;
}



    .toast{position:fixed;right:16px;bottom:16px;max-width:360px;background:var(--success-bg);border:1px solid var(--success-border);color:var(--brand-deep);padding:.75rem 1rem;border-radius:var(--radius-md);box-shadow:var(--shadow-lg);display:none;z-index:1200;font-size:.9rem;font-weight:500;backdrop-filter:blur(8px);animation:toast-slide .3s ease}
    .toast.show{display:flex;align-items:center;gap:8px}
    .toast::before{content:'\f058';font-family:'Font Awesome 6 Free';font-weight:900;color:var(--brand-accent);font-size:1.1em}
    @keyframes toast-slide{from{transform:translateY(12px);opacity:0}to{transform:translateY(0);opacity:1}}

    @media(max-width:768px){
      .content{flex-direction:column}
      .state-filter{grid-template-columns:1fr}
    }

    @media (max-width: 768px){
      .layers-panel.toolbar{
        top: auto;
        bottom: 16px;
        right: 16px;
        left: 16px;
        justify-content: flex-start;
      }
      .layers-panel.toolbar #osmStatus{
        flex-basis: 100%;
        margin-left: 0;
        margin-top: 4px;
      }
    }

    .layers-panel.toolbar.in-header h3{
      color: var(--brand-deep);
    }
    .layers-panel.toolbar.in-header .state-checkbox label{
      color: var(--ink);
    }
    .layers-panel.toolbar.in-header .state-checkbox small{
      color: var(--muted);
    }
    .layers-panel.toolbar.in-header .filter-btn.secondary{
      color: #fff;
    }
    .layers-panel.toolbar.in-header .filter-btn{
      color: #fff;
    }
    .layers-panel.toolbar.in-header{
      background: #ffffff;
      border-color: #e7e7e7;
    }


    /* Hamburger in header */
    .hamburger{
      background: linear-gradient(135deg, var(--brand-deep), var(--brand-main)); color:#fff; border:0; border-radius:var(--radius-sm);
      padding:.5rem .7rem; font-size:1.05rem; display:none; gap:.5rem;
      box-shadow:var(--shadow-sm);
      transition:all .2s;
    }
    .hamburger:hover{ box-shadow:var(--shadow-md); }
    .hamburger i{ font-size:1.1rem }

    /* Drawer close button (inside sidebar) */
    .sidebar-close{
      display:none; position:sticky; top:0; margin:-6px -6px 10px auto;
      background:#f8fafc; border:1px solid var(--line); border-radius:var(--radius-sm); padding:.4rem .55rem;
      cursor:pointer; transition:all .15s; z-index:10;
    }
    .sidebar-close:hover{ background:#ecfdf5; border-color:var(--brand-muted); }

    .backdrop{
      position:fixed; inset:0; background:rgba(0,0,0,.35); z-index:1099; display:none;
      backdrop-filter:blur(4px);
      -webkit-backdrop-filter:blur(4px);
    }

    @media (max-width: 768px){
      .content{ flex-direction:column; gap:0 }
      .map-container{ min-width:0; height: calc(100vh - 42px); }
      .hamburger{ display:inline-flex; }
      .sidebar{
        width: min(88vw, 400px);
        max-width: 92vw;
        position: fixed;
        z-index: 1100;
        top: 0;
        right: 0;
        height: 100vh;
        height: 100dvh;
        overflow-y: auto;
        transform: translateX(100%);
        transition: transform .28s cubic-bezier(.4,0,.2,1);
        box-shadow: -8px 0 24px rgba(0,0,0,.18);
        border-left: 1px solid var(--line);
      }
      .sidebar.open{ transform: translateX(0); }
      .sidebar-close{ display:inline-flex; }
      .backdrop.show{ display:block; }
      .layers-panel.toolbar{
        position: fixed;
        bottom: 14px; left: 14px; right: 14px; top: auto;
        padding:10px; border-radius:12px;
        box-shadow:0 10px 24px rgba(0,0,0,.18);
      }
      .substation-item{ padding:.7rem }
      .search-box input{ padding:.6rem .8rem }
      .state-checkbox{ padding:.45rem .6rem }
    }

    .gm-style-iw-d { font-size: 13px; line-height: 1.45; }
    .gm-style-iw-d h3 { font-size: 16px; margin-bottom: 6px; }

    @media (max-width:768px){
      .state-checkbox input{ width:22px; height:22px; }
      .filter-btn{ padding:.7rem 1rem; font-size:14px; }
    }

    .map-search{
      display:flex; align-items:center; gap:8px;
      background:#fff; border:1px solid var(--line); border-radius:var(--radius-md);
      padding:6px 12px; max-width:420px;
      box-shadow:var(--shadow-sm);
      transition:all .2s;
    }
    .map-search:focus-within{
      border-color:var(--brand-accent);
      box-shadow:0 0 0 3px rgba(74,222,128,.12);
    }
    .map-search i{ color:var(--muted); }
    .map-search input{
      border:0; outline:none; flex:1; font-size:13px; font-family:inherit;
    }
    .map-search button{
      background:linear-gradient(135deg,var(--brand-deep),var(--brand-main)); color:#fff; border:0; border-radius:var(--radius-sm); padding:6px 10px; cursor:pointer; transition:all .2s;
    }
    .map-search button:hover{
      box-shadow:0 2px 8px rgba(13,74,26,.3);
    }
    @media (max-width:768px){
      .map-search{ width:100%; margin-top:8px; }
    }

    .as-control{
      background:rgba(255,255,255,.96);
      border:1px solid var(--line);
      border-radius:var(--radius-lg);
      box-shadow:var(--shadow-lg);
      padding:12px;
      margin:10px;
      max-width: 420px;
      backdrop-filter:blur(12px);
    }


    .gm-style .gm-style-iw-c{
      padding: 6px 10px 8px 10px !important;
      border-radius: var(--radius-md) !important;
      box-shadow: var(--shadow-lg) !important;
    }
    .gm-ui-hover-effect{
      display: none !important;
    }
    .iw-compact{ position:relative; line-height:1.25; }
    .iw-compact h3{ font-size:14px; margin:0 0 4px 0; color:var(--brand-deep); font-weight:600; }
    .iw-compact p{ font-size:12px; margin:0 0 2px 0; color:var(--ink); }
    .iw-close{
      position:absolute; top:4px; right:6px;
      padding:0; border:0; background:transparent; cursor:pointer;
      font-size:16px; line-height:1; color:var(--muted);
      transition:color .15s;
    }
    .iw-close:hover{ color:var(--ink); }

    .boundary-controls .row{ display:flex; gap:8px; flex-wrap:wrap; }
    .boundary-controls input[type="text"]{ width:100%; padding:8px 10px; border:1px solid #e0e0e0; border-radius:8px; font-size:13px; }
    .boundary-controls .filter-btn{ padding:.5rem .7rem; }

    .as-control.compact{
      padding: 8px 10px;
      max-width: 360px;
      border-radius: 12px;
    }
    .boundary-controls h3{
      margin: 0 0 6px 0;
      font-size: 16px;
    }
    .boundary-controls input[type="text"]{
      height: 38px;
      font-size: 13px;
    }
    .boundary-controls .row{ gap:6px }
    .boundary-controls .filter-btn{
      padding:.45rem .6rem;
      font-size:14px;
      border-radius:10px;
    }
    .boundary-controls small{ font-size:12px }

    .filter-section.boundary .row{ display:flex; gap:8px; flex-wrap:wrap; }
    .filter-section.boundary input[type="text"]{
      width:100%; padding:8px 10px; border:1px solid #e0e0e0; border-radius:8px; font-size:13px;
    }
    .filter-section.boundary .filter-btn{ padding:.45rem .6rem; font-size:14px; border-radius:10px; }
    .filter-section.boundary small{ font-size:12px }

    .min-btn{
      position:absolute;
      right:8px;
      top:6px;
      width:24px; height:24px;
      display:flex;
      align-items:center;
      justify-content:center;
      line-height:1;
      border:1px solid var(--line); border-radius:6px;
      background:#f8fafc;
      cursor:pointer; font-weight:700; font-size:13px;
      box-shadow:var(--shadow-sm);
      z-index:2;
      color:var(--muted);
      transition:all .15s;
    }
    .min-btn:hover{ background:#ecfdf5; color:var(--brand-deep); border-color:var(--brand-muted); }
    .min-btn:active{ transform:translateY(1px); }

    .card-header,
    #layersPanel h3{
      margin:0 0 8px 0;
      line-height:1.1;
      position:relative;
      padding-right:30px;
    }
    #mapToolsControl .min-btn{ top:6px; }
    #layersPanel .min-btn{ top:6px; }

    #layersPanel h3{
      font-size: 12px;
      line-height: 1.2;
      font-weight: 700;
    }
    #layersPanel h3 i{
      font-size: 1.05em;
    }
    #layersPanel .min-btn{ top: 4px; }
    @media (max-width:768px){
      #layersPanel h3{ font-size: 18px; }
    }
    #layersPanel .state-checkbox label{
      font-size: 0.75rem;
    }
    #layersPanel .state-checkbox small{
      font-size: 0.75rem;
    }
    .collapsible .card-body{ display:block; }
    .collapsible.collapsed .card-body{ display:none; }


    #layersPanel .min-btn, #mapToolsControl .min-btn{ top:4px; }

    .state-dropdown {
      border: 1px solid var(--line);
      border-radius: var(--radius-sm);
      background: #fff;
      overflow: hidden;
      transition: box-shadow .2s;
    }
    .state-dropdown:hover {
      box-shadow: var(--shadow-sm);
    }
    .state-dropdown .dropdown-header {
      padding: 10px 14px;
      display: flex;
      justify-content: space-between;
      cursor: pointer;
      font-weight: 600;
      font-size: .85rem;
      color: var(--ink);
      transition: background .15s;
    }
    .state-dropdown .dropdown-header:hover {
      background: #f8fafc;
    }
    .state-dropdown .dropdown-body {
      max-height: 0;
      overflow: hidden;
      transition: max-height .25s ease;
      border-top: 1px solid var(--line);
    }
    .state-dropdown.open .dropdown-body {
      max-height: 260px;
      overflow-y: auto;
    }
    .state-dropdown .dropdown-body label {
      display: flex;
      align-items: center;
      padding: 8px 14px;
      border-bottom: 1px solid #f8fafc;
      font-size: .85rem;
      transition: background .1s;
    }
    .state-dropdown .dropdown-body label:hover {
      background: #ecfdf5;
    }
    .state-dropdown input[type="checkbox"] {
      margin-right: 10px;
      transform: scale(1.15);
      accent-color: var(--brand-main);
    }

    .filter-section.import-section{
      display:none !important;
    }


    .sidebar .filter-section{
      margin-bottom: 0.3rem;
      padding-bottom: 0.6rem;
    }
    .sidebar .filter-section + .filter-section{
      margin-top: 0.1rem !important;
    }
    

    button.mini-btn:hover{
      background:#e0e0e0;
    }



/* --- Sidebar layout clean-up --- */
.msm-panel-header {
  display: flex;
  align-items: center;
  gap: 6px;
  margin-bottom: 8px;
}

.msm-panel-header h3 {
  margin: 0;
}



/* Saved Lands Summary card */
.msm-summary-wrapper {
  margin: 0 -12px;
  width: calc(100% + 24px);
  padding: 10px 0 16px;
}
@media (min-width: 1440px){
  .msm-summary-wrapper {
    margin: 0 -8px;
    width: calc(100% + 16px);
  }
}

.msm-summary-wrapper > .filter-section.boundary {
  border-radius: 0;
  border-left: 0;
  border-right: 0;
  margin: 0;
  padding: 12px 16px;
}

.msm-summary-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 8px;
}

.msm-summary-meta {
  font-size: 11px;
  color: var(--muted);
}

/* Table tidy-up */
.msm-summary-table-wrapper {
  max-height: 180px;
  overflow: auto;
  border: 1px solid var(--line);
  border-radius: var(--radius-sm);
  background: #ffffff;
  scrollbar-width: thin;
  scrollbar-color: var(--brand-muted) transparent;
}

.msm-summary-table {
  width: 100%;
  border-collapse: separate;
  border-spacing: 0;
  font-size: 12px;
  font-family: 'DM Sans', system-ui, sans-serif;
}

.msm-summary-table thead th {
  position: sticky;
  top: 0;
  z-index: 2;
  background: linear-gradient(135deg, #f0fdf4, #ecfdf5);
  font-weight: 600;
  padding: 8px 8px;
  border-bottom: none;
  box-shadow: inset 0 -2px 0 #d1fae5;  /* replaces border-bottom; works with border-separate */
  text-align: left;
  white-space: nowrap;
  color: var(--brand-deep);
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: 0.04em;
}

.msm-summary-table tbody td {
  padding: 5px 8px;
  border-bottom: 1px solid #f1f5f9;
}

.msm-summary-table tbody tr:nth-child(odd) {
  background: #fafffe;
}
.msm-summary-table tbody tr:hover {
  background: #ecfdf5;
  cursor: pointer;
}

/* Inputs in table */
.msm-summary-table input[type="text"] {
  width: 100%;
  border-radius: 6px;
  border: 1px solid var(--line);
  padding: 3px 6px;
  font-size: 12px;
  font-family: inherit;
  transition: all .15s;
}
.msm-summary-table input[type="text"]:focus {
  outline: none;
  border-color: var(--brand-accent);
  box-shadow: 0 0 0 2px rgba(74,222,128,.12);
}

/* Compact Export to Excel button */
.msm-export-btn {
  flex: 0 0 auto;
  padding: 6px 14px;
  font-size: 11px;
  border-radius: 999px;
  background: linear-gradient(135deg, var(--brand-deep), var(--brand-main)) !important;
  box-shadow: var(--shadow-sm);
  transition: all .2s;
}
.msm-export-btn:hover {
  box-shadow: 0 4px 12px rgba(13,74,26,.25);
  transform: translateY(-1px);
}

.sidebar-section-title{
  margin:14px 0 6px;
  font-size:13px;
  font-weight:700;
  color:var(--brand-deep);
  display:flex;
  align-items:center;
  gap:6px;
  letter-spacing:-0.01em;
}

.sidebar-section-title i{
  color:var(--brand-main);
}

.msm-card{
  background:#FFFFFF;
  border-radius:var(--radius-lg);
  padding:14px;
  box-shadow:var(--shadow-md);
  border:1px solid var(--line);
  margin-bottom:12px;
  transition:box-shadow .2s;
}
.msm-card:hover{
  box-shadow:var(--shadow-lg);
}

.msm-card h4{
  margin-bottom:8px;
  font-size:13px;
  font-weight:600;
  color:var(--ink);
}

.msm-note{
  font-size:11px;
  color:var(--muted);
}

.compact-row{
  display:flex;
  gap:8px;
  flex-wrap:wrap;
}

.boundary-fields{
  display:flex;
  flex-direction:column;
  gap:6px;
  margin-bottom:6px;
}

.boundary-field label{
  display:block;
  font-size:11px;
  font-weight:600;
  color:var(--brand-deep);
  margin-bottom:2px;
  letter-spacing:0.02em;
}

.boundary-field input{
  width:100%;
  padding:5px 8px;
  border-radius:6px;
  border:1px solid var(--line);
  font-size:12px;
  font-family:inherit;
  transition:all .15s;
}
.boundary-field input:focus{
  outline:none;
  border-color:var(--brand-accent);
  box-shadow:0 0 0 2px rgba(74,222,128,.12);
}

.msm-boundary-status{
  margin:4px 0 6px;
  font-size:11px;
  color:var(--muted);
}

/* 3×2 grid for buttons */
.msm-boundary-controls{
  display:grid;
  grid-template-columns:repeat(3,minmax(0,1fr));
  gap:6px;
  margin-top:4px;
}

.msm-boundary-controls .filter-btn{
  width:100%;
}

/* Small pill-style delete button */
button.mini-btn{
  font-size:11px;
  padding:3px 8px;
  border-radius:999px;
  border:1px solid var(--line);
  background:#f8fafc;
  cursor:pointer;
  transition:all .15s;
  font-family:inherit;
  color:var(--muted);
}
button.mini-btn:hover{
  background:#fef2f2;
  border-color:#fecaca;
  color:#b91c1c;
}

/* ===== Compact layout for the whole left sidebar ===== */

/* Slightly tighter padding on the whole sidebar */
.sidebar{
  padding:0.8rem 1rem !important;   /* was 1.2rem 1.5rem */
}

/* Less vertical gap between blocks in the sidebar */
.sidebar > * + *{
  margin-top:6px !important;        /* was 14px via earlier rule */
}

/* Make all cards (filter-section) denser */
.sidebar .filter-section{
  padding:0.6rem 0.7rem !important; /* reduce inner padding */
  margin-bottom:0.3rem !important;
}

/* Titles closer to content and slightly smaller */
.sidebar .filter-section h3{
  margin-bottom:0.4rem !important;
  font-size:0.95rem !important;
}

/* Inputs smaller across the sidebar */
.sidebar input[type="text"],
.sidebar input[type="number"],
.sidebar select,
.sidebar textarea{
  padding:3px 6px !important;
  font-size:0.85rem !important;
}

/* Buttons (including the green ones) more compact */
.sidebar .filter-btn,
.sidebar button{
  padding:0.35rem 0.7rem !important;
  font-size:0.85rem !important;
}

/* Grids / rows in cards use smaller gaps */
.sidebar .state-filter,
.sidebar .msm-boundary-controls,
.sidebar .route-buttons{
  gap:4px !important;
}

/* Filter row in Saved Lands table */
/* Small pill-style selects in header */
.msm-summary-filter-select {
  width: 100%;
  font-size: 11px;
  padding: 3px 6px;
  border-radius: 999px;
  border: 1px solid var(--line);
  background: #fafffe;
  box-sizing: border-box;
  font-family: inherit;
  transition: all .15s;
  cursor: pointer;
}
.msm-summary-filter-select:focus {
  outline: none;
  border-color: var(--brand-accent);
  box-shadow: 0 0 0 2px rgba(74,222,128,.12);
}
.msm-summary-filter-select:hover {
  border-color: var(--brand-muted);
}

/* (mobile map-container height consolidated in main @media block above) */

/* ===== Global vertical compaction ===== */

/* Reduce top & bottom body spacing */
body{
  margin:0;
}

/* Reduce vertical padding everywhere */
.filter-section,
.msm-card{
  padding-top:8px !important;
  padding-bottom:8px !important;
}

/* Reduce vertical gaps between sections */
.sidebar > * + *{
  margin-top:4px !important;
}

/* Reduce headers spacing */
h3, h4{
  margin-top:4px !important;
  margin-bottom:6px !important;
}
.sidebar{
  max-height: calc(100vh - 42px);
  overflow-y:auto;
}
html, body{
  height: 100%;
  margin: 0;
}

/* ===============================
   iPhone / iOS Safari Fix
   Ensure bottom buttons are visible
================================== */

/* Use dynamic viewport height instead of 100vh (iOS Safari issue) */
@supports (height: 100dvh) {
  .content {
    height: calc(100dvh - 42px) !important;
  }
  .sidebar {
    max-height: calc(100dvh - 42px) !important;
  }
  @media (max-width:768px){
    .map-container{
      height: calc(100dvh - 42px) !important;
    }
    .sidebar{
      height: 100dvh !important;
      max-height: 100dvh !important;
    }
  }
}

/* Ensure sidebar can scroll fully (so bottom buttons won't be hidden) */
.sidebar {
  padding-bottom: 90px !important;     /* extra space for last buttons */
  overflow-y: auto !important;
  -webkit-overflow-scrolling: touch;   /* smooth scroll on iPhone */
}

/* Safety: avoid panels exceeding sidebar height */
#layersPanel {
  max-height: none !important;
}

/* If the panel is collapsible, ensure content isn't clipped */
#layersPanelBody {
  overflow: visible !important;
}

/* ===== Final design polish ===== */

/* Smooth scrollbar for sidebar */
.sidebar::-webkit-scrollbar { width: 5px; }
.sidebar::-webkit-scrollbar-track { background: transparent; }
.sidebar::-webkit-scrollbar-thumb { background: var(--brand-muted); border-radius: 999px; }
.sidebar::-webkit-scrollbar-thumb:hover { background: var(--brand-accent); }

/* Substation list scrollbar */
.substation-list::-webkit-scrollbar { width: 4px; }
.substation-list::-webkit-scrollbar-track { background: transparent; }
.substation-list::-webkit-scrollbar-thumb { background: var(--line); border-radius: 999px; }

/* Summary table scrollbar */
.msm-summary-table-wrapper::-webkit-scrollbar { width: 4px; height: 4px; }
.msm-summary-table-wrapper::-webkit-scrollbar-track { background: transparent; }
.msm-summary-table-wrapper::-webkit-scrollbar-thumb { background: var(--brand-muted); border-radius: 999px; }

/* Smooth transitions on all interactive inputs */
input, select, textarea, button {
  font-family: inherit;
}

/* Active row highlight in Saved Lands table */
.msm-summary-table tbody tr.active-row,
.msm-summary-table tbody tr[data-shape-id]:focus-within {
  background: #ecfdf5 !important;
  outline: 2px solid var(--brand-accent);
  outline-offset: -2px;
}

/* Better layer checkbox alignment */
#layersPanelBody .state-checkbox {
  padding: .4rem .6rem;
  margin-bottom: 3px;
}
#layersPanelBody .state-checkbox label {
  font-size: 0.8rem;
  font-weight: 500;
}

/* Import status text */
.import-status {
  font-size: .82rem;
  color: var(--brand-main);
  font-weight: 500;
}

/* Route summary text */
#routeSummary, #routeSummaryFS {
  font-family: 'JetBrains Mono', monospace;
  font-size: 12px;
  color: var(--brand-deep);
  font-weight: 500;
}

/* Mobile: nicer sidebar slide */
@media (max-width: 768px) {
  .sidebar {
    box-shadow: -12px 0 40px rgba(0,0,0,.2);
  }
}

/* ===== HEADER REDESIGN ===== */
.header-left{ display:flex; align-items:center; gap:10px; }
.header-brand{ display:flex; align-items:center; gap:10px; }
.logo-icon{ font-size:1.5rem; color:var(--brand-accent); }
.header-right{ display:flex; align-items:center; gap:12px; }

.header-stats{
  display:flex; align-items:center; gap:6px;
  background:rgba(255,255,255,.1);
  border-radius:var(--radius-sm);
  padding:4px 8px;
}
.header-stat{
  display:flex; flex-direction:row; align-items:baseline; gap:3px;
}
.header-stat + .header-stat::before{
  content:'·';
  color:rgba(255,255,255,.35);
  font-size:.8rem;
  margin-right:2px;
}
.header-stat:hover{ background:none; }
.header-stat-value{
  font-size:.82rem;
  font-weight:700;
  color:#fff;
  font-family:'JetBrains Mono', monospace;
}
.header-stat-label{
  font-size:.68rem;
  color:rgba(255,255,255,.65);
  text-transform:uppercase;
  letter-spacing:0.03em;
  font-weight:500;
}

.dark-toggle{
  background:rgba(255,255,255,.12);
  border:1px solid rgba(255,255,255,.2);
  color:#fff;
  border-radius:var(--radius-sm);
  width:36px; height:36px;
  display:flex; align-items:center; justify-content:center;
  cursor:pointer;
  transition:all .2s;
  font-size:1rem;
}
.dark-toggle:hover{ background:rgba(255,255,255,.2); }

/* Drawing Page link button in header */
.header-link-btn{
  display:inline-flex;
  align-items:center;
  gap:6px;
  background:rgba(255,255,255,.12);
  border:1px solid rgba(255,255,255,.2);
  color:#fff;
  text-decoration:none;
  border-radius:var(--radius-sm);
  padding:0 12px;
  height:36px;
  font-size:.82rem;
  font-weight:600;
  letter-spacing:.02em;
  cursor:pointer;
  transition:all .2s;
  white-space:nowrap;
}
.header-link-btn:hover{
  background:rgba(255,255,255,.22);
  border-color:rgba(255,255,255,.35);
  transform:translateY(-1px);
}
.header-link-btn i{ font-size:.9rem; }
.header-link-btn .link-label-short{ display:none; }

@media (max-width:900px){
  .header-link-btn .link-label-full{ display:none; }
  .header-link-btn .link-label-short{ display:inline; }
}
@media (max-width:600px){
  .header-link-btn{ padding:0 10px; height:32px; font-size:.75rem; }
  .header-link-btn .link-label-full,
  .header-link-btn .link-label-short{ display:none; }
}

@media (max-width:768px){
  .header-stats{ display:none; }
  .header-content{ gap:8px; }
  header h1{ font-size:1.1rem; }
  header p{ font-size:.75rem; }
  .header-brand{ gap:8px; }
  .logo-icon{ font-size:1.2rem; }
  .dark-toggle{ width:32px; height:32px; font-size:.85rem; }
}

/* ===== SIDEBAR TABS ===== */
.sidebar-tabs{
  display:flex;
  gap:0;
  background:var(--line);
  border-radius:var(--radius-sm);
  padding:3px;
  margin:8px 0;
}
.sidebar-tab{
  flex:1;
  padding:8px 4px;
  border:0;
  background:transparent;
  border-radius:6px;
  cursor:pointer;
  font-size:.78rem;
  font-weight:600;
  color:var(--muted);
  transition:all .2s;
  display:flex;
  align-items:center;
  justify-content:center;
  gap:5px;
  font-family:inherit;
  white-space:nowrap;
}
.sidebar-tab i{ font-size:.75rem; }
.sidebar-tab:hover{ color:var(--ink); }
.sidebar-tab.active{
  background:#fff;
  color:var(--brand-deep);
  box-shadow:var(--shadow-sm);
}

.tab-panel{ display:none; }
.tab-panel.active{ display:block; }

/* ===== COORDINATE BAR ===== */
.coord-bar{
  position:absolute;
  bottom:8px;
  left:50%;
  transform:translateX(-50%);
  background:rgba(0,0,0,.75);
  color:#fff;
  padding:5px 14px;
  border-radius:999px;
  font-size:11px;
  font-family:'JetBrains Mono', monospace;
  z-index:100;
  pointer-events:none;
  backdrop-filter:blur(8px);
  -webkit-backdrop-filter:blur(8px);
  white-space:nowrap;
  display:flex;
  align-items:center;
  gap:8px;
}
.coord-divider{
  color:rgba(255,255,255,.3);
}
@media (max-width:768px){
  .coord-bar{
    bottom:60px;
    font-size:10px;
    padding:4px 10px;
  }
}

/* ===== MOBILE: Tabs + Summary + Layout fixes ===== */
@media (max-width: 768px){
  /* Tabs: bigger touch targets */
  .sidebar-tabs{
    padding:3px;
    margin:6px 0;
    position:sticky;
    top:0;
    z-index:5;
    background:var(--line);
  }
  .sidebar-tab{
    padding:10px 6px;
    font-size:.8rem;
  }
  .sidebar-tab i{ font-size:.85rem; }

  /* Summary table: horizontal scroll + full bleed */
  .msm-summary-wrapper{
    margin:0;
    width:100%;
    padding:8px 0 12px;
  }
  .msm-summary-wrapper > .filter-section.boundary{
    padding:10px 10px;
    border-radius:0;
    border-left:0;
    border-right:0;
  }
  .msm-summary-table-wrapper{
    -webkit-overflow-scrolling:touch;
    overflow-x:auto;
  }
  .msm-summary-table{
    min-width:700px;
  }
  .msm-summary-header{
    flex-direction:column;
    align-items:flex-start;
    gap:6px;
  }
  .msm-export-btn{
    align-self:flex-end;
  }

  /* Filter section cards */
  .filter-section{
    border-radius:var(--radius-sm) !important;
  }

  /* Search box full width */
  .search-box input{
    font-size:.85rem;
    padding:.6rem .8rem;
  }

  /* Substation items - bigger touch target */
  .substation-item{
    padding:.7rem .6rem;
  }
  .substation-name{
    font-size:.88rem;
  }

  /* Boundary controls grid */
  .msm-boundary-controls{
    grid-template-columns:repeat(2,1fr) !important;
  }

  /* Info window on small screens */
  .gm-style-iw-d{
    font-size:12px;
    max-width:260px !important;
  }
}

/* ===== DARK MODE ===== */
body.dark-mode{
  --bg: #111827;
  --card: #1f2937;
  --ink: #f1f5f9;
  --muted: #94a3b8;
  --line: #374151;
  --brand-deep: #22c55e;
  --brand-main: #16a34a;
  --brand-muted: #064e3b;
  --success-bg: #064e3b;
  --success-border: #059669;
}
body.dark-mode header{
  background:linear-gradient(135deg, #0f172a, #1e293b);
  border-bottom:1px solid #334155;
}
body.dark-mode .sidebar{
  background:#1f2937;
  border-right-color:#374151;
}
body.dark-mode .filter-section{
  background:#1f2937;
  border-color:#374151;
}
body.dark-mode .sidebar-tabs{
  background:#374151;
}
body.dark-mode .sidebar-tab.active{
  background:#1f2937;
  color:var(--brand-accent);
}
body.dark-mode .state-checkbox{
  background:#1f2937;
  border-color:#374151;
}
body.dark-mode .state-checkbox:hover{
  background:#111827;
}
body.dark-mode .state-dropdown{
  border-color:#374151;
}
body.dark-mode .state-dropdown .dropdown-header{
  color:#f1f5f9;
}
body.dark-mode .state-dropdown .dropdown-body label{
  border-bottom-color:#374151;
}
body.dark-mode .state-dropdown .dropdown-body label:hover{
  background:#111827;
}
body.dark-mode .search-box input,
body.dark-mode .map-search{
  background:#1f2937;
  border-color:#374151;
  color:#f1f5f9;
}
body.dark-mode .map-search input{ color:#f1f5f9; }
body.dark-mode .substation-list{
  background:#1f2937;
  border-color:#374151;
}
body.dark-mode .substation-item:hover{
  background:#111827;
}
body.dark-mode .substation-name{
  color:var(--brand-accent);
}
body.dark-mode .min-btn{
  background:#374151;
  border-color:#4b5563;
  color:#94a3b8;
}
body.dark-mode .min-btn:hover{
  background:#111827;
  color:var(--brand-accent);
}
body.dark-mode .toast{
  background:#1f2937;
  border-color:#059669;
  color:#f1f5f9;
}
body.dark-mode .msm-summary-table thead th{
  background:#1e293b;
  box-shadow: inset 0 -2px 0 #374151;
  color:var(--brand-accent);
}
body.dark-mode .msm-summary-table tbody td{
  border-bottom-color:#374151;
}
body.dark-mode .msm-summary-table tbody tr:nth-child(odd){
  background:#1e293b;
}
body.dark-mode .msm-summary-table tbody tr:hover{
  background:#111827;
}
body.dark-mode .msm-summary-table input[type="text"]{
  background:#1f2937;
  border-color:#374151;
  color:#f1f5f9;
}
body.dark-mode .msm-summary-table-wrapper{
  border-color:#374151;
  background:#1f2937;
}
body.dark-mode .msm-summary-wrapper > .filter-section.boundary{
  background:#1f2937;
  border-color:#374151;
}
body.dark-mode .msm-summary-filter-select{
  background:#1f2937;
  border-color:#374151;
  color:#f1f5f9;
}
body.dark-mode .dark-toggle i::before{
  content:"\\f185";
}
body.dark-mode .boundary-field input{
  background:#1f2937;
  border-color:#374151;
  color:#f1f5f9;
}
body.dark-mode button.mini-btn{
  background:#374151;
  border-color:#4b5563;
  color:#94a3b8;
}
body.dark-mode .backdrop{
  background:rgba(0,0,0,.55);
}
body.dark-mode .sidebar-close{
  background:#374151;
  border-color:#4b5563;
}
@media (max-width:768px){
  body.dark-mode .sidebar{
    border-left-color:#374151;
  }
}


    /* ===== SOLAR BID ENHANCEMENTS ===== */
    .status-badge{
      display:inline-block; padding:2px 8px; border-radius:999px;
      font-size:10px; font-weight:700; letter-spacing:.03em; white-space:nowrap;
    }
    .status-Prospecting{ background:#e8f5e9; color:#1b5e20; }
    .status-Shortlisted{ background:#fff9c4; color:#e65100; }
    .status-Submitted  { background:#e3f2fd; color:#0d47a1; }
    .status-Awarded    { background:#c8e6c9; color:#1b5e20; border:1px solid #2e7d32; }
    .status-Rejected   { background:#ffcdd2; color:#b71c1c; }

    .solar-scorecard{
      background:#f8fffe; border:1px solid var(--line);
      border-radius:var(--radius-sm); padding:8px 10px; margin-top:6px;
      font-size:11px; line-height:1.6;
    }
    .solar-scorecard h4{
      font-size:11px; font-weight:700; color:var(--brand-deep);
      margin:0 0 4px; display:flex; align-items:center; gap:5px;
    }
    .solar-scorecard table{ width:100%; border-collapse:collapse; }
    .solar-scorecard td{ padding:1px 4px; }
    .solar-scorecard td:first-child{ color:var(--muted); width:45%; }
    .solar-scorecard td:last-child{ font-weight:600; color:var(--ink); }
    .solar-scorecard .sc-divider{
      border:none; border-top:1px solid var(--line); margin:4px 0;
    }
    .solar-scorecard .sc-section{
      font-size:10px; font-weight:700; color:var(--brand-main);
      text-transform:uppercase; letter-spacing:.04em; margin:4px 0 2px;
    }
    .boundary-field select, .boundary-field textarea{
      width:100%; border-radius:6px; border:1px solid var(--line);
      padding:4px 6px; font-size:12px; font-family:inherit;
      transition:all .15s; background:#fff;
    }
    .boundary-field textarea{ resize:vertical; min-height:42px; }
    .boundary-field select:focus, .boundary-field textarea:focus{
      outline:none; border-color:var(--brand-accent);
      box-shadow:0 0 0 2px rgba(74,222,128,.12);
    }
    body.dark-mode .boundary-field select,
    body.dark-mode .boundary-field textarea{
      background:#1f2937; border-color:#374151; color:#f1f5f9;
    }
    body.dark-mode .solar-scorecard{ background:#1f2937; border-color:#374151; }

    /* ===== PASSCODE LOCK ===== */
    .lock-btn{
      background:rgba(255,255,255,.12);
      border:1px solid rgba(255,255,255,.2);
      color:#fff;
      border-radius:var(--radius-sm);
      width:36px; height:36px;
      display:flex; align-items:center; justify-content:center;
      cursor:pointer;
      transition:all .2s;
      font-size:1rem;
    }
    .lock-btn:hover{ background:rgba(255,255,255,.22); }
    .lock-btn.unlocked{ background:rgba(74,222,128,.25); border-color:rgba(74,222,128,.5); }
    .lock-btn.unlocked:hover{ background:rgba(74,222,128,.35); }

    /* Modal overlay */
    .msm-modal-overlay{
      position:fixed; inset:0; background:rgba(0,0,0,.55);
      z-index:9000; display:none; align-items:center; justify-content:center;
      backdrop-filter:blur(4px);
    }
    .msm-modal-overlay.show{ display:flex; }
    .msm-modal{
      background:#fff; border-radius:var(--radius-lg); padding:28px 28px 22px;
      box-shadow:var(--shadow-lg); width:min(360px,90vw);
      animation:toast-slide .2s ease;
    }
    body.dark-mode .msm-modal{ background:#1f2937; color:#f1f5f9; }
    .msm-modal h2{ font-size:1.05rem; font-weight:700; margin:0 0 6px; color:var(--brand-deep); }
    .msm-modal p{ font-size:.85rem; color:var(--muted); margin:0 0 16px; }
    .msm-modal input[type=password]{
      width:100%; padding:9px 12px; border:1px solid var(--line);
      border-radius:var(--radius-sm); font-size:.95rem; font-family:inherit;
      outline:none; transition:all .2s; margin-bottom:12px;
    }
    body.dark-mode .msm-modal input[type=password]{
      background:#111827; border-color:#374151; color:#f1f5f9;
    }
    .msm-modal input[type=password]:focus{ border-color:var(--brand-accent); box-shadow:0 0 0 3px rgba(74,222,128,.15); }
    .msm-modal-actions{ display:flex; gap:8px; }
    .msm-modal-error{ font-size:.82rem; color:#b91c1c; margin:-4px 0 10px; display:none; }
    .msm-modal-error.show{ display:block; }

    /* ===== FILTER BAR (replaces in-table filter row) ===== */
    .msm-filter-bar{
      display:flex; align-items:center; gap:6px; flex-wrap:wrap;
      padding:6px 0 8px; margin-bottom:4px;
    }
    .msm-filter-bar-label{
      font-size:10px; font-weight:700; color:var(--muted);
      text-transform:uppercase; letter-spacing:.04em; white-space:nowrap;
    }
    .msm-filter-bar select{
      padding:3px 22px 3px 8px; border-radius:999px;
      border:1px solid var(--line); background:#fafffe;
      font-size:11px; font-family:inherit; color:var(--ink);
      cursor:pointer; transition:all .15s;
      appearance:none; -webkit-appearance:none;
      background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%236b7280'/%3E%3C/svg%3E");
      background-repeat:no-repeat; background-position:right 7px center;
      max-width:120px;
    }
    .msm-filter-bar select:focus{
      outline:none; border-color:var(--brand-accent);
      box-shadow:0 0 0 2px rgba(74,222,128,.15);
    }
    .msm-filter-bar select:hover{ border-color:var(--brand-muted); }
    .msm-filter-bar .filter-active{ border-color:var(--brand-main); background:#ecfdf5; }
    .msm-filter-clear{
      padding:3px 10px; border-radius:999px; border:1px solid var(--line);
      background:#f8fafc; font-size:11px; font-family:inherit; color:var(--muted);
      cursor:pointer; transition:all .15s; white-space:nowrap;
    }
    .msm-filter-clear:hover{ background:#fef2f2; border-color:#fecaca; color:#b91c1c; }
    body.dark-mode .msm-filter-bar select{
      background-color:#1f2937; border-color:#374151; color:#f1f5f9;
    }
    body.dark-mode .msm-filter-clear{
      background:#374151; border-color:#4b5563; color:#94a3b8;
    }

    /* ===== IMPROVED SUMMARY TABLE ===== */
    .msm-summary-table thead th{
      padding: 7px 8px;
      font-size: 10px;
      letter-spacing: 0.05em;
      border-bottom: 2px solid #d1fae5;
    }
    .msm-summary-table tbody td{
      padding: 6px 8px;
      vertical-align: middle;
    }
    /* Column group shading */
    .msm-summary-table td.col-identity,
    .msm-summary-table th.col-identity{ background:rgba(74,222,128,.04); }
    .msm-summary-table td.col-grid{ background:rgba(59,130,246,.04); }
    .msm-summary-table th.col-grid{ background:#e8f0fe; } /* solid — must be opaque for sticky */
    body.dark-mode .msm-summary-table th.col-grid{ background:#1e3a5f; }
    .msm-summary-table tbody tr:hover td{ background:#ecfdf5 !important; }

    /* Delete button only visible when editing is unlocked */
    .delete-shape-btn { display: none !important; }
    body.msm-unlocked .delete-shape-btn { display: inline-flex !important; }

    /* ===== GRID CONNECTION CHOICE MODAL ===== */
    .msm-choice-modal-overlay{
      position:fixed; inset:0; background:rgba(0,0,0,.55);
      z-index:9100; display:none; align-items:center; justify-content:center;
      backdrop-filter:blur(4px);
    }
    .msm-choice-modal-overlay.show{ display:flex; }
    .msm-choice-modal{
      background:#fff; border-radius:var(--radius-lg); padding:22px 22px 16px;
      box-shadow:var(--shadow-lg); width:min(420px,92vw);
      animation:toast-slide .2s ease;
    }
    body.dark-mode .msm-choice-modal{ background:#1f2937; color:#f1f5f9; }
    .msm-choice-modal h2{
      font-size:.95rem; font-weight:700; margin:0 0 4px; color:var(--brand-deep);
    }
    .msm-choice-modal p{
      font-size:.8rem; color:var(--muted); margin:0 0 14px;
    }
    .msm-choice-section{ margin-bottom:14px; }
    .msm-choice-section-title{
      font-size:.7rem; font-weight:700; text-transform:uppercase;
      letter-spacing:.05em; color:var(--brand-main); margin-bottom:6px;
    }
    .msm-choice-option{
      display:flex; align-items:center; gap:10px;
      padding:8px 10px; border-radius:var(--radius-sm);
      border:2px solid var(--line); cursor:pointer;
      transition:all .15s; margin-bottom:6px; background:#fafffe;
    }
    body.dark-mode .msm-choice-option{ background:#111827; border-color:#374151; }
    .msm-choice-option:hover{ border-color:var(--brand-accent); background:#ecfdf5; }
    .msm-choice-option.selected{ border-color:var(--brand-main); background:#ecfdf5; }
    body.dark-mode .msm-choice-option.selected{ background:#064e3b; border-color:#059669; }
    .msm-choice-option input[type=radio]{ accent-color:var(--brand-main); width:16px; height:16px; flex-shrink:0; }
    .msm-choice-label{ flex:1; }
    .msm-choice-label strong{ font-size:.85rem; display:block; color:var(--ink); }
    .msm-choice-label small{ font-size:.75rem; color:var(--muted); }
    .msm-choice-actions{ display:flex; gap:8px; margin-top:4px; }

    /* ===== MOBILE IMPROVEMENTS ===== */
    @media (max-width: 768px){
      /* Filter bar wraps nicely on small screens */
      .msm-filter-bar{ flex-wrap:wrap; gap:5px; }
      .msm-filter-bar select{ max-width:calc(50% - 6px); flex:1 1 auto; font-size:10px; }

      /* Solar scorecard full width */
      .solar-scorecard{ font-size:11px; }
      .solar-scorecard table{ table-layout:fixed; width:100%; }

      /* Boundary fields — full width inputs */
      .boundary-field input,
      .boundary-field select,
      .boundary-field textarea{ font-size:13px; padding:6px 8px; }

      /* Choice modal full screen on mobile */
      .msm-choice-modal{ width:95vw; padding:16px 14px 12px; }
      .msm-choice-modal h2{ font-size:.9rem; }

      /* Lock modal full width on mobile */
      .msm-modal{ width:92vw; padding:20px 16px 16px; }

      /* Saved Lands summary table — allow horizontal scroll */
      .msm-summary-table{ min-width:1100px; }

      /* Header stats hidden on very small screens (already done) */
      /* Route summary */
      #routeSummary{ font-size:12px; }

      /* Find connection button full width */
      #findConnectionBtn{ font-size:.82rem; padding:.5rem; }
    }

    /* Custom PMU pin option in choice modal */
    .msm-choice-option.custom-pin-opt{
      border-style: dashed;
      background: #fafff8;
    }
    body.dark-mode .msm-choice-option.custom-pin-opt{ background:#0f2318; }
    .msm-choice-option.custom-pin-opt:hover{ border-color:var(--brand-accent); }
    .custom-pin-name-row{
      display:none; margin-top:8px; gap:6px; align-items:center;
    }
    .custom-pin-name-row.show{ display:flex; }
    .custom-pin-name-row input{
      flex:1; padding:5px 8px; border-radius:6px; border:1px solid var(--line);
      font-size:12px; font-family:inherit;
    }
    body.dark-mode .custom-pin-name-row input{
      background:#111827; border-color:#374151; color:#f1f5f9;
    }

    /* Map cursor during pin-drop mode */
    body.msm-pin-drop-mode #map,
    body.msm-pin-drop-mode #map *{ cursor: crosshair !important; }
    .msm-pin-drop-banner{
      position:fixed; top:52px; left:50%; transform:translateX(-50%);
      background:rgba(13,74,26,.92); color:#fff; padding:8px 20px;
      border-radius:999px; font-size:13px; font-weight:600;
      z-index:9200; display:none; align-items:center; gap:8px;
      box-shadow:0 4px 16px rgba(0,0,0,.25);
      backdrop-filter:blur(6px);
    }
    .msm-pin-drop-banner.show{ display:flex; }
    .msm-pin-drop-banner button{
      background:rgba(255,255,255,.2); border:1px solid rgba(255,255,255,.4);
      color:#fff; border-radius:999px; padding:3px 10px; font-size:11px;
      cursor:pointer; font-family:inherit;
    }
    .msm-pin-drop-banner button:hover{ background:rgba(255,255,255,.35); }

    /* Explicit column widths to prevent header overlap */
    .msm-summary-table th:nth-child(1)  { min-width:30px;  width:30px; }   /* # */
    .msm-summary-table th:nth-child(2)  { min-width:90px;  width:90px; }   /* Status */
    .msm-summary-table th:nth-child(3)  { min-width:70px;  width:70px; }   /* Cap */
    .msm-summary-table th:nth-child(4)  { min-width:130px; }               /* Name */
    .msm-summary-table th:nth-child(5)  { min-width:110px; }               /* Land owner */
    .msm-summary-table th:nth-child(6)  { min-width:90px;  }               /* State */
    .msm-summary-table th:nth-child(7)  { min-width:90px;  }               /* City */
    .msm-summary-table th:nth-child(8)  { min-width:100px; }               /* Rental */
    .msm-summary-table th:nth-child(9)  { min-width:80px;  }               /* Acres */
    .msm-summary-table th:nth-child(10) { min-width:70px;  }               /* m² */
    .msm-summary-table th:nth-child(11) { min-width:130px; }               /* Nearest Sub */
    .msm-summary-table th:nth-child(12) { min-width:90px;  }               /* Grid Conn */
    .msm-summary-table th:nth-child(13) { min-width:50px;  width:50px; }   /* Action */
    .msm-summary-table th.col-grid{ white-space:nowrap; }

    /* Custom substation search inside grid choice modal */
    .custom-sub-search-wrap{
      display:none; margin-top:8px; position:relative;
    }
    .custom-sub-search-wrap.show{ display:block; }
    .custom-sub-search-wrap input{
      width:100%; padding:7px 10px; border-radius:8px;
      border:1px solid var(--line); font-size:13px; font-family:inherit;
      box-sizing:border-box; transition:all .15s;
    }
    body.dark-mode .custom-sub-search-wrap input{
      background:#111827; border-color:#374151; color:#f1f5f9;
    }
    .custom-sub-search-wrap input:focus{
      outline:none; border-color:var(--brand-accent);
      box-shadow:0 0 0 2px rgba(74,222,128,.15);
    }
    .custom-sub-dropdown{
      position:absolute; top:calc(100% + 4px); left:0; right:0;
      background:#fff; border:1px solid var(--line);
      border-radius:8px; box-shadow:var(--shadow-lg);
      max-height:180px; overflow-y:auto; z-index:9300;
      display:none;
    }
    body.dark-mode .custom-sub-dropdown{ background:#1f2937; border-color:#374151; }
    .custom-sub-dropdown.show{ display:block; }
    .custom-sub-item{
      padding:8px 12px; cursor:pointer; font-size:12px;
      border-bottom:1px solid #f1f5f9; transition:background .1s;
    }
    body.dark-mode .custom-sub-item{ border-bottom-color:#374151; color:#f1f5f9; }
    .custom-sub-item:last-child{ border-bottom:none; }
    .custom-sub-item:hover, .custom-sub-item.highlighted{
      background:#ecfdf5;
    }
    body.dark-mode .custom-sub-item:hover,
    body.dark-mode .custom-sub-item.highlighted{ background:#064e3b; }
    .custom-sub-item strong{ display:block; }
    .custom-sub-item small{ color:var(--muted); }
    .custom-sub-selected{
      display:none; margin-top:5px; font-size:11px;
      color:var(--brand-deep); font-weight:600;
    }
    .custom-sub-selected.show{ display:block; }
    .custom-sub-no-results{
      padding:10px 12px; font-size:12px; color:var(--muted); text-align:center;
    }
    .custom-pin-hint{
      font-size:11px; color:var(--muted); margin-top:5px; display:none;
    }
    .custom-pin-hint.show{ display:block; }

    /* ===== CONNECTION LINE OVERLAY ===== */
    .msm-conn-label{
      position:absolute;
      background:rgba(13,74,26,.88);
      color:#fff;
      font-size:11px;
      font-weight:600;
      padding:4px 9px;
      border-radius:999px;
      pointer-events:none;
      white-space:nowrap;
      box-shadow:0 2px 8px rgba(0,0,0,.25);
      backdrop-filter:blur(4px);
      z-index:200;
      transform:translate(-50%,-50%);
      display:none;
    }
    .msm-conn-label.show{ display:block; }
    .msm-conn-label.line-label{
      background:rgba(21,101,192,.88);
    }

    /* ========================================================
       MOBILE POLISH v2 — consolidated mobile UX fixes
       These rules appear LAST so they reliably override earlier
       conflicting mobile styles.
       ======================================================== */
    @media (max-width: 768px){

      /* --- Header tightening --- */
      header{
        padding: 6px 10px !important;
      }
      .header-content{
        gap: 6px !important;
      }
      header h1{
        font-size: 0.95rem !important;
        line-height: 1.15 !important;
      }
      header p{
        font-size: 0.65rem !important;
        line-height: 1.1 !important;
      }
      .logo-icon{
        font-size: 1.1rem !important;
      }
      .header-brand{ min-width: 0; }

      /* --- Give the map the full screen below the header --- */
      .content{
        height: calc(100dvh - 40px) !important;
        gap: 0 !important;
      }
      .map-container{
        height: calc(100dvh - 40px) !important;
        width: 100% !important;
        min-width: 0 !important;
      }

      /* --- Sidebar drawer: trim bottom padding, smoother motion --- */
      .sidebar{
        padding-bottom: 32px !important;   /* was 90px — freed up ~60px */
        width: min(90vw, 380px) !important;
      }

      /* --- Hide the floating bottom transmission toolbar on mobile.
             Users can reach those controls from the sidebar's Layers tab
             to prevent overlap with the map + coord bar + Google controls. --- */
      .layers-panel.toolbar{
        display: none !important;
      }

      /* --- Coordinate bar: push above the Google Maps controls bar --- */
      .coord-bar{
        bottom: 72px !important;
        font-size: 10px !important;
        padding: 4px 10px !important;
        max-width: calc(100vw - 24px) !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
      }

      /* --- Sidebar tabs: clearer touch targets, prevent wrap --- */
      .sidebar-tabs{
        padding: 3px !important;
        margin: 4px 0 8px !important;
      }
      .sidebar-tab{
        padding: 9px 2px !important;
        font-size: 0.72rem !important;
        gap: 3px !important;
      }
      .sidebar-tab i{
        font-size: 0.8rem !important;
      }
      /* hide the text label on very narrow screens, show only icons */
      @media (max-width: 380px){
        .sidebar-tab span.tab-label{ display: none !important; }
        .sidebar-tab{ padding: 11px 4px !important; }
        .sidebar-tab i{ font-size: 1rem !important; }
      }

      /* --- Saved Lands table: make it stack as readable cards
             instead of forcing a 700px horizontal scroll --- */
      .msm-summary-table{
        min-width: 0 !important;
      }
      .msm-summary-table-wrapper{
        overflow-x: visible !important;
      }
      .msm-summary-table thead{
        display: none !important;
      }
      .msm-summary-table tbody tr{
        display: block !important;
        border: 1px solid var(--line) !important;
        border-radius: var(--radius-sm) !important;
        padding: 8px !important;
        margin-bottom: 8px !important;
        background: #fff !important;
      }
      .msm-summary-table tbody tr td{
        display: flex !important;
        justify-content: space-between !important;
        align-items: center !important;
        gap: 8px !important;
        padding: 4px 0 !important;
        border: 0 !important;
        font-size: 0.82rem !important;
      }
      .msm-summary-table tbody tr td::before{
        font-weight: 600;
        color: var(--brand-deep);
        font-size: 0.72rem;
        text-transform: uppercase;
        letter-spacing: 0.02em;
        min-width: 82px;
        flex-shrink: 0;
      }
      .msm-summary-table tbody tr td:nth-child(1)::before  { content: "#"; }
      .msm-summary-table tbody tr td:nth-child(2)::before  { content: "Status"; }
      .msm-summary-table tbody tr td:nth-child(3)::before  { content: "Cap (MWp)"; }
      .msm-summary-table tbody tr td:nth-child(4)::before  { content: "Name"; }
      .msm-summary-table tbody tr td:nth-child(5)::before  { content: "Land owner"; }
      .msm-summary-table tbody tr td:nth-child(6)::before  { content: "State"; }
      .msm-summary-table tbody tr td:nth-child(7)::before  { content: "City"; }
      .msm-summary-table tbody tr td:nth-child(8)::before  { content: "Rent (RM)"; }
      .msm-summary-table tbody tr td:nth-child(9)::before  { content: "Area (ac)"; }
      .msm-summary-table tbody tr td:nth-child(10)::before { content: "Area (m²)"; }
      .msm-summary-table tbody tr td:nth-child(11)::before { content: "Nearest Sub"; }
      .msm-summary-table tbody tr td:nth-child(12)::before { content: "Grid Conn"; }
      .msm-summary-table tbody tr td:nth-child(13)::before { content: "Action"; }
      /* hide the "No lands drawn yet" placeholder's pseudo-label */
      .msm-summary-table tbody tr td[colspan]::before{
        content: none !important;
      }
      .msm-summary-table tbody tr td[colspan]{
        display: block !important;
        text-align: center !important;
      }
      .msm-summary-table tbody tr td input,
      .msm-summary-table tbody tr td select{
        flex: 1 !important;
        min-width: 0 !important;
      }
      body.dark-mode .msm-summary-table tbody tr{
        background: #1f2937 !important;
        border-color: #374151 !important;
      }

      /* --- Filter sections: tighter on phones --- */
      .filter-section{
        padding: 0.7rem !important;
        margin-bottom: 0.5rem !important;
      }
      .filter-section h3{
        font-size: 0.88rem !important;
        margin-bottom: 0.6rem !important;
      }

      /* --- Buttons: comfortable touch --- */
      .filter-btn{
        min-height: 40px !important;
      }

      /* --- Search input: larger to avoid iOS zoom-on-focus --- */
      .search-box input{
        font-size: 16px !important;   /* prevents iOS auto-zoom */
      }
      input[type="text"],
      input[type="url"],
      input[type="number"],
      input[type="search"],
      textarea,
      select{
        font-size: 16px !important;   /* prevents iOS auto-zoom */
      }

      /* --- Substation list: more breathing room --- */
      .substation-list{
        max-height: 50vh !important;
      }

      /* --- Hamburger: easier to tap --- */
      .hamburger{
        padding: 0.55rem 0.75rem !important;
        font-size: 1rem !important;
      }
    }

    /* Ultra-narrow phones */
    @media (max-width: 360px){
      header h1{ font-size: 0.85rem !important; }
      header p{ display: none !important; }
      .sidebar{ width: 95vw !important; }
    }
  </style>
</head>
<body>
<div class="full-bleed">
  <div id="msm-app" class="container">
      <header>
        <div class="header-content">
          <div class="header-left">
            <div class="header-brand">
              <i class="fas fa-bolt logo-icon"></i>
              <div>
                <h1>Substation Map</h1>
                <p>Malaysian Grid Infrastructure Viewer</p>
              </div>
            </div>
          </div>
          <div class="header-right">
            <button id="lockToggleBtn" class="dark-toggle lock-btn locked" title="Unlock editing (passcode required)">
              <i class="fas fa-lock"></i>
            </button>
            <a href="https://samaiden-ai.com.my/solar-farm-drawing-and-engineering/"
               target="_blank"
               rel="noopener noreferrer"
               class="header-link-btn"
               title="Open Solar Farm Drawing &amp; Engineering page">
              <i class="fas fa-drafting-compass"></i>
              <span class="link-label-full">Drawing Page</span>
              <span class="link-label-short">Drawing</span>
            </a>
            <a href="https://samaiden-ai.com.my/site-survey-for-solar-farm/"
               target="_blank"
               rel="noopener noreferrer"
               class="header-link-btn"
               title="Open page">
              <i class="fas fa-external-link-alt"></i>
              <span class="link-label-full">Site Survey</span>
              <span class="link-label-short">Site Survey</span>
            </a>
            <button id="mobileMenuBtn" class="hamburger" aria-label="Open filters">
              <i class="fas fa-sliders-h"></i>
            </button>
          </div>
        </div>
      </header>

      <div class="content">
        <div class="sidebar">
          <button id="closeSidebarBtn" class="sidebar-close" aria-label="Close">
            <i class="fas fa-times"></i>
          </button>

          <div class="map-search">
            <i class="fas fa-search"></i>
            <input id="mapSearchInput" type="text" placeholder="Search places, addresses…">
            <button id="locateMeBtn" title="Use my location"><i class="fas fa-location-arrow"></i></button>
          </div>

          <!-- Sidebar Tabs -->
          <div class="sidebar-tabs">
            <button class="sidebar-tab active" data-tab="tab-search">
              <i class="fas fa-search"></i> <span class="tab-label">Search</span>
            </button>
            <button class="sidebar-tab" data-tab="tab-layers">
              <i class="fas fa-layer-group"></i> <span class="tab-label">Layers</span>
            </button>
            <button class="sidebar-tab" data-tab="tab-tools">
              <i class="fas fa-tools"></i> <span class="tab-label">Tools</span>
            </button>
          </div>

          <!-- TAB 1: Search & Filter -->
          <div class="tab-panel active" id="tab-search">

          <!-- Compact tools that will follow the map into fullscreen -->
          <div id="mapToolsControl" class="as-control collapsible" style="display:none; position:relative;">
            <div class="card-header" style="position:relative; padding-right:42px; margin-bottom:8px;">
              <strong>Search / Route</strong>
              <button class="min-btn" aria-expanded="true" aria-controls="mapToolsBody">–</button>
            </div>
            <div id="mapToolsBody" class="card-body">
              <div class="map-search" style="margin-bottom:8px;">
                <i class="fas fa-search"></i>
                <input id="mapSearchInputFS" type="text" placeholder="Search places, addresses…">
                <button id="locateMeBtnFS" title="Use my location"><i class="fas fa-location-arrow"></i></button>
              </div>

              <div style="display:flex; gap:8px; flex-wrap:wrap;">
                <button id="routeStartBtnFS" class="filter-btn"><i class="fas fa-flag"></i> Start</button>
                <button id="routeEndBtnFS" class="filter-btn"><i class="fas fa-location-dot"></i> End</button>
                <button id="clearRouteBtnFS" class="filter-btn secondary"><i class="fas fa-eraser"></i> Clear</button>
              </div>

              <div id="routeSummaryFS" style="margin-top:6px;font-size:12px;color:#1b5e20;">
                Distance: – | Time: –
              </div>
            </div>
          </div>

          <div class="search-box">
            <input type="text" id="searchInput" placeholder="Filter substations by name…">
          </div>
    
          <div class="filter-section">
            <h3><i class="fas fa-filter"></i> Filter by State</h3>
            <div class="state-dropdown" id="stateDropdown">
              <div class="dropdown-header" id="stateDropdownHeader">
                <span id="stateDropdownLabel">All States Selected</span>
                <i class="fas fa-chevron-down"></i>
              </div>
              <div class="dropdown-body" id="stateDropdownBody"></div>
            </div>
            <div class="filter-actions">
              <button class="filter-btn" id="selectAllBtn"><i class="fas fa-check-square"></i> Select All</button>
              <button class="filter-btn secondary" id="clearAllBtn"><i class="fas fa-times-circle"></i> Clear All</button>
            </div>
          </div>

          <div class="filter-section">
            <h3><i class="fas fa-map-marker-alt"></i> Substations</h3>
            <div class="substation-list" id="substationList"></div>
          </div>

          <!-- Saved Candidate Sites -->
          <div class="filter-section">
            <h3><i class="fas fa-bookmark"></i> Saved Candidate Sites</h3>
            <div class="substation-list" id="savedSitesList"></div>
            <div class="filter-actions" style="margin-top:6px;">
              <button class="filter-btn secondary" id="clearSavedSitesBtn">
                <i class="fas fa-trash-alt"></i> Clear Saved Sites
              </button>
            </div>
          </div>

          </div><!-- end tab-search -->

          <!-- TAB 2: Layers -->
          <div class="tab-panel" id="tab-layers">

          <!-- Transmission Lines (OSM) -->
          <div id="layersPanel" class="filter-section layers-panel toolbar collapsible">
            <h3 style="position:relative; padding-right:42px;">
              <i class="fas fa-bolt"></i> Transmission Lines
              <button class="min-btn" aria-expanded="true" aria-controls="layersPanelBody">–</button>
            </h3>
            <div id="layersPanelBody" class="card-body">
              <div class="state-checkbox" style="justify-content:space-between;">
                <label><input type="checkbox" id="toggle11" checked style="margin-right:.5rem"> Show 11 kV</label>
                <span style="display:inline-flex;align-items:center;gap:6px;">
                  <span style="display:inline-block;width:24px;height:3px;background:#9C27B0;border-radius:2px"></span>
                  <small>11 kV</small>
                </span>
              </div>
              <div class="state-checkbox" style="justify-content:space-between;">
                <label><input type="checkbox" id="toggle33" checked style="margin-right:.5rem"> Show 33 kV</label>
                <span style="display:inline-flex;align-items:center;gap:6px;">
                  <span style="display:inline-block;width:24px;height:3px;background:#FF9800;border-radius:2px"></span>
                  <small>33 kV</small>
                </span>
              </div>
              <div class="state-checkbox" style="justify-content:space-between;">
                <label><input type="checkbox" id="toggle132" checked style="margin-right:.5rem"> Show 132 kV</label>
                <span style="display:inline-flex;align-items:center;gap:6px;">
                  <span style="display:inline-block;width:24px;height:3px;background:#FF1744;border-radius:2px"></span>
                  <small>132 kV</small>
                </span>
              </div>
              <div class="state-checkbox" style="justify-content:space-between;">
                <label><input type="checkbox" id="toggle275" checked style="margin-right:.5rem"> Show 275 kV</label>
                <span style="display:inline-flex;align-items:center;gap:6px;">
                  <span style="display:inline-block;width:24px;height:3px;background:#00A8FF;border-radius:2px"></span>
                  <small>275 kV</small>
                </span>
              </div>
              <div class="state-checkbox" style="justify-content:space-between;">
                <label><input type="checkbox" id="toggle500" checked style="margin-right:.5rem"> Show 500 kV</label>
                <span style="display:inline-flex;align-items:center;gap:6px;">
                  <span style="display:inline-block;width:24px;height:3px;background:#C97125;border-radius:2px"></span>
                  <small>500 kV</small>
                </span>
              </div>

              <!-- KTM Rail -->
              <div class="state-checkbox" style="justify-content:space-between;">
                <label>
                  <input type="checkbox" id="toggleKTM" checked style="margin-right:.5rem">
                  Show KTM Rail
                </label>
                <span style="display:inline-flex;align-items:center;gap:6px;">
                  <span style="display:inline-block;width:24px;height:3px;background:#F0C50E;border-radius:2px"></span>
                  <small>KTM</small>
                </span>
              </div>

              <!-- Highways -->
              <div class="state-checkbox" style="justify-content:space-between;">
                <label><input type="checkbox" id="toggleHwy" checked style="margin-right:.5rem"> Show Highways</label>
                <span style="display:inline-flex;align-items:center;gap:6px;">
                  <span style="display:inline-block;width:24px;height:3px;background:#FBC02D;border-radius:2px"></span>
                  <small>Highway</small>
                </span>
              </div>

              <div class="filter-actions">
                <button class="filter-btn" id="selectAllLinesBtn">
                  <i class="fas fa-check-square"></i> Select All
                </button>
                <button class="filter-btn secondary" id="clearLinesBtn">
                  <i class="fas fa-eraser"></i> Clear
                </button>
              </div>
              <div id="osmStatus" style="margin-top:6px;font-size:12px;color:var(--brand-deep);"></div>
            </div>
          </div>

          <!-- Route Measure -->
          <div class="filter-section">
            <h3><i class="fas fa-route"></i> Measure via Road</h3>
            <div class="filter-actions">
              <button id="routeStartBtn" class="filter-btn"><i class="fas fa-flag"></i> Start</button>
              <button id="routeEndBtn" class="filter-btn"><i class="fas fa-location-dot"></i> Destination</button>
            </div>
            <div class="filter-actions" style="margin-top:8px;">
              <button id="clearRouteBtn" class="filter-btn secondary"><i class="fas fa-eraser"></i> Clear</button>
            </div>
            <div id="routeSummary" style="margin-top:8px;font-size:13px;">
              Distance: – | Time: –
            </div>
            <small style="color:var(--muted);font-size:11px;">Click "Start" then "Destination" on the map to measure driving distance</small>
          </div>

          </div><!-- end tab-layers -->

          <!-- TAB 3: Tools -->
          <div class="tab-panel" id="tab-tools">

          <!-- Site Boundary panel will be inserted here by JS -->

          </div><!-- end tab-tools -->

          <!-- Upload (hidden) -->
          <div class="filter-section import-section">
            <h3><i class="fas fa-file-import"></i> Import from Excel / CSV</h3>
            <div id="dropZone" class="uploader">
              <div class="row">
                <input id="mediaUrl" type="url" placeholder="https://samaiden-ai.com.my/wp-content/uploads/2025/05/TNB-Substation-Peninsular-Malaysia.xlsx" />
                <button id="loadFromUrl" class="sample-btn" title="Load the file from this URL">Load URL</button>
              </div>
              <div class="row">
                <input id="fileInput" type="file" accept=".xlsx,.xls,.csv" style="display:none"/>
                <button id="importExcelBtn" class="sample-btn" title="Import a local Excel/CSV file">Import from Excel</button>
              </div>
              <div id="importStatus" class="import-status"></div>
            </div>
          </div>

        </div>

        <div class="map-container">
          <div id="map"></div>
          <!-- Floating coordinate bar -->
          <div id="coordBar" class="coord-bar">
            <span id="coordText">Move cursor over map</span>
            <span class="coord-divider">|</span>
            <span id="zoomText">Zoom: 6</span>
          </div>
        </div>
        <div id="backdrop" class="backdrop"></div>
      </div>
      
<!-- Global Saved Lands Summary under the map -->
<div class="msm-summary-wrapper">
  <div class="filter-section boundary">
    <div class="msm-summary-header">
      <div>
        <h3 style="margin:0;">
          <i class="fas fa-draw-polygon"></i> Saved Lands Summary
        </h3>
        <div class="msm-summary-meta" id="savedLandsMeta">
          0 sites saved.
        </div>
      </div>
        <div style="display:flex; gap:6px; flex-wrap:wrap;">
          <button id="importKmzBtn"
                  class="filter-btn secondary msm-export-btn"
                  type="button"
                  title="Import lands from a KMZ/KML file (unlock editing first)"
                  style="white-space:nowrap;">
            <i class="fas fa-file-import"></i> Import KMZ
          </button>
          <input id="kmzFileInput" type="file" accept=".kmz,.kml,application/vnd.google-earth.kmz,application/vnd.google-earth.kml+xml" style="display:none" />
          <button id="exportKmzBtn"
                  class="filter-btn msm-export-btn"
                  type="button"
                  title="Export all saved lands to a KMZ file (Google Earth)"
                  style="white-space:nowrap;background:linear-gradient(135deg,#1a5490,#2e7cd6);">
            <i class="fas fa-earth-asia"></i> Export KMZ
          </button>
          <button id="exportExcelBtn"
                  class="filter-btn msm-export-btn"
                  type="button"
                  style="white-space:nowrap;">
            <i class="fas fa-file-excel"></i> Export to Excel
          </button>
        </div>
    </div>

    <!-- Filter bar — above the table, outside thead -->
    <div class="msm-filter-bar">
      <span class="msm-filter-bar-label"><i class="fas fa-filter"></i> Filter</span>
      <select id="filterStatus" title="Status">
        <option value="">All status</option>
        <option value="Prospecting">Prospecting</option>
        <option value="Shortlisted">Shortlisted</option>
        <option value="Submitted">Submitted</option>
        <option value="Awarded">Awarded</option>
        <option value="Rejected">Rejected</option>
      </select>
      <select id="filterLandName" title="Site name"><option value="">All names</option></select>
      <select id="filterLandOwner" title="Land owner"><option value="">All owners</option></select>
      <select id="filterState" title="State"><option value="">All states</option></select>
      <select id="filterCity" title="City"><option value="">All cities</option></select>
      <select id="filterRent" title="Rental fee"><option value="">All rents</option></select>
      <button class="msm-filter-clear" id="clearFiltersBtn"><i class="fas fa-times"></i> Clear</button>
    </div>

    <div class="msm-summary-table-wrapper">
            <table class="msm-summary-table">
              <thead>
                <tr>
                  <th>#</th>
                  <th>Status</th>
                  <th>Cap (MWp)</th>
                  <th>Name</th>
                  <th>Land owner</th>
                  <th>State</th>
                  <th>City</th>
                  <th>Rental fee (RM)</th>
                  <th>Area (acres)</th>
                  <th>Area (m²)</th>
                  <th class="col-grid">Nearest Sub</th>
                  <th class="col-grid">Grid Conn</th>
                  <th>Action</th>
                </tr>
              </thead>
            
              <tbody id="boundarySummaryBody">

            <tr>
              <td colspan="13">No lands drawn yet.</td>
            </tr>
          </tbody>
        </table>

    </div>

    <small style="font-size:11px;color:#555;display:block;margin-top:4px;">
      Tip: click a shape to select; vertices are draggable. Summary will update automatically.
    </small>
  </div>
</div>


      
    </div>
</div>

<!-- Passcode Modal -->
<div id="passcodeModal" class="msm-modal-overlay">
  <div class="msm-modal">
    <h2><i class="fas fa-lock" style="margin-right:6px;color:var(--brand-main)"></i>Unlock Editing</h2>
    <p>Enter the passcode to enable drawing and saving.</p>
    <input type="password" id="passcodeInput" placeholder="Enter passcode…" autocomplete="off">
    <div id="passcodeError" class="msm-modal-error">Incorrect passcode. Please try again.</div>
    <div class="msm-modal-actions">
      <button id="passcodeSubmitBtn" class="filter-btn" style="flex:1"><i class="fas fa-unlock"></i> Unlock</button>
      <button id="passcodeCancelBtn" class="filter-btn secondary" style="flex:1">Cancel</button>
    </div>
  </div>
</div>

<!-- Delete Passcode Modal (second-layer protection) -->
<div id="deletePasscodeModal" class="msm-modal-overlay">
  <div class="msm-modal">
    <h2><i class="fas fa-triangle-exclamation" style="margin-right:6px;color:#b91c1c"></i>Confirm Delete</h2>
    <p style="color:#b91c1c;font-weight:500;">
      This will permanently delete the selected land boundary.<br>
      Enter the <strong>delete passcode</strong> to continue.
    </p>
    <input type="password" id="deletePasscodeInput" placeholder="Enter delete passcode…" autocomplete="off">
    <div id="deletePasscodeError" class="msm-modal-error">Incorrect delete passcode. Please try again.</div>
    <div class="msm-modal-actions">
      <button id="deletePasscodeSubmitBtn" class="filter-btn" style="flex:1;background:linear-gradient(135deg,#991b1b,#b91c1c);">
        <i class="fas fa-trash-alt"></i> Delete
      </button>
      <button id="deletePasscodeCancelBtn" class="filter-btn secondary" style="flex:1">Cancel</button>
    </div>
  </div>
</div>

<!-- Pin-drop mode banner -->
<div id="pinDropBanner" class="msm-pin-drop-banner">
  <i class="fas fa-map-pin"></i>
  Click the map to place the custom PMU location
  <button id="pinDropCancelBtn">Cancel</button>
</div>

<!-- Grid Connection Choice Modal -->
<div id="gridChoiceModal" class="msm-choice-modal-overlay">
  <div class="msm-choice-modal">
    <h2><i class="fas fa-plug" style="margin-right:6px;color:var(--brand-main)"></i>Choose Connection Type</h2>
    <p>Select <strong>one</strong> connection type for this site.</p>

    <!-- Connection type switcher -->
    <div style="display:flex;gap:6px;margin-bottom:14px;">
      <button id="connTypePmu" class="filter-btn" style="flex:1;font-size:12px;">
        <i class="fas fa-building"></i> PMU / Substation
      </button>
      <button id="connTypeGrid" class="filter-btn secondary" style="flex:1;font-size:12px;">
        <i class="fas fa-bolt"></i> Grid Line
      </button>
    </div>

    <!-- PMU panel -->
    <div id="connPanelPmu" class="msm-choice-section">
      <div class="msm-choice-section-title"><i class="fas fa-building"></i> Select PMU / Substation</div>
      <label class="msm-choice-option" id="subOpt1">
        <input type="radio" name="subChoice" value="1">
        <div class="msm-choice-label"><strong id="subOpt1Name">—</strong><small id="subOpt1Dist"></small></div>
      </label>
      <label class="msm-choice-option" id="subOpt2">
        <input type="radio" name="subChoice" value="2">
        <div class="msm-choice-label"><strong id="subOpt2Name">—</strong><small id="subOpt2Dist"></small></div>
      </label>
      <label class="msm-choice-option custom-pin-opt" id="subOptCustom">
        <input type="radio" name="subChoice" value="custom">
        <div class="msm-choice-label">
          <strong><i class="fas fa-search" style="color:var(--brand-main)"></i> Search substation list</strong>
          <small>Type to find any substation, or drop a pin if not found</small>
        </div>
      </label>
      <div class="custom-sub-search-wrap" id="customSubSearchWrap">
        <input type="text" id="customSubSearchInput" placeholder="Type substation name…" autocomplete="off">
        <div class="custom-sub-dropdown" id="customSubDropdown"></div>
        <div class="custom-sub-selected" id="customSubSelected"></div>
        <div class="custom-pin-hint" id="customPinHint">
          <i class="fas fa-map-pin"></i> Not found? <a href="#" id="switchToPinDrop" style="color:var(--brand-main);">Drop a pin on the map instead</a>
        </div>
      </div>
    </div>

    <!-- Grid line panel (hidden by default) -->
    <div id="connPanelGrid" class="msm-choice-section" style="display:none;">
      <div class="msm-choice-section-title"><i class="fas fa-bolt"></i> Select Transmission Line</div>
      <label class="msm-choice-option" id="lineOpt1">
        <input type="radio" name="lineChoice" value="1">
        <div class="msm-choice-label"><strong id="lineOpt1Kv">—</strong><small id="lineOpt1Dist"></small></div>
      </label>
      <label class="msm-choice-option" id="lineOpt2">
        <input type="radio" name="lineChoice" value="2">
        <div class="msm-choice-label"><strong id="lineOpt2Kv">—</strong><small id="lineOpt2Dist"></small></div>
      </label>
    </div>

    <div class="msm-choice-actions">
      <button id="gridChoiceConfirmBtn" class="filter-btn" style="flex:1"><i class="fas fa-check"></i> Confirm</button>
      <button id="gridChoiceCancelBtn" class="filter-btn secondary" style="flex:1">Cancel</button>
    </div>
  </div>
</div>

<!-- Toast -->
<div id="toast" class="toast"></div>

<script>

  // ===== PASSCODE LOCK SYSTEM =====
  window._msmPasscode = null;   // in-memory only — cleared on lock

  function msmIsUnlocked(){ return !!window._msmPasscode; }

  function msmSetLockUI(unlocked){
    const btn  = document.getElementById('lockToggleBtn');
    const icon = btn && btn.querySelector('i');
    if (!btn) return;
    document.body.classList.toggle('msm-unlocked', unlocked);
    if (unlocked){
      btn.classList.add('unlocked');
      btn.classList.remove('locked');
      btn.title = 'Lock editing';
      if (icon){ icon.className = 'fas fa-lock-open'; }
    } else {
      btn.classList.remove('unlocked');
      btn.classList.add('locked');
      btn.title = 'Unlock editing (passcode required)';
      if (icon){ icon.className = 'fas fa-lock'; }
    }
    // Toggle draw button availability
    ['drawPolygonBtn','drawRectangleBtn','clearBoundaryBtn','saveBoundaryBtn','findConnectionBtn','detectTerrainBtn']
      .forEach(id => {
        const el = document.getElementById(id);
        if (el){ el.style.opacity = unlocked ? '' : '0.4'; el.style.pointerEvents = unlocked ? '' : 'none'; }
      });
  }

  function msmLock(){
    window._msmPasscode = null;
    msmSetLockUI(false);
    if (window.drawingManager) window.drawingManager.setDrawingMode(null);
    showToast('Editing locked.');
  }

  function msmShowUnlockModal(){
    const modal = document.getElementById('passcodeModal');
    const input = document.getElementById('passcodeInput');
    const err   = document.getElementById('passcodeError');
    if (!modal) return;
    input.value = '';
    err.classList.remove('show');
    modal.classList.add('show');
    setTimeout(() => input.focus(), 80);
  }

  function msmHideUnlockModal(){
    const modal = document.getElementById('passcodeModal');
    if (modal) modal.classList.remove('show');
  }

  // Wire up modal interactions once DOM is ready
  document.addEventListener('DOMContentLoaded', function(){
    const lockBtn    = document.getElementById('lockToggleBtn');
    const submitBtn  = document.getElementById('passcodeSubmitBtn');
    const cancelBtn  = document.getElementById('passcodeCancelBtn');
    const input      = document.getElementById('passcodeInput');
    const err        = document.getElementById('passcodeError');

    if (lockBtn){
      lockBtn.addEventListener('click', () => {
        if (msmIsUnlocked()){ msmLock(); }
        else { msmShowUnlockModal(); }
      });
    }

    function attemptUnlock(){
      const val = input ? input.value : '';
      // Send to backend to verify
      const url = window.SHAPE_API_SAVE_URL;
      if (!url){ showToast('API not configured.'); return; }
      fetch(url.replace('/land-shapes', '/verify-passcode'), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ passcode: val })
      })
      .then(r => r.json())
      .then(data => {
        if (data && data.success){
          window._msmPasscode = val;
          msmSetLockUI(true);
          msmHideUnlockModal();
          showToast('Editing unlocked.');
        } else {
          if (err){ err.classList.add('show'); }
          if (input){ input.value = ''; input.focus(); }
        }
      })
      .catch(() => {
        if (err){ err.classList.add('show'); }
      });
    }

    if (submitBtn) submitBtn.addEventListener('click', attemptUnlock);
    if (cancelBtn) cancelBtn.addEventListener('click', msmHideUnlockModal);
    if (input){
      input.addEventListener('keydown', e => {
        if (e.key === 'Enter') attemptUnlock();
        if (e.key === 'Escape') msmHideUnlockModal();
      });
    }

    // Initialise UI in locked state
    msmSetLockUI(false);
  });

  // ===== DELETE PASSCODE SYSTEM (second-layer protection) =====
  // The delete passcode is NEVER stored — it must be entered each time.
  // This is on purpose: delete is destructive, so no "remember me" shortcut.
  window._msmDeleteCallback = null;

  function msmShowDeleteModal(onConfirm){
    const modal = document.getElementById('deletePasscodeModal');
    const input = document.getElementById('deletePasscodeInput');
    const err   = document.getElementById('deletePasscodeError');
    if (!modal) return;
    window._msmDeleteCallback = onConfirm || null;
    input.value = '';
    err.classList.remove('show');
    modal.classList.add('show');
    setTimeout(() => input.focus(), 80);
  }

  function msmHideDeleteModal(){
    const modal = document.getElementById('deletePasscodeModal');
    if (modal) modal.classList.remove('show');
    const input = document.getElementById('deletePasscodeInput');
    if (input) input.value = '';
    window._msmDeleteCallback = null;
  }

  // Wire up delete modal interactions
  document.addEventListener('DOMContentLoaded', function(){
    const submitBtn = document.getElementById('deletePasscodeSubmitBtn');
    const cancelBtn = document.getElementById('deletePasscodeCancelBtn');
    const input     = document.getElementById('deletePasscodeInput');
    const err       = document.getElementById('deletePasscodeError');

    function attemptDeleteUnlock(){
      const val = input ? input.value : '';
      const url = window.SHAPE_API_SAVE_URL;
      if (!url){ showToast('API not configured.'); return; }

      // Disable button during request
      if (submitBtn) submitBtn.disabled = true;

      fetch(url.replace('/land-shapes', '/verify-delete-passcode'), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ passcode: val })
      })
      .then(r => r.json())
      .then(data => {
        if (submitBtn) submitBtn.disabled = false;
        if (data && data.success){
          const cb = window._msmDeleteCallback;
          msmHideDeleteModal();
          if (typeof cb === 'function') cb();
        } else {
          if (err){ err.classList.add('show'); }
          if (input){ input.value = ''; input.focus(); }
        }
      })
      .catch(() => {
        if (submitBtn) submitBtn.disabled = false;
        if (err){ err.classList.add('show'); }
      });
    }

    if (submitBtn) submitBtn.addEventListener('click', attemptDeleteUnlock);
    if (cancelBtn) cancelBtn.addEventListener('click', msmHideDeleteModal);
    if (input){
      input.addEventListener('keydown', e => {
        if (e.key === 'Enter') attemptDeleteUnlock();
        if (e.key === 'Escape') msmHideDeleteModal();
      });
    }
  });

  // ---------- Data + Map State ----------
  // --- Global map + search landmark state ---
  let map, infoWindow;
  let searchMarker   = null;
  let searchIW       = null;
  let searchGeocoder = null;
  let searchLastTitle = '';
  let savedSites = [];   // [{id,name,lat,lng}]
  let shapeGeocoder = null;

    function copyShareLink(lat, lng){
      try{
        const z = map ? map.getZoom() : 17;
        const url = new URL(window.location.href);
        url.searchParams.set('lat', lat.toFixed(6));
        url.searchParams.set('lng', lng.toFixed(6));
        url.searchParams.set('z', z);
    
        navigator.clipboard.writeText(url.toString())
          .then(()=> showToast('Share link copied to clipboard.'))
          .catch(()=> showToast('Unable to copy link.'));
      }catch(e){
        console.error('copyShareLink failed', e);
        showToast('Error creating share link.');
      }
    }

function applyShareLinkFromUrl(){
  try{
    const params = new URLSearchParams(window.location.search);
    const plat = params.get('lat');
    const plng = params.get('lng');
    const pz   = params.get('z');

    if (!plat || !plng) return;

    const lat = parseFloat(plat);
    const lng = parseFloat(plng);
    if (!isFinite(lat) || !isFinite(lng)) return;

    const center = new google.maps.LatLng(lat, lng);
    map.setCenter(center);

    const z = parseInt(pz, 10);
    if (isFinite(z)) map.setZoom(z);

    // Show landmark with a default label if none yet
    setSearchLandmark(center, 'Shared view');
  }catch(e){
    console.error('applyShareLinkFromUrl failed', e);
  }
}

  // Close only the info window (used by the X button)
  function closeSearchInfo(){
    if (searchIW){
      searchIW.close();
    }
  }

  // Completely clear the landmark + info window
  function clearSearchLandmark(){
    if (searchMarker){
      searchMarker.setMap(null);
      searchMarker = null;
    }
    if (searchIW){
      searchIW.close();
      searchIW = null;
    }
  }

async function loadSavedSites(){
  try{
    // Public read: anyone can load if URL is configured
    if (!window.CANDIDATE_SITES_API_URL){
      savedSites = [];
      rebuildSavedSitesList();
      return;
    }

    const resp = await fetch(window.CANDIDATE_SITES_API_URL, {
      method: 'GET',
      headers: { 'Accept': 'application/json' }
    });

    const data = await resp.json();
    savedSites = Array.isArray(data?.sites) ? data.sites : [];
  }catch(e){
    console.warn('Failed to load saved sites from WordPress', e);
    savedSites = [];
  }
  rebuildSavedSitesList();
}

async function persistSavedSites(){
  try{
    // If not configured, do nothing
    if (!window.CANDIDATE_SITES_API_URL) return;

    // Check unlock
    if (!msmIsUnlocked()){
      showToast('Editing is locked. Click the lock icon to unlock.');
      return;
    }

    const resp = await fetch(window.CANDIDATE_SITES_API_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-WP-Nonce': window.CANDIDATE_SITES_API_NONCE || '',
        'X-Edit-Passcode': window._msmPasscode || ''
      },
      body: JSON.stringify({ sites: savedSites })
    });

    if (!resp.ok){
      showToast('WordPress save failed (HTTP ' + resp.status + ').');
    }
  }catch(e){
    console.warn('Failed to save sites to WordPress', e);
    showToast('Failed to save sites to WordPress.');
  }
}


  function saveCurrentSite(){
    if (!searchMarker){
      showToast('No landmark to save.');
      return;
    }
    const pos = searchMarker.getPosition();
    const lat = pos.lat();
    const lng = pos.lng();
    const defaultName = searchLastTitle || 'Candidate site';

    const name = prompt('Site name?', defaultName);
    if (!name) return;

    const site = {
      id: Date.now(),
      name: name.trim(),
      lat,
      lng
    };
    savedSites.push(site);
    persistSavedSites();
    rebuildSavedSitesList();
    showToast('Site saved.');
  }

  function rebuildSavedSitesList(){
    const box = document.getElementById('savedSitesList');
    if (!box) return;
    box.innerHTML = '';

    if (!savedSites.length){
      box.innerHTML = '<p style="font-size:12px;color:#666;padding:4px 0;">No saved sites yet.</p>';
      return;
    }

    savedSites.forEach(site => {
      const el = document.createElement('div');
      el.className = 'substation-item';  // reuse same styling
      el.style.cursor = 'pointer';
      el.innerHTML = `
        <div class="substation-name">${escapeHtml(site.name)}</div>
        <div class="substation-meta" style="font-size:11px;color:#555;">
          (${site.lat.toFixed(5)}, ${site.lng.toFixed(5)})
        </div>
        <button class="mini-btn" data-id="${site.id}" style="margin-top:4px;">Delete</button>
      `;
      el.addEventListener('click', (e)=>{
        // avoid click when pressing Delete
        if (e.target && e.target.matches('button.mini-btn')) return;

        const latLng = new google.maps.LatLng(site.lat, site.lng);
        setSearchLandmark(latLng, site.name);
        map.setZoom(18);
        map.panTo(latLng);
      });
      box.appendChild(el);
    });

    // wire delete buttons
    box.querySelectorAll('button.mini-btn').forEach(btn=>{
      btn.addEventListener('click', (e)=>{
        e.stopPropagation();
        const id = Number(btn.getAttribute('data-id'));
        savedSites = savedSites.filter(s => s.id !== id);
        persistSavedSites();
        rebuildSavedSitesList();
      });
    });
  }

  function clearAllSavedSites(){
    if (!savedSites.length) return;
    if (!confirm('Clear all saved candidate sites?')) return;
    savedSites = [];
    persistSavedSites();
    rebuildSavedSitesList();
  }

// --- Main function to place / update the landmark + box ---
function setSearchLandmark(latLng, title = 'Searched location'){
  searchLastTitle = title || 'Searched location';

  if (!searchMarker){
    searchMarker = new google.maps.Marker({
      map,
      position: latLng,
      title: searchLastTitle,
      animation: google.maps.Animation.DROP,
      draggable: true
    });

    // When user drags marker, reverse geocode + update box
    searchMarker.addListener('dragend', (e) => {
      const newPos = e.latLng;
      map.panTo(newPos);

      if (!searchGeocoder){
        searchGeocoder = new google.maps.Geocoder();
      }
      searchGeocoder.geocode({ location: newPos }, (results, status) => {
        let newTitle = searchLastTitle;
        if (status === 'OK' && results && results[0]) {
          newTitle = results[0].formatted_address;
        }
        setSearchLandmark(newPos, newTitle);
      });
    });

  } else {
    searchMarker.setPosition(latLng);
    searchMarker.setTitle(searchLastTitle);
    searchMarker.setDraggable(true);
  }

  const lat = latLng.lat();
  const lng = latLng.lng();
  const point = new google.maps.LatLng(lat, lng);

  // Return nearest substation: name + voltage (if any) + distance in km
  function computeNearestSubstation(){
    if (!Array.isArray(substations) || !substations.length) return null;

    // Try to guess voltage text from capacity / remark / name
    function extractVoltage(sub){
      const parts = [];
      if (sub.capacity != null) parts.push(String(sub.capacity));
      if (sub.remark   != null) parts.push(String(sub.remark));
      if (sub.name     != null) parts.push(String(sub.name));
      const text = parts.join(' ').toLowerCase();

      // Look for 500/275/132/66/33/22/11 kV patterns
      const m = text.match(/\b(500|275|230|220|132|110|66|45|33|22|11)\s*k?v\b/);
      if (m) return `${m[1]}kV`;
      return ''; // no clear voltage found
    }

    let min = Infinity;
    let nearest = null;

    substations.forEach(s => {
      if (s.lat == null || s.lng == null) return;
      const q = new google.maps.LatLng(s.lat, s.lng);
      const d = google.maps.geometry.spherical.computeDistanceBetween(point, q);
      if (d < min){
        min = d;
        nearest = s;
      }
    });

    if (!nearest || min === Infinity) return null;

    return {
      name: nearest.name || 'Unknown Substation',
      voltage: extractVoltage(nearest),      // e.g. "132kV"
      distKm: (min / 1000)
    };
  }



  function computeNearestPolylineKm(lineArrays){
    if (!lineArrays || !lineArrays.length) return null;
    let min = Infinity;
    lineArrays.forEach(line => {
      if (!line || !line.getPath) return;
      const path = line.getPath();
      const len = path.getLength();
      for (let i = 0; i < len; i++){
        const q = path.getAt(i);
        const d = google.maps.geometry.spherical.computeDistanceBetween(point, q);
        if (d < min) min = d;
      }
    });
    return min === Infinity ? null : (min / 1000);
  }

  const nearestSub = computeNearestSubstation();
  const dTLKm   = computeNearestPolylineKm([
    ...tl_500, ...tl_275, ...tl_132, ...tl_33, ...tl_11
  ]);
  const dRailKm = computeNearestPolylineKm(ktmLines);
  const dHwyKm  = computeNearestPolylineKm(hwyLines);

  const distLines = [];
  if (nearestSub){
    const voltPart = nearestSub.voltage ? ` (${nearestSub.voltage})` : '';
    distLines.push(
      `Nearest PMU/Substation: ${nearestSub.name}${voltPart} (${nearestSub.distKm.toFixed(2)} km)`
    );
  }

  if (dTLKm   != null) distLines.push(`Nearest Transmission Line: ${dTLKm.toFixed(2)} km`);
  if (dRailKm != null) distLines.push(`Nearest KTM Rail: ${dRailKm.toFixed(2)} km`);
  if (dHwyKm  != null) distLines.push(`Nearest Highway: ${dHwyKm.toFixed(2)} km`);

  const distHtml = distLines.length
    ? `<div style="margin-top:4px;font-size:11px;color:#444;">${distLines.join('<br>')}</div>`
    : '';

  const labelHtml = `
    <div style="
      font-size:11px;
      line-height:1.25;
      position:relative;
      padding:6px 22px 6px 6px;
      border-radius:10px;
      box-shadow:0 2px 6px rgba(0,0,0,0.25);
      border:1px solid rgba(0,0,0,0.15);
      background:#fff;
      max-width:220px;           /* 👈 Limit width */
      white-space:normal;         /* 👈 Allow sentence wrap */
      word-break:break-word;      /* 👈 No overflow on long names */
    ">
          <div style="margin-bottom:2px;">${searchLastTitle}</div>
          <div style="color:#555;font-size:12px;margin-bottom:4px;">
            (${lat.toFixed(6)}, ${lng.toFixed(6)})
          </div>
          ${distHtml}
              <div style="display:flex;gap:4px;font-size:10px;margin-top:4px;flex-wrap:wrap;">
                <button onclick="navigator.clipboard.writeText('${lat.toFixed(6)}, ${lng.toFixed(6)}')"
                        style="border:1px solid #ccc;border-radius:10px;padding:1px 6px;background:#f5f5f5;cursor:pointer;font-size:10px;">
                  Copy
                </button>
                <button onclick="saveCurrentSite()"
                        style="border:1px solid #1b5e20;border-radius:10px;padding:1px 6px;background:#e6f3ea;color:#1b5e20;cursor:pointer;font-size:10px;">
                  Save
                </button>
                <button onclick="copyShareLink(${lat.toFixed(6)}, ${lng.toFixed(6)})"
                        style="border:1px solid #1565c0;border-radius:10px;padding:1px 6px;background:#e3f2fd;color:#1565c0;cursor:pointer;font-size:10px;">
                  Share
                </button>
                <button onclick="window.open('https://www.google.com/maps?q=${lat.toFixed(6)},${lng.toFixed(6)}','_blank')"
                        style="border:1px solid #0f3f17;border-radius:10px;padding:1px 6px;background:#d4edda;color:#0f3f17;cursor:pointer;font-size:10px;">
                  Open in Google Maps
                </button>
              </div>
            <button onclick="closeSearchInfo()"
                    style="position:absolute;right:4px;top:2px;border:none;background:transparent;font-weight:bold;font-size:12px;cursor:pointer;line-height:1;">
              ×
            </button>
        </div>
      `;
    
      if (!searchIW){
        searchIW = new google.maps.InfoWindow({
          content: labelHtml,
          disableAutoPan: true
        });
      } else {
        searchIW.setContent(labelHtml);
      }
      searchIW.open({ anchor: searchMarker, map });
    }




  let markers = [];
  let substations = [];
  let selectedStates = [];

  // Transmission line overlays
  let tl_33  = [];
  let tl_11  = [];
  let tl_132 = [];
  let tl_275 = [];
  let tl_500 = [];
  let ktmLines = [];
  let hwyLines = [];
  let tl_loaded = false;
  let TL_SEEN_IDS = new Set();

  // Site boundary globals
  let drawingManager;
  let drawnShapes = []; // {id, type:'polygon', gobj, name, owner, rental}
  let activeShape = null;

  // ---------- Helpers ----------
  const byKey = (obj, key) => Object.keys(obj).find(k => k.trim().toLowerCase() === key);
  function escapeHtml(str){
    const d = document.createElement('div');
    d.textContent = str;
    return d.innerHTML;
  }
  function showToast(msg){
    const t=document.getElementById('toast'); t.innerText=msg; t.classList.add('show');
    setTimeout(()=>t.classList.remove('show'), 4200);
  }
  function setStatus(msg){ document.getElementById('importStatus').innerText = msg; }
  function setOsmStatus(msg){
    const el = document.getElementById('osmStatus');
    if(el) el.textContent = msg;
  }

  function clearTransmissionLines(){
    tl_275.forEach(p => p.setMap(null));
    tl_132.forEach(p => p.setMap(null));
    tl_33 .forEach(p => p.setMap(null));
    tl_11 .forEach(p => p.setMap(null));
    tl_500.forEach(p => p.setMap(null));
    ktmLines.forEach(p => p.setMap(null));
    hwyLines.forEach(p => p.setMap(null));

    tl_275.length = tl_132.length = tl_33.length = tl_11.length = 0;
    tl_500.length = 0;
    ktmLines.length = 0;
    hwyLines.length = 0;

    tl_loaded = false;
    TL_SEEN_IDS.clear();

    ['toggle11','toggle33','toggle132','toggle275','toggle500','toggleKTM','toggleHwy']
      .forEach(id => {
        const cb = document.getElementById(id);
        if (cb) { cb.checked = true; cb.indeterminate = false; }
      });

    updateTransmissionVisibility();
    setOsmStatus('Cleared. All layers are OFF.');
  }

function attachPlacesSearch(){
  const input = document.getElementById('mapSearchInput');
  if(!input) return;

  const geocoder = new google.maps.Geocoder();

  function flyTo(latLng, zoom=18, title='Searched location'){
    map.setCenter(latLng);
    map.setZoom(zoom);
    setSearchLandmark(latLng, title);
  }

  function geocodeAndGo(query){
    if(!query) return;

    // 1) Try to interpret the query as "lat,lng" first
    const coord = parseCoordinate(query);
    if (
      coord &&
      coord.lat >= -90 && coord.lat <= 90 &&
      coord.lng >= -180 && coord.lng <= 180
    ) {
      const latLng = new google.maps.LatLng(coord.lat, coord.lng);
      // Zoom in to around building level, you can adjust 18 -> 17/19 if you like
      flyTo(
        latLng,
        18,
        `Lat ${coord.lat.toFixed(6)}, Lng ${coord.lng.toFixed(6)}`
      );
      return; // IMPORTANT: stop here, don’t call geocoder
    }

    // 2) If not lat,lng, fall back to your existing address geocode
    geocoder.geocode({ address: query }, (results, status) => {
      if(status === 'OK' && results && results[0]){
        const loc = results[0].geometry.location;
        flyTo(
          loc,
          results[0].geometry.viewport ? map.getZoom() : 18,
          results[0].formatted_address
        );
        if(results[0].geometry.viewport){
          map.fitBounds(results[0].geometry.viewport);
          google.maps.event.addListenerOnce(map, 'idle', () => {
            if(map.getZoom() < 16) map.setZoom(16);
          });
        }
      }else{
        showToast('Could not find that address.');
      }
    });
  }

  // Coordinate shortcut — capture phase fires BEFORE Places Autocomplete
  input.addEventListener('keydown', function coordCapture(e){
    if (e.key !== 'Enter') return;
    const coord = parseCoordinate(input.value.trim());
    if (coord && coord.lat >= -90 && coord.lat <= 90 && coord.lng >= -180 && coord.lng <= 180){
      e.preventDefault();
      e.stopImmediatePropagation();
      flyTo(new google.maps.LatLng(coord.lat, coord.lng), 18,
        'Lat ' + coord.lat.toFixed(6) + ', Lng ' + coord.lng.toFixed(6));
    }
  }, true); // capture = true → fires before Places Autocomplete

  const ac = new google.maps.places.Autocomplete(input, {
    fields: ['geometry', 'name', 'formatted_address'],
  });
  ac.bindTo('bounds', map);

  ac.addListener('place_changed', ()=>{
    const place = ac.getPlace();
    if(place && place.geometry){
      if(place.geometry.viewport){
        map.fitBounds(place.geometry.viewport);
        google.maps.event.addListenerOnce(map, 'idle', () => {
          if(map.getZoom() < 16) map.setZoom(16);
        });
      }else{
        flyTo(place.geometry.location, 18, place.formatted_address || place.name);
      }
    }else{
      geocodeAndGo(input.value.trim());
    }
  });

  input.addEventListener('keydown', (e)=>{
    if(e.key === 'Enter'){ e.preventDefault(); geocodeAndGo(input.value.trim()); }
  });
}


function attachLocateMe(){
  const btn = document.getElementById('locateMeBtn');
  if (!btn) return;

  // Change icon meaning
  btn.title = 'Search this location';
  btn.disabled = false;
  btn.style.opacity = 1;

  btn.addEventListener('click', () => {
    const input = document.getElementById('mapSearchInput');
    if (!input) return;

    const val = input.value.trim();
    if (!val) {
      showToast('Please type a place, address, or "lat,lng" first.');
      input.focus();
      return;
    }

    // Simulate pressing Enter so it uses the SAME logic as your search box
    const evt = new KeyboardEvent('keydown', { key: 'Enter' });
    input.dispatchEvent(evt);
  });
}


  function setLayerChecks(val){
    ['toggle11','toggle33','toggle132','toggle275','toggle500','toggleKTM','toggleHwy']
      .forEach(id=>{
        const cb = document.getElementById(id);
        if (cb){ cb.checked = val; cb.indeterminate = false; }
      });
  }

  function selectAllLines(){
    setLayerChecks(true);
    if (!tl_loaded) {
      loadTransmissionLines();
    } else {
      updateTransmissionVisibility();
    }
    setOsmStatus('All layers ON.');
  }

  function clearAllLines(){
    setLayerChecks(false);
    updateTransmissionVisibility();
    setOsmStatus('All layers OFF.');
  }

  function updateTransmissionVisibility(){
    const show500 = document.getElementById('toggle500')?.checked ?? true;
    const show275 = document.getElementById('toggle275')?.checked ?? true;
    const show132 = document.getElementById('toggle132')?.checked ?? true;
    const show33  = document.getElementById('toggle33') ?.checked ?? true;
    const show11  = document.getElementById('toggle11') ?.checked ?? true;
    const showKTM = document.getElementById('toggleKTM')?.checked ?? true;
    const showHwy = document.getElementById('toggleHwy')?.checked ?? true;
  
    tl_500.forEach(p => p.setMap(show500 ? map : null));
    tl_275.forEach(p => p.setMap(show275 ? map : null));
    tl_132.forEach(p => p.setMap(show132 ? map : null));
    tl_33 .forEach(p => p.setMap(show33  ? map : null));
    tl_11 .forEach(p => p.setMap(show11  ? map : null));
  
    ktmLines.forEach(p => p.setMap(showKTM ? map : null));
    hwyLines.forEach(p => p.setMap(showHwy ? map : null));
  }

  const GEOJSON_URL = 'https://samaiden-ai.com.my/wp-content/uploads/2026/03/transmission-lines-all2.geojson';

  async function loadTransmissionLines(){
    if (tl_loaded) {
      updateTransmissionVisibility();
      return;
    }
    setOsmStatus('Fetching GeoJSON…');
    try {
      const resp = await fetch(GEOJSON_URL, { cache: 'no-store' });
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const geojson = await resp.json();

      const drawLine = (coords, klass) => {
        let color  = '#C97125';
        let weight = 4.2;
        if (klass === '275') { color = '#00A8FF'; weight = 3.8; }
        else if (klass === '132') { color = '#FF1744'; weight = 3.8; }
        else if (klass === '33')  { color = '#FF9800'; weight = 3.0; }
        else if (klass === '11')  { color = '#9C27B0'; weight = 2.8; }
        else if (klass === 'ktm') { color = '#F0C50E'; weight = 3.0; }
        else if (klass === 'hwy') { color = '#FBC02D'; weight = 2.8; }

        const line = new google.maps.Polyline({
          path: coords.map(([lng, lat]) => ({ lat, lng })),
          geodesic: true,
          clickable: false,
          strokeColor: color,
          strokeOpacity: 1.0,
          strokeWeight: weight,
          map
        });

        if (klass === '500')      tl_500.push(line);
        else if (klass === '275') tl_275.push(line);
        else if (klass === '132') tl_132.push(line);
        else if (klass === '33')  tl_33.push(line);
        else if (klass === '11')  tl_11.push(line);
        else if (klass === 'ktm') ktmLines.push(line);
        else if (klass === 'hwy') hwyLines.push(line);
      };

      let c500 = 0, c275 = 0, c132 = 0, c33 = 0, c11 = 0, cKTM = 0;

      (geojson.features || []).forEach(f => {
        const props = f.properties || {};
        const g = f.geometry;
        if (!g) return;

        const op  = (props.operator || '').toLowerCase();
        const net = (props.network  || '').toLowerCase();
        const hasRail       = !!props.railway;
        const isTrainRoute  = props.route === 'train';
        const isKtmOperator = op.includes('ktm') || op.includes('keretapi tanah melayu');
        const isKtmNetwork  = net.includes('ktm');
        const isKTM = (hasRail || isTrainRoute) && (isKtmOperator || isKtmNetwork);

        if (isKTM) {
          const klass = 'ktm';
          if (g.type === 'LineString') {
            drawLine(g.coordinates, klass);
          } else if (g.type === 'MultiLineString') {
            g.coordinates.forEach(seg => drawLine(seg, klass));
          }
          cKTM++;
          return;
        }

        const rawV = props.voltage ?? props.voltages ?? props.volt ?? '';
        const v = Array.isArray(rawV) ? rawV.join(';') : String(rawV);
        if (!v) return;

        let klass = null;
        if (/\b500(?:\s*kV|000)?\b/i.test(v))      klass = '500';
        else if (/\b275(?:\s*kV|000)?\b/i.test(v)) klass = '275';
        else if (/\b132(?:\s*kV|000)?\b/i.test(v)) klass = '132';
        else if (/\b33(?:\s*kV|000)?\b/i.test(v))  klass = '33';
        else if (/\b11(?:\s*kV|000)?\b/i.test(v))  klass = '11';
        else return;

        if (g.type === 'LineString') {
          drawLine(g.coordinates, klass);
        } else if (g.type === 'MultiLineString') {
          g.coordinates.forEach(seg => drawLine(seg, klass));
        } else {
          return;
        }

        if (klass === '500')      c500++;
        else if (klass === '275') c275++;
        else if (klass === '132') c132++;
        else if (klass === '33')  c33++;
        else if (klass === '11')  c11++;
      });

      tl_loaded = true;
      updateTransmissionVisibility();

      setOsmStatus(
        `${c500} (500kV), ${c275} (275kV), ${c132} (132kV), ${c33} (33kV), ${c11} (11kV), ${cKTM} KTM segments.`
      );
    } catch (err) {
      console.error(err);
      setOsmStatus('Failed to load GeoJSON.');
    }
  }

  async function onLayerToggleChange() {
    const anyOn = [
      'toggle11','toggle33','toggle132','toggle275','toggle500',
      'toggleKTM','toggleHwy'
    ].some(id => document.getElementById(id)?.checked);
    if (anyOn && !tl_loaded) {
      await loadTransmissionLines();
    } else {
      updateTransmissionVisibility();
    }
  }

  function parseCoordinate(coord){
    if(!coord) return null;
    const parts = String(coord).split(',');
    if(parts.length<2) return null;
    const lat = parseFloat(parts[0].trim());
    const lng = parseFloat(parts[1].trim());
    if(Number.isNaN(lat) || Number.isNaN(lng)) return null;
    return {lat, lng};
  }

  // ---------- Excel/CSV Import ----------
  async function handleFile(file){
    setStatus(`Reading ${file.name} …`);
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type:'array', dense:true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    let rows = XLSX.utils.sheet_to_json(ws, { defval:"" });

    if(rows.length === 0){ showToast('No rows found in the first sheet.'); setStatus(''); return; }

    const sample = rows[0];
    const kState   = byKey(sample,'state');
    const kDist    = byKey(sample,'district');
    const kName    = byKey(sample,'substation name');
    const kCoord   = byKey(sample,'coordinate');
    const kCap    = byKey(sample,'capacity');
    const kRemark = byKey(sample,'remark');

    if(!kState || !kName || !kCoord){
      showToast('Missing required columns. Need: State, Substation Name, Coordinate. (District optional)');
      setStatus('');
      return;
    }

    const parsed = [];
    let bad = 0;
    rows.forEach(r=>{
      const state = String(r[kState]).trim();
      const district = kDist ? String(r[kDist]).trim() : '';
      const name = String(r[kName]).trim();
      const coord = parseCoordinate(r[kCoord]);
      const capRaw = kCap ? String(r[kCap]).trim() : '';
      const remark = kRemark ? String(r[kRemark]).trim() : '';
      const cap = capRaw && !Number.isNaN(Number(capRaw)) ? Number(capRaw) : capRaw;
      if(!state || !name || !coord){ bad++; return; }
      parsed.push({ state, district, name, lat: coord.lat, lng: coord.lng, capacity: cap, remark });
    });

    substations = parsed;
    if(substations.length === 0){ showToast('Couldn’t find any valid rows (check coordinates format).'); setStatus(''); return; }

    setStatus(`Imported ${substations.length} rows${bad?` (${bad} skipped as invalid)`:''}.`);
    rebuildUIFromData();
  }

  async function handleUrl(url){
    if(!url){ showToast('Please paste a Media Library file URL.'); return; }
    try{
      setStatus('Fetching file from URL …');
      const resp = await fetch(url, { credentials: 'omit' });
      if(!resp.ok){
        showToast(`Failed to fetch file (${resp.status}). Check that the URL is public and correct.`);
        setStatus('');
        return;
      }
      const data = await resp.arrayBuffer();
      const wb = XLSX.read(data, { type:'array', dense:true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      let rows = XLSX.utils.sheet_to_json(ws, { defval:"" });
  
      if(rows.length === 0){ showToast('No rows found in the first sheet.'); setStatus(''); return; }
  
      const sample = rows[0];
      const kState   = byKey(sample,'state');
      const kDist    = byKey(sample,'district');
      const kName    = byKey(sample,'substation name');
      const kCoord   = byKey(sample,'coordinate');
      const kCap     = byKey(sample,'capacity');
      const kRemark  = byKey(sample,'remark');
  
      if(!kState || !kName || !kCoord){
        showToast('Missing required columns. Need: State, Substation Name, Coordinate. (District optional)');
        setStatus('');
        return;
      }
  
      const parsed = [];
      let bad = 0;
      rows.forEach(r=>{
        const state    = String(r[kState]).trim();
        const district = kDist ? String(r[kDist]).trim() : '';
        const name     = String(r[kName]).trim();
        const coord    = parseCoordinate(r[kCoord]);
        const capRaw   = kCap ? String(r[kCap]).trim() : '';
        const remark   = kRemark ? String(r[kRemark]).trim() : '';
        const cap      = capRaw && !Number.isNaN(Number(capRaw)) ? Number(capRaw) : capRaw;
        if(!state || !name || !coord){ bad++; return; }
        parsed.push({ state, district, name, lat: coord.lat, lng: coord.lng, capacity: cap, remark });
      });
  
      substations = parsed;
      if(substations.length === 0){
        showToast('Couldn’t find any valid rows (check coordinates format).');
        setStatus('');
        return;
      }
  
      setStatus(`Imported ${substations.length} rows from URL${bad?` (${bad} skipped)`:''}.`);
      rebuildUIFromData();
    }catch(err){
      console.error(err);
      showToast('Error loading file from URL. Open console for details.');
      setStatus('');
    }
  }

  function rebuildUIFromData(){
    selectedStates = [...new Set(substations.map(s=>s.state))];
    buildStateFilter();
    updateMarkers();
    updateSubstationList();
    updateStats();
    fitToMarkers();
  }

  function buildStateFilter(){
    const container = document.getElementById('stateDropdownBody');
    const label = document.getElementById('stateDropdownLabel');

    container.innerHTML = '';
    const states = [...new Set(substations.map(s=>s.state))].sort();

    states.forEach(state=>{
      const id = 'state-' + CSS.escape(state);
      const row = document.createElement('label');
      row.innerHTML = `
        <input type="checkbox" id="${id}" value="${state}" checked>
        ${state}
      `;
      row.querySelector('input').addEventListener('change', function(){
        if(this.checked)
          selectedStates.push(this.value);
        else
          selectedStates = selectedStates.filter(s => s !== this.value);
        updateMarkers();
        updateSubstationList();
        updateStats();
        if(selectedStates.length === states.length)
          label.textContent = 'All States Selected';
        else
          label.textContent = selectedStates.join(', ');
      });
      container.appendChild(row);
    });

    document.getElementById('stateDropdownHeader')
      .addEventListener('click', () => {
        document.getElementById('stateDropdown')
          .classList.toggle('open');
      });
  }

  function selectAllStates(){
    const checkboxes = document.querySelectorAll(
      '#stateDropdownBody input[type="checkbox"]'
    );
    const states = [...new Set(substations.map(s => s.state))].sort();
    checkboxes.forEach(cb => cb.checked = true);
    selectedStates = [...states];
    const label = document.getElementById('stateDropdownLabel');
    if (label) label.textContent = 'All States Selected';
    updateMarkers();
    updateSubstationList();
    updateStats();
    fitToMarkers();
  }
    
  function clearAllStates(){
    const checkboxes = document.querySelectorAll(
      '#stateDropdownBody input[type="checkbox"]'
    );
    checkboxes.forEach(cb => cb.checked = false);
    selectedStates = [];
    const label = document.getElementById('stateDropdownLabel');
    if (label) label.textContent = 'No State Selected';
    updateMarkers();
    updateSubstationList();
    updateStats();
  }

  function filterSubstations(){
    const term = document.getElementById('searchInput').value.toLowerCase();
    updateSubstationList(term);
  }

  function updateMarkers(){
    markers.forEach(m=>m.setMap(null)); markers = [];
    const filtered = substations.filter(s=>selectedStates.includes(s.state));
    filtered.forEach(sub=>{
      const marker = new google.maps.Marker({
        position:{lat:sub.lat,lng:sub.lng},
        map,
        title: sub.name,
        icon:{
          url:'data:image/svg+xml;base64,'+btoa(`
            <svg width="22" height="22" viewBox="0 0 22 22" xmlns="http://www.w3.org/2000/svg">
              <circle cx="11" cy="11" r="9" fill="#D81B60" stroke="white" stroke-width="2"/> 
            </svg>`),
          scaledSize:new google.maps.Size(15,15),
          anchor:new google.maps.Point(11,11)
        }
      });
      marker.addListener('click', ()=>{
        if (infoWindow) infoWindow.close();
        infoWindow.setContent(`
          <div class="iw-compact" style="max-width:260px;">
            <button class="iw-close" onclick="closePMUInfo()" aria-label="Close">&times;</button>
            <h3>${sub.name}</h3>
            <p><strong>State:</strong> ${sub.state}</p>
            ${sub.district ? `<p><strong>District:</strong> ${sub.district}</p>` : ''}
            <p><strong>Coordinates:</strong> ${sub.lat.toFixed(6)}, ${sub.lng.toFixed(6)}</p>
            ${typeof sub.capacity !== 'undefined' && String(sub.capacity).trim() !== '' ? 
              `<p><strong>Capacity:</strong> ${typeof sub.capacity === 'number' ? sub.capacity.toLocaleString() : sub.capacity}</p>` : ''
            }
            ${sub.remark ? `<p><strong>Remark:</strong> ${sub.remark}</p>` : ''}
          </div>
        `);
        infoWindow.open(map, marker);
      });
      markers.push(marker);
    });
    updateStats();
  }

  function fitToMarkers(){
    if(!map || markers.length===0) return;
    const b = new google.maps.LatLngBounds();
    markers.forEach(m=>b.extend(m.getPosition()));
    const leftPad = window.innerWidth <= 768 ? 16 : 412;
    map.fitBounds(b, { top:40, right:16, bottom:40, left:leftPad });
  }

  function updateSubstationList(searchTerm=''){
    const box = document.getElementById('substationList');
    box.innerHTML = '';
    let list = substations.filter(s=>selectedStates.includes(s.state));
    if (searchTerm){
      const t = searchTerm.toLowerCase();
      list = list.filter(s =>
        s.name.toLowerCase().includes(t) ||
        s.state.toLowerCase().includes(t) ||
        (s.district||'').toLowerCase().includes(t) ||
        (s.remark||'').toLowerCase().includes(t) ||
        (s.capacity!==undefined ? String(s.capacity).toLowerCase().includes(t) : false)
      );
    }
    if(list.length===0){
      box.innerHTML = '<p style="text-align:center;padding:16px;color:#666;">No substations match the current filters</p>';
      return;
    }
    list.forEach(sub=>{
      const el = document.createElement('div');
      el.className = 'substation-item';
      el.innerHTML = `
        <div class="substation-name">${sub.name}</div>
        <div class="substation-location"><i class="fas fa-location-dot"></i> ${sub.district?sub.district+', ':''}${sub.state}</div>`;
      el.addEventListener('click', ()=>{
        map.setCenter({lat:sub.lat,lng:sub.lng}); map.setZoom(17);
        const m = markers.find(mm => mm.getPosition().lat()===sub.lat && mm.getPosition().lng()===sub.lng);
        if(m) google.maps.event.trigger(m,'click');
      });
      box.appendChild(el);
    });
  }

  function updateStats(){
    const total = document.getElementById('totalSubstations');
    const visible = document.getElementById('visibleSubstations');
    const states = document.getElementById('selectedStates');
    if (total)   total.textContent   = substations.length;
    if (visible) visible.textContent = substations.filter(s=>selectedStates.includes(s.state)).length;
    if (states)  states.textContent  = selectedStates.length;
  }

  // ==========================================================================
  // DATA SOURCE — Google Sheets (published as CSV)
  // --------------------------------------------------------------------------
  // To update substation data, just edit the Google Sheet in your browser.
  // Changes go live within ~5 minutes — no re-upload, no HTML edit needed.
  //
  // To set this up (one-time):
  //   1. Open your Google Sheet
  //   2. File → Share → Publish to web
  //   3. Choose the sheet tab, format = "Comma-separated values (.csv)"
  //   4. Keep "Automatically republish when changes are made" CHECKED
  //   5. Click Publish, copy the URL, paste it below
  //
  // Required column headers (row 1):
  //   State | District | Substation Name | Coordinate | Capacity | Remark
  //   (District, Capacity, Remark are optional)
  //
  // Coordinate format: "lat,lng"  e.g.  3.1390,101.6869
  // ==========================================================================
  const MEDIA_FILE_URL = 'https://samaiden-ai.com.my/wp-content/uploads/2026/05/tnb-substation-peninsular-malaysia.xlsx';

  // ---------- Site Boundary helpers ----------
  function nextId(){ return 'shape_'+Math.random().toString(36).slice(2,9); }

  function setActiveShape(shape){
    if (activeShape && activeShape.gobj) activeShape.gobj.set('zIndex', 1);
    activeShape = shape || null;
    if (activeShape && activeShape.gobj) activeShape.gobj.set('zIndex', 1000);
    updateShapeMetrics();
    msmUpdateSolarPanel(activeShape);
    if (typeof msmDrawConnLines === 'function') msmDrawConnLines(activeShape);
  }

  function polygonMetrics(poly){
    if (!poly || !poly.getPath) return null;
    const path = poly.getPath();
    if (path.getLength() < 3) return null;
    const pts = [];
    for (let i = 0; i < path.getLength(); i++){
      pts.push(path.getAt(i));
    }
    const areaM2 = google.maps.geometry.spherical.computeArea(pts);
    const areaAc = areaM2 / 4046.8564224;
    let sumLat = 0, sumLng = 0;
    pts.forEach(pt => {
      sumLat += pt.lat();
      sumLng += pt.lng();
    });
    const centroid = {
      lat: sumLat / pts.length,
      lng: sumLng / pts.length,
    };
    return { areaM2, areaAc, centroid };
  }

function autoFillStateCityForShape(shape){
  if (!shape || !shape.gobj || !shapeGeocoder) return;

  const metrics = polygonMetrics(shape.gobj);
  if (!metrics) return;

  const center = new google.maps.LatLng(metrics.centroid.lat, metrics.centroid.lng);

  shapeGeocoder.geocode({ location: center }, (results, status) => {
    if (status !== 'OK' || !results || !results[0]) return;

    const comps = results[0].address_components || [];
    let state = '';
    let city  = '';

    comps.forEach(c => {
      if (c.types.includes('administrative_area_level_1')) {
        // e.g. "Selangor", "Johor"
        state = c.long_name;
      }
      if (
        c.types.includes('locality') ||
        c.types.includes('administrative_area_level_2') ||
        c.types.includes('sublocality')
      ) {
        if (!city) city = c.long_name;
      }
    });

    // Update shape object
    if (state) shape.state = state;
    if (city)  shape.city  = city;

    // Re-render the table so the inputs show the new values
    refreshBoundarySummary();
  });
}

  function updateShapeMetrics(){
    const el = document.getElementById('shapeStatus');
    if (!el) return;
    if (!activeShape || !activeShape.gobj){
      el.textContent = 'No shape selected.';
      return;
    }
    const metrics = polygonMetrics(activeShape.gobj);
    if (!metrics){
      el.textContent = 'Draw at least 3 points to get area.';
      return;
    }
    const name = activeShape.name || '(unnamed)';
    el.textContent =
      `${name} — Area: ${metrics.areaAc.toFixed(4)} acres `
      + `(${metrics.areaM2.toFixed(0)} m²); `
      + `Centroid: ${metrics.centroid.lat.toFixed(6)}, `
      + `${metrics.centroid.lng.toFixed(6)}`;
  }

function refreshBoundarySummary(){
  const tbody = document.getElementById('boundarySummaryBody');
  if (!tbody) return;

  // No data
  if (!drawnShapes.length){
    tbody.innerHTML =
      '<tr><td colspan="13" style="padding:4px;font-size:12px;color:#666;">' +
      'No lands drawn yet.' +
      '</td></tr>';
    return;
  }

  const rows = drawnShapes.map((s, index) => {
    const metrics = polygonMetrics(s.gobj);

    const safeName  = (s.name  || '').replace(/"/g, '&quot;');
    const safeOwner = (s.owner || '').replace(/"/g, '&quot;');
    const safeRent  = (s.rental|| '').replace(/"/g, '&quot;');
    const safeState = (s.state || '').replace(/"/g, '&quot;');
    const safeCity  = (s.city  || '').replace(/"/g, '&quot;');
    const safeStatus   = s.status   || 'Prospecting';
    const safeCapacity = s.capacity || '';

    // When metrics missing
    if (!metrics){
      return `
        <tr data-shape-id="${s.id}">
          <td>${index + 1}</td>
          <td><span class="status-badge status-${safeStatus}">${safeStatus}</span></td>
          <td style="text-align:right;">${safeCapacity}</td>
          <td><input type="text" class="land-name-input"
                     data-shape-id="${s.id}"
                     value="${safeName}"
                     style="width:100%;font-size:12px;"></td>
          <td><input type="text" class="land-owner-input"
                     data-shape-id="${s.id}"
                     value="${safeOwner}"
                     style="width:100%;font-size:12px;"></td>
          <td><input type="text" class="land-state-input"
                     data-shape-id="${s.id}"
                     value="${safeState}"
                     style="width:100%;font-size:12px;"></td>
          <td><input type="text" class="land-city-input"
                     data-shape-id="${s.id}"
                     value="${safeCity}"
                     style="width:100%;font-size:12px;"></td>
          <td><input type="text" class="land-rent-input"
                     data-shape-id="${s.id}"
                     value="${safeRent}"
                     style="width:100%;font-size:12px;"></td>
          <td></td>
          <td></td>
          <td style="font-size:11px;color:var(--ink);max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${s.nearestSub||''}">
            ${s.nearestSub || '—'}
            ${s.nearestSubDist ? '<br><small style=\"color:var(--muted)\";white-space:nowrap>' + s.nearestSubDist + ' km</small>' : ''}
          </td>
          <td style="font-size:11px;">${s.connVoltage ? s.connVoltage + '<br><small style=\"color:var(--muted)\">' + s.connDistKm + ' km</small>' : '—'}</td>
          <td style="text-align:center;">
            <button type="button"
                    class="mini-btn delete-shape-btn"
                    data-shape-id="${s.id}"
                    title="Delete this land">
              <i class="fas fa-trash-alt"></i>
            </button>
          </td>
        </tr>`;
    }

    // With metrics
    return `
      <tr data-shape-id="${s.id}">
        <td>${index + 1}</td>
        <td><span class="status-badge status-${safeStatus}">${safeStatus}</span></td>
        <td style="text-align:right;">${safeCapacity}</td>
        <td>
          <input type="text"
                 class="land-name-input"
                 data-shape-id="${s.id}"
                 value="${safeName}"
                 style="width:100%;font-size:12px;">
        </td>
        <td>
          <input type="text"
                 class="land-owner-input"
                 data-shape-id="${s.id}"
                 value="${safeOwner}"
                 style="width:100%;font-size:12px;">
        </td>
        <td>
          <input type="text"
                 class="land-state-input"
                 data-shape-id="${s.id}"
                 value="${safeState}"
                 style="width:100%;font-size:12px;">
        </td>
        <td>
          <input type="text"
                 class="land-city-input"
                 data-shape-id="${s.id}"
                 value="${safeCity}"
                 style="width:100%;font-size:12px;">
        </td>
        <td>
          <input type="text"
                 class="land-rent-input"
                 data-shape-id="${s.id}"
                 value="${safeRent}"
                 style="width:100%;font-size:12px;">
        </td>
        <td style="text-align:right;">${metrics.areaAc.toFixed(4)}</td>
        <td style="text-align:right;">${metrics.areaM2.toFixed(0)}</td>
        <td style="font-size:11px;color:var(--ink);max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${s.nearestSub||''}">
          ${s.nearestSub || '—'}
          ${s.nearestSubDist ? '<br><small style=\"color:var(--muted);white-space:nowrap\">' + s.nearestSubDist + ' km</small>' : ''}
        </td>
        <td style="font-size:11px;text-align:center;">${s.connVoltage ? '<span style=\"font-weight:600;color:var(--brand-deep)\">' + s.connVoltage + '</span><br><small style=\"color:var(--muted)\">' + s.connDistKm + ' km</small>' : '—'}</td>
        <td style="text-align:center;">
          <button type="button"
                  class="mini-btn delete-shape-btn"
                  data-shape-id="${s.id}"
                  title="Delete this land">
            <i class="fas fa-trash-alt"></i>
          </button>
        </td>
      </tr>`;
  });

  tbody.innerHTML = rows.join('');

  // Rebuild filter dropdowns and reapply active filters after every render
  if (typeof msmBuildSavedLandFilters === 'function')         msmBuildSavedLandFilters();
  if (typeof msmApplySavedLandDropdownFilter === 'function')  msmApplySavedLandDropdownFilter();

  // INPUT CHANGE HANDLERS (keep what you already had)
  tbody.querySelectorAll('.land-name-input').forEach(input => {
    input.addEventListener('change', () => {
      const shape = drawnShapes.find(s => s.id === input.dataset.shapeId);
      if (shape){
        shape.name = input.value.trim();
        updateShapeMetrics();
      }
    });
  });

  tbody.querySelectorAll('.land-owner-input').forEach(input => {
    input.addEventListener('change', () => {
      const shape = drawnShapes.find(s => s.id === input.dataset.shapeId);
      if (shape){
        shape.owner = input.value.trim();
      }
    });
  });

  tbody.querySelectorAll('.land-state-input').forEach(input => {
    input.addEventListener('change', () => {
      const shape = drawnShapes.find(s => s.id === input.dataset.shapeId);
      if (shape){
        shape.state = input.value.trim();
      }
    });
  });

  tbody.querySelectorAll('.land-city-input').forEach(input => {
    input.addEventListener('change', () => {
      const shape = drawnShapes.find(s => s.id === input.dataset.shapeId);
      if (shape){
        shape.city = input.value.trim();
      }
    });
  });

  tbody.querySelectorAll('.land-rent-input').forEach(input => {
    input.addEventListener('change', () => {
      const shape = drawnShapes.find(s => s.id === input.dataset.shapeId);
      if (shape){
        shape.rental = input.value.trim();
      }
    });
  });

  // Attach click handler only once
  if (!tbody.dataset.boundaryClickBound) {
    tbody.addEventListener('click', (e) => {
      // 1) Delete button
      const delBtn = e.target.closest('.delete-shape-btn');
        if (delBtn) {
          const id = delBtn.getAttribute('data-shape-id');
          if (!id) return;
        
          // 1) Find the shape for this row
          const shape = drawnShapes.find(s => s.id === id);
          if (!shape) return;

          // Guard: must be unlocked
          if (!msmIsUnlocked()) {
            showToast('Unlock editing to delete.');
            e.stopPropagation(); return;
          }

          // 2) Require the row/shape to be selected first
          if (!activeShape || activeShape.id !== id) {
            // You already use showToast elsewhere; if not, replace with alert(...)
            showToast('Please click this land row first to select it, then click delete again.');
            e.stopPropagation();
            return;
          }

          // 3) SECOND-LAYER PROTECTION: require separate delete passcode
          e.stopPropagation();
          msmShowDeleteModal(() => {
            // This runs ONLY after the delete passcode is verified
            // Re-check the shape still exists (user may have changed state)
            const currentShape = drawnShapes.find(s => s.id === id);
            if (!currentShape) {
              showToast('Shape no longer exists.');
              return;
            }

            // 4) Remove from map
            if (currentShape.gobj) {
              currentShape.gobj.setMap(null);
            }

            // 5) Remove from drawnShapes list
            const idx = drawnShapes.findIndex(s => s.id === id);
            if (idx !== -1) {
              drawnShapes.splice(idx, 1);
            }

            // 6) Clear active selection, refresh table, and save
            setActiveShape(null);
            refreshBoundarySummary();
            if (typeof saveToWordPress === 'function') {
              saveToWordPress();
            }
            showToast('Land boundary deleted.');
          });
          return;
        }


      // 2) Row click → select/zoom
      const row = e.target.closest('tr[data-shape-id]');
      if (!row) return;
      const id = row.getAttribute('data-shape-id');
      const shape = drawnShapes.find(s => s.id === id);
      if (!shape) return;
      setActiveShape(shape);
      const metrics = polygonMetrics(shape.gobj);
      if (metrics && window.map) {
        map.setCenter(metrics.centroid);
        map.setZoom(16);
      }
    });

    tbody.dataset.boundaryClickBound = '1';
  }
}



  document.addEventListener('DOMContentLoaded', ()=>{
    const tbody = document.getElementById('boundarySummaryBody');
    if (!tbody) return;

tbody.addEventListener('input', (e)=>{
  const target = e.target;
  const id = target.getAttribute('data-shape-id');
  if (!id) return;
  const shape = drawnShapes.find(s => s.id === id);
  if (!shape) return;

  if (target.classList.contains('land-name-input')) {
    shape.name = target.value;
  } else if (target.classList.contains('land-owner-input')) {
    shape.owner = target.value;
  } else if (target.classList.contains('land-state-input')) {
    shape.state = target.value;
  } else if (target.classList.contains('land-city-input')) {
    shape.city = target.value;
  } else if (target.classList.contains('land-rent-input')) {
    shape.rental = target.value;
  }

  // Debounced save — avoids hammering the API on every keystroke
  if (typeof scheduleShapeSave === 'function') {
    scheduleShapeSave();
  }
});

  });



  function rectToPolygon(rect){
    const b = rect.getBounds();
    const ne = b.getNorthEast();
    const sw = b.getSouthWest();
    const nw = new google.maps.LatLng(ne.lat(), sw.lng());
    const se = new google.maps.LatLng(sw.lat(), ne.lng());
    rect.setMap(null);
    const poly = new google.maps.Polygon({
      paths: [nw, ne, se, sw],
      map,
      clickable: true,
      editable: true,
      strokeColor: '#00c853',
      strokeOpacity: 1,
      strokeWeight: 3,
      fillColor: '#b9f6ca',
      fillOpacity: 0.45,
    });
    return poly;
  }

let __shapeSaveTimer = null;

function scheduleShapeSave(){
  // Debounce to avoid spamming WP REST while dragging vertices
  clearTimeout(__shapeSaveTimer);
  __shapeSaveTimer = setTimeout(() => {
    if (typeof saveToWordPress === 'function') saveToWordPress();
  }, 800);
}


function attachShapeEvents(entry){
  const g = entry.gobj;
  g.addListener('click', ()=>{
    setActiveShape(entry);
    refreshBoundarySummary();
  });
  ['insert_at', 'remove_at', 'set_at'].forEach(evt=>{
    g.getPath().addListener(evt, ()=>{
      refreshBoundarySummary();
      updateShapeMetrics();
      scheduleShapeSave();
      // Optional: re-geocode when shape changes (might hit quota if you edit a lot)
      // autoFillStateCityForShape(entry);
    });
  });
}

    function addPolygonFromPath(
      path,
      name     = '',
      owner    = '',
      rental   = '',
      state    = '',
      city     = '',
      status   = 'Prospecting',
      capacity = '',
      tenure   = '',
      flood    = 'Unknown',
      terrain  = 'Unknown',
      notes    = '',
      connVoltage   = '',
      connDistKm    = '',
      nearestSub    = '',
      nearestSubDist = '',
      isNew = false
    ){
      const poly = new google.maps.Polygon({
        paths: path,
        clickable: true,
        editable: true,
        strokeColor: '#00c853',
        strokeOpacity: 1,
        strokeWeight: 3,
        fillColor: '#b9f6ca',
        fillOpacity: 0.45
      });
      poly.setMap(map);
      const shapeId = 'shape-' + Date.now() + '-' + Math.floor(Math.random()*1e6);
      const shape = {
        id: shapeId,
        gobj: poly,
        name:     name     || '',
        owner:    owner    || '',
        rental:   rental   || '',
        state:    state    || '',
        city:     city     || '',
        status:   status   || 'Prospecting',
        capacity: capacity || '',
        tenure:   tenure   || '',
        flood:    flood    || 'Unknown',
        terrain:  terrain  || 'Unknown',
        notes:    notes    || '',
        connVoltage:    connVoltage    || '',
        connDistKm:     connDistKm     || '',
        nearestSub:     nearestSub     || '',
        nearestSubDist: nearestSubDist || ''
      };
      drawnShapes.push(shape);
      msmApplyStatusStyle(shape);
      attachShapeEvents(shape);

      if (isNew) {
        // Only run these for freshly drawn shapes — NOT during batch load from WordPress.
        // During batch load (isNew = false) we skip per-shape UI updates to avoid:
        //  • setActiveShape firing 13 times and panning the map to the last shape's connections
        //  • 13 unnecessary geocoder requests for state/city that are already saved
        //  • refreshBoundarySummary rebuilding the table on every single shape
        refreshBoundarySummary();
        setActiveShape(shape);
        autoFillStateCityForShape(shape);
        autoDetectTerrain(shape);
      }
    }



// Auto-detect terrain slope using Google Elevation API
// Samples 9 points across the polygon bounding box, computes max elevation diff
function autoDetectTerrain(shape) {
  if (!shape || !shape.gobj || !window.msmElevationService) return;
  const path = shape.gobj.getPath().getArray();
  if (path.length < 3) return;

  // Bounding box
  let minLat = Infinity, maxLat = -Infinity, minLng = Infinity, maxLng = -Infinity;
  path.forEach(p => {
    minLat = Math.min(minLat, p.lat()); maxLat = Math.max(maxLat, p.lat());
    minLng = Math.min(minLng, p.lng()); maxLng = Math.max(maxLng, p.lng());
  });

  // 16 sample points in a 4×4 grid across bounding box
  const locations = [];
  for (let r = 0; r < 4; r++) {
    for (let c = 0; c < 4; c++) {
      locations.push({
        lat: minLat + (maxLat - minLat) * (r / 3),
        lng: minLng + (maxLng - minLng) * (c / 3)
      });
    }
  }

  window.msmElevationService.getElevationForLocations({ locations }, (results, status) => {
    if (status !== 'OK' || !results || results.length < 2) return;

    // For every pair of sample points, compute the actual slope between them.
    // Take the steepest pair — this correctly handles ridges and local slopes.
    let maxSlopeDeg = 0;
    for (let i = 0; i < results.length; i++) {
      for (let j = i + 1; j < results.length; j++) {
        const elevDiff = Math.abs(results[i].elevation - results[j].elevation); // metres
        if (elevDiff < 0.5) continue; // ignore sub-0.5m noise

        // Horizontal distance between the two sample points (haversine, metres)
        const distKm = msmHaversineKm(
          results[i].location.lat(), results[i].location.lng(),
          results[j].location.lat(), results[j].location.lng()
        );
        const distM = distKm * 1000;
        if (distM < 1) continue; // skip coincident points

        const slopeDeg = Math.atan(elevDiff / distM) * (180 / Math.PI);
        if (slopeDeg > maxSlopeDeg) maxSlopeDeg = slopeDeg;
      }
    }

    // Classify by max pairwise slope
    // Flat < 3°, Gentle Slope 3–15°, Steep > 15°
    let terrain = 'Flat';
    if (maxSlopeDeg >= 15) terrain = 'Steep';
    else if (maxSlopeDeg >= 3) terrain = 'Gentle Slope';

    shape.terrain = terrain;

    // Update panel dropdown if this shape is currently active
    if (activeShape && activeShape.id === shape.id) {
      const sel = document.getElementById('shapeTerrainInput');
      if (sel) sel.value = terrain;
      msmUpdateSolarPanel(shape);
    }

    refreshBoundarySummary();
    showToast('Terrain: ' + terrain + ' (max slope ' + maxSlopeDeg.toFixed(1) + '°)');
  });
}

  function startDraw(mode){
    if (!drawingManager) return;
    if (!msmIsUnlocked()){
      showToast('Editing is locked. Click the lock icon to unlock.');
      return;
    }
    drawingManager.setDrawingMode(mode === 'polygon'
      ? google.maps.drawing.OverlayType.POLYGON
      : google.maps.drawing.OverlayType.RECTANGLE);
  }

// Internal helper: clears shapes from map + table
function internalClearShapes(options){
  const opts = options || {};
  const clearStorage = !!opts.clearStorage;

  drawnShapes.forEach(s => s.gobj.setMap(null));
  drawnShapes = [];
  setActiveShape(null);
  refreshBoundarySummary();

  // Shape data is stored in WordPress only — no localStorage to clear
}


function clearShapes(){
  if (!msmIsUnlocked()){
    showToast('Editing is locked. Click the lock icon to unlock.');
    return;
  }
  // Only allow clearing if a shape is selected
  if (!activeShape) {
    showToast('No site boundary selected to clear.');
    return;
  }

  const ok = confirm(
    'This will remove the selected site boundary from the map and summary.\n' +
    'Click Save if you also want to update your saved records in WordPress.\n\n' +
    'Do you want to continue?'
  );
  if (!ok) return;

  // Remove the selected shape from the map
  if (activeShape.gobj) {
    activeShape.gobj.setMap(null);
  }

  // Remove it from the drawnShapes list
  drawnShapes = drawnShapes.filter(s => s !== activeShape);

  // Clear active selection and refresh UI
  setActiveShape(null);        // also updates metrics & "No shape selected" text
  refreshBoundarySummary();    // rebuilds the "Saved Lands Summary" table

  // Optional: immediately persist the change to WordPress
  if (typeof saveToWordPress === 'function') {
    saveToWordPress();
  }

  showToast('Selected site boundary cleared.');
}




  const SHAPE_API_SAVE_URL = window.SHAPE_API_SAVE_URL || '';
  const SHAPE_API_NONCE    = window.SHAPE_API_NONCE    || '';

  function toGeoJSON(){
    const features = drawnShapes.map(s => {
      const path = s.gobj.getPath().getArray().map(p => [p.lng(), p.lat()]);
      return {
        type: 'Feature',
        properties: {
          name:           s.name     || '',
          landOwner:      s.owner    || '',
          rentalFee:      s.rental   || '',
          state:          s.state    || '',
          city:           s.city     || '',
          status:         s.status   || 'Prospecting',
          capacityMwp:    s.capacity || '',
          tenure:         s.tenure   || '',
          floodZone:      s.flood    || '',
          terrain:        s.terrain  || '',
          notes:          s.notes    || '',
          connVoltage:    s.connVoltage    || '',
          connDistKm:     s.connDistKm     || '',
          nearestSub:     s.nearestSub     || '',
          nearestSubDist: s.nearestSubDist || ''
        },

        geometry: {
          type: 'Polygon',
          coordinates: [ path ]
        }
      };
    });
    return { type:'FeatureCollection', features };
  }
  
function buildSummaryPayload(){
  const summary = [];

  drawnShapes.forEach((s, index) => {
    const metrics = polygonMetrics(s.gobj);

    let areaAc    = '';
    let areaM2    = '';
    let centerLat = '';
    let centerLng = '';

    if (metrics){
      areaAc    = metrics.areaAc.toFixed(4);
      areaM2    = metrics.areaM2.toFixed(0);
      centerLat = metrics.centroid.lat.toFixed(6);
      centerLng = metrics.centroid.lng.toFixed(6);
    }

    summary.push({
      index: index + 1,
      status: s.status || 'Prospecting',
      name: s.name || '',
      land_owner: s.owner || '',
      state: s.state || '',
      city: s.city || '',
      rental_fee_rm: s.rental || '',
      capacity_mwp: s.capacity || '',
      tenure: s.tenure || '',
      flood_zone: s.flood || '',
      terrain: s.terrain || '',
      nearest_substation: s.nearestSub || '',
      nearest_sub_dist_km: s.nearestSubDist || '',
      connection_voltage: s.connVoltage || '',
      connection_dist_km: s.connDistKm || '',
      notes: s.notes || '',
      area_acres: areaAc,
      area_m2: areaM2,
      center_lat: centerLat,
      center_lng: centerLng
    });

  });

  return summary;
}



async function saveToWordPress(){
  // Check unlock first before doing any work
  if (!msmIsUnlocked()){
    showToast('Editing is locked. Click the lock icon to unlock.');
    return;
  }

  const fc      = toGeoJSON();           // full geometry
  const summary = buildSummaryPayload(); // table rows

  // If no WP endpoint configured, warn and stop
  if (!window.SHAPE_API_SAVE_URL){
    showToast('ERROR: WordPress save URL not configured. Please contact admin.');
    return;
  }

  try {
    showToast('Saving to WordPress…');

    const resp = await fetch(window.SHAPE_API_SAVE_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-WP-Nonce': window.SHAPE_API_NONCE || '',
        'X-Edit-Passcode': window._msmPasscode || ''
      },
      body: JSON.stringify({
        shapes:  fc,
        summary: summary
      })
    });

    if (!resp.ok) {
      showToast('WordPress save failed (HTTP ' + resp.status + ').');
      return;
    }

    await resp.json();
    showToast('Saved to WordPress successfully.');
  } catch (err) {
    console.error(err);
    showToast('WordPress save error (see console).');
  }
}





  function loadFromGeoJSON(fc){
     internalClearShapes({ clearStorage: false });
    if(!fc || fc.type!=='FeatureCollection') return;
    (fc.features||[]).forEach(f=>{
      if(!f.geometry || f.geometry.type!=='Polygon') return;
      const ring = f.geometry.coordinates && f.geometry.coordinates[0];
      if(!Array.isArray(ring)) return;
      const path = ring.map(([lng,lat]) => ({lat, lng}));
        const name     = f.properties?.name          || '';
        const owner    = f.properties?.landOwner     || '';
        const rent     = f.properties?.rentalFee     || '';
        const state    = f.properties?.state         || '';
        const city     = f.properties?.city          || '';
        const status   = f.properties?.status        || 'Prospecting';
        const capacity = f.properties?.capacityMwp   || '';
        const tenure   = f.properties?.tenure        || '';
        const flood    = f.properties?.floodZone     || 'Unknown';
        const terrain  = f.properties?.terrain       || 'Unknown';
        const notes    = f.properties?.notes         || '';
        const connVoltage    = f.properties?.connVoltage    || '';
        const connDistKm     = f.properties?.connDistKm     || '';
        const nearestSub     = f.properties?.nearestSub     || '';
        const nearestSubDist = f.properties?.nearestSubDist || '';
        
        addPolygonFromPath(path, name, owner, rent, state, city,
          status, capacity, tenure, flood, terrain, notes,
          connVoltage, connDistKm, nearestSub, nearestSubDist);
          // isNew is NOT passed → defaults to false → skips per-shape UI updates

    });

    // Single final render after ALL shapes are loaded — nothing pre-selected
    refreshBoundarySummary();
    setActiveShape(null);
  }

async function loadFromWordPress(){
  // Warn if there are unsaved shapes already on screen
  if (drawnShapes.length > 0) {
    const ok = confirm('Loading from WordPress will replace all current shapes on screen.\nAny unsaved changes will be lost. Continue?');
    if (!ok) return;
  }
  if (!window.SHAPE_API_SAVE_URL){
    showToast('ERROR: WordPress load URL not configured.');
    return;
  }

  try {
    const resp = await fetch(window.SHAPE_API_SAVE_URL, {
      method: 'GET',
      headers: {
        'X-WP-Nonce': window.SHAPE_API_NONCE || ''
      }
    });

    if (!resp.ok) {
      showToast('Failed to load shapes from WordPress (HTTP ' + resp.status + ').');
      return;
    }

    const data = await resp.json();

    if (!data || !data.shapes) {
      showToast('No saved shapes found in WordPress.');
      return;
    }

    // data.shapes is the same FeatureCollection we saved
    loadFromGeoJSON(data.shapes);
    showToast('Loaded saved shapes from WordPress.');
  } catch (err) {
    console.error(err);
    showToast('Error loading shapes from WordPress.');
  }
}



  function downloadFile(filename, content, mime='application/octet-stream'){
    const blob = new Blob([content], {type: mime});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; a.click();
    URL.revokeObjectURL(url);
  }

  function exportGeoJSON(){
    const fc = toGeoJSON();
    downloadFile('site-boundaries.geojson', JSON.stringify(fc, null, 2), 'application/geo+json');
  }

  function exportKML(){
    const placemarks = drawnShapes.map(s=>{
      const coords = s.gobj.getPath().getArray()
        .concat([s.gobj.getPath().getAt(0)])
        .map(p => `${p.lng()},${p.lat()},0`).join(' ');
      const nm = (s.name || 'Site');
      const metrics = polygonMetrics(s.gobj);
      let desc = '';
      if (metrics){
        desc =
          `Area: ${metrics.areaAc.toFixed(4)} acres `
          + `(${metrics.areaM2.toFixed(2)} m²)\n`
          + `Centroid: ${metrics.centroid.lat.toFixed(6)}, `
          + `${metrics.centroid.lng.toFixed(6)}`;
      }
      return `
      <Placemark>
        <name>${nm}</name>
        <description><![CDATA[${desc}]]></description>
        <Style>
          <LineStyle><color>ff205e1b</color><width>2</width></LineStyle>
          <PolyStyle><color>406abb66</color></PolyStyle>
        </Style>
        <Polygon>
          <outerBoundaryIs>
            <LinearRing>
              <coordinates>${coords}</coordinates>
            </LinearRing>
          </outerBoundaryIs>
        </Polygon>
      </Placemark>`;
    }).join('\n');

    const kml = `<?xml version="1.0" encoding="UTF-8"?>
  <kml xmlns="http://www.opengis.net/kml/2.2">
    <Document>
      <name>Samaiden Site Boundaries</name>
      ${placemarks}
    </Document>
  </kml>`;
    downloadFile('site-boundaries.kml', kml, 'application/vnd.google-earth.kml+xml');
  }

  // ==========================================================================
  // KMZ EXPORT — Google Earth / QGIS compatible
  // --------------------------------------------------------------------------
  // Produces proper KML (uses <n>, not the broken <n> in the old exportKML),
  // embeds ALL field data in <ExtendedData> for round-trip import, wraps in
  // a zip as .kmz so Google Earth opens it directly.
  // ==========================================================================
  // XML-entity escape. Used everywhere — names, Data values, descriptions.
  // Also strips XML-1.0-illegal control chars (everything below 0x20 except
  // tab, LF, CR), which would otherwise produce a malformed file.
  function msmEscapeXml(s){
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;')
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
  }

  function msmKmlStyleForStatus(status){
    // KML colors are AABBGGRR (alpha, blue, green, red) — note the byte order
    // is the REVERSE of CSS #RRGGBB. Values below are derived directly from
    // MSM_STATUS_STYLES so the KMZ renders in Google Earth with the same
    // palette the app uses on the live map. Fill alpha 0x73 ≈ 0.45 to match
    // fillOpacity: 0.45 in the polygon constructor.
    const map = {
      'Prospecting': { line: 'ff53c800', fill: '73caf6b9' }, // green   (#00c853 / #b9f6ca)
      'Shortlisted': { line: 'ff25a8f9', fill: '73c4f9ff' }, // amber   (#f9a825 / #fff9c4)
      'Submitted':   { line: 'ffc06515', fill: '73fbdebb' }, // blue    (#1565c0 / #bbdefb)
      'Awarded':     { line: 'ff205e1b', fill: '73a7d6a5' }, // d.green (#1b5e20 / #a5d6a7)
      'Rejected':    { line: 'ff2828c6', fill: '73d2cdff' }  // red     (#c62828 / #ffcdd2)
    };
    return map[status] || map['Prospecting'];
  }

  function msmBuildKml(){
    const placemarks = drawnShapes.map(s => {
      if (!s.gobj || !s.gobj.getPath) return '';
      const path = s.gobj.getPath().getArray();
      if (path.length < 3) return '';

      // KML LinearRing needs the first point repeated as the last
      const coordStr = path.concat([path[0]])
        .map(p => `${p.lng()},${p.lat()},0`).join(' ');

      const metrics = polygonMetrics(s.gobj);
      const statusKey = (s.status || 'Prospecting').toLowerCase();

      const fields = {
        shapeId:        s.id,
        name:           s.name,
        owner:          s.owner,
        rental:         s.rental,
        state:          s.state,
        city:           s.city,
        status:         s.status,
        capacity:       s.capacity,
        tenure:         s.tenure,
        flood:          s.flood,
        terrain:        s.terrain,
        notes:          s.notes,
        connVoltage:    s.connVoltage,
        connDistKm:     s.connDistKm,
        nearestSub:     s.nearestSub,
        nearestSubDist: s.nearestSubDist,
        areaAcres:      metrics ? metrics.areaAc.toFixed(4) : '',
        areaM2:         metrics ? metrics.areaM2.toFixed(2) : ''
      };

      const extendedData = Object.entries(fields)
        .map(([k, v]) => `        <Data name="${msmEscapeXml(k)}"><value>${msmEscapeXml(v)}</value></Data>`)
        .join('\n');

      // Build description as plain text, then XML-escape the whole thing.
      // No CDATA — that's what was breaking Google Earth's parser when user
      // input contained sequences like ']]>'. Google Earth still renders
      // <br/> as a line break after entity decoding (description balloons
      // are HTML-rendered).
      const descPieces = [
        s.owner           ? `Owner: ${s.owner}`                                                          : '',
        (s.state||s.city) ? `Location: ${[s.city, s.state].filter(Boolean).join(', ')}`                  : '',
        s.status          ? `Status: ${s.status}`                                                        : '',
        s.capacity        ? `Capacity: ${s.capacity} MWp`                                                : '',
        metrics           ? `Area: ${metrics.areaAc.toFixed(4)} acres (${metrics.areaM2.toFixed(2)} m²)` : '',
        s.connVoltage     ? `Grid: ${s.connVoltage} @ ${s.connDistKm} km`                                : '',
        s.nearestSub      ? `Nearest sub: ${s.nearestSub}`                                               : '',
        s.notes           ? `Notes: ${s.notes}`                                                          : ''
      ].filter(Boolean).join('<br/>');
      const descSafe = msmEscapeXml(descPieces);

      return `    <Placemark>
      <name>${msmEscapeXml(s.name || 'Site')}</name>
      <description>${descSafe}</description>
      <styleUrl>#style_${statusKey}</styleUrl>
      <ExtendedData>
${extendedData}
      </ExtendedData>
      <Polygon>
        <tessellate>1</tessellate>
        <altitudeMode>clampToGround</altitudeMode>
        <outerBoundaryIs>
          <LinearRing>
            <coordinates>${coordStr}</coordinates>
          </LinearRing>
        </outerBoundaryIs>
      </Polygon>
    </Placemark>`;
    }).filter(Boolean).join('\n');

    const styles = ['Prospecting','Shortlisted','Submitted','Awarded','Rejected'].map(st => {
      const c = msmKmlStyleForStatus(st);
      return `    <Style id="style_${st.toLowerCase()}">
      <LineStyle><color>${c.line}</color><width>3</width></LineStyle>
      <PolyStyle><color>${c.fill}</color></PolyStyle>
    </Style>`;
    }).join('\n');

    return `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>Samaiden Site Boundaries</name>
    <description>Exported ${msmEscapeXml(new Date().toISOString())} — ${drawnShapes.length} site(s)</description>
${styles}
${placemarks}
  </Document>
</kml>`;
  }

  async function exportKMZ(){
    if (!drawnShapes.length){
      showToast('No lands to export.');
      return;
    }
    if (typeof JSZip === 'undefined'){
      showToast('KMZ library not loaded — please refresh the page.');
      return;
    }
    try {
      const kml = msmBuildKml();
      const zip = new JSZip();
      // Google Earth expects the main KML inside the zip to be named doc.kml
      zip.file('doc.kml', kml);
      const blob = await zip.generateAsync({
        type: 'blob',
        mimeType: 'application/vnd.google-earth.kmz',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 }
      });
      const filename = `samaiden-lands-${new Date().toISOString().slice(0,10)}.kmz`;
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      // Defer cleanup so the browser has time to grab the blob bytes.
      // Calling revokeObjectURL synchronously after click() can produce
      // empty/corrupt downloads in Firefox and some Safari builds.
      setTimeout(() => {
        if (a.parentNode) a.parentNode.removeChild(a);
        URL.revokeObjectURL(url);
      }, 1000);
      showToast(`Exported ${drawnShapes.length} land(s) to ${filename}`);
    } catch (err) {
      console.error('KMZ export failed', err);
      showToast('KMZ export failed. See console for details.');
    }
  }

  // ==========================================================================
  // KMZ IMPORT — accepts .kmz or .kml
  // --------------------------------------------------------------------------
  // Reads every Placemark with a Polygon, pulls ExtendedData back into shape
  // fields, draws them via addPolygonFromPath(). Gated by the edit lock.
  // ==========================================================================
  async function importKmzFile(file){
    if (!msmIsUnlocked()){
      showToast('Unlock editing first to import KMZ.');
      return;
    }
    if (!file) return;

    const lname = (file.name || '').toLowerCase();
    const isKmz = lname.endsWith('.kmz');
    const isKml = lname.endsWith('.kml');
    if (!isKmz && !isKml){
      showToast('Please select a .kmz or .kml file.');
      return;
    }

    setStatus(`Reading ${file.name}…`);

    try {
      let kmlText = '';
      if (isKmz){
        if (typeof JSZip === 'undefined'){
          showToast('KMZ library not loaded — please refresh the page.');
          setStatus('');
          return;
        }
        const buf = await file.arrayBuffer();
        const zip = await JSZip.loadAsync(buf);
        // Find the first .kml inside the archive (usually doc.kml)
        let kmlEntry = null;
        zip.forEach((path, entry) => {
          if (!kmlEntry && !entry.dir && path.toLowerCase().endsWith('.kml')){
            kmlEntry = entry;
          }
        });
        if (!kmlEntry){
          showToast('No KML file found inside the KMZ.');
          setStatus('');
          return;
        }
        kmlText = await kmlEntry.async('string');
      } else {
        kmlText = await file.text();
      }

      const parser = new DOMParser();
      const doc = parser.parseFromString(kmlText, 'application/xml');
      if (doc.querySelector('parsererror')){
        showToast('KML is malformed — could not parse.');
        setStatus('');
        return;
      }

      const placemarks = doc.getElementsByTagName('Placemark');
      let imported = 0;
      let skipped = 0;
      const beforeCount = drawnShapes.length;

      for (let i = 0; i < placemarks.length; i++){
        const pm = placemarks[i];
        const polys = pm.getElementsByTagName('Polygon');
        if (!polys.length){ skipped++; continue; }

        const ring = polys[0].getElementsByTagName('LinearRing')[0];
        if (!ring){ skipped++; continue; }
        const coordsEl = ring.getElementsByTagName('coordinates')[0];
        if (!coordsEl){ skipped++; continue; }

        const raw = (coordsEl.textContent || '').trim();
        if (!raw){ skipped++; continue; }

        // KML coords: "lng,lat,alt lng,lat,alt ..."  (alt optional)
        const path = raw.split(/\s+/).map(pt => {
          const parts = pt.split(',');
          const lng = parseFloat(parts[0]);
          const lat = parseFloat(parts[1]);
          if (Number.isNaN(lat) || Number.isNaN(lng)) return null;
          return { lat, lng };
        }).filter(Boolean);

        if (path.length < 3){ skipped++; continue; }

        // Drop duplicated closing vertex
        const first = path[0], last = path[path.length - 1];
        if (first && last && first.lat === last.lat && first.lng === last.lng){
          path.pop();
        }
        if (path.length < 3){ skipped++; continue; }

        // Pull ExtendedData field values
        const data = {};
        const ext = pm.getElementsByTagName('ExtendedData')[0];
        if (ext){
          const datas = ext.getElementsByTagName('Data');
          for (let j = 0; j < datas.length; j++){
            const key = datas[j].getAttribute('name');
            const valEl = datas[j].getElementsByTagName('value')[0];
            if (key && valEl){
              data[key] = (valEl.textContent || '').trim();
            }
          }
        }

        const nameEl = pm.getElementsByTagName('name')[0];
        const pmName = nameEl ? (nameEl.textContent || '').trim() : '';

        addPolygonFromPath(
          path,
          data.name || pmName || 'Imported site',
          data.owner || '',
          data.rental || '',
          data.state || '',
          data.city || '',
          data.status || 'Prospecting',
          data.capacity || '',
          data.tenure || '',
          data.flood || 'Unknown',
          data.terrain || 'Unknown',
          data.notes || '',
          data.connVoltage || '',
          data.connDistKm || '',
          data.nearestSub || '',
          data.nearestSubDist || '',
          false  // isNew=false → skip per-shape side effects during batch
        );
        imported++;
      }

      // Single refresh + save after batch
      refreshBoundarySummary();
      if (typeof saveToWordPress === 'function'){
        saveToWordPress();
      }

      // Zoom to the new shapes
      if (imported > 0 && map && drawnShapes.length > beforeCount){
        try {
          const bounds = new google.maps.LatLngBounds();
          drawnShapes.slice(beforeCount).forEach(s => {
            if (s.gobj && s.gobj.getPath){
              s.gobj.getPath().forEach(p => bounds.extend(p));
            }
          });
          if (!bounds.isEmpty()) map.fitBounds(bounds);
        } catch (e){ /* non-fatal */ }
      }

      setStatus('');
      const skippedMsg = skipped ? ` (${skipped} non-polygon placemark(s) skipped)` : '';
      showToast(`Imported ${imported} land(s) from ${file.name}${skippedMsg}`);
    } catch (err) {
      console.error('KMZ import failed', err);
      setStatus('');
      showToast('KMZ import failed. See console for details.');
    }
  }
  function exportSummaryExcel(){
    if (!drawnShapes.length){
      showToast('No lands to export.');
      return;
    }

    // Header row matches the table
    const rows = [[
      '#',
      'Status',
      'Name',
      'Land owner',
      'State (MY)',
      'City (MY)',
      'Rental fee (RM)',
      'Capacity (MWp)',
      'Tenure',
      'Flood zone',
      'Terrain',
      'Nearest substation',
      'Sub dist (km)',
      'Connection voltage',
      'Line dist (km)',
      'Notes',
      'Area (acres)',
      'Area (m²)',
      'Nearest Sub',
      'Sub dist (km)',
      'Grid voltage',
      'Line dist (km)'
    ]];


    drawnShapes.forEach((s, index) => {
      const metrics = polygonMetrics(s.gobj);
      if (!metrics){
        rows.push([
          index + 1,
          s.status   || 'Prospecting',
          s.name     || '',
          s.owner    || '',
          s.state    || '',
          s.city     || '',
          s.rental   || '',
          s.capacity || '',
          s.tenure   || '',
          s.flood    || '',
          s.terrain  || '',
          s.nearestSub     || '',
          s.nearestSubDist || '',
          s.connVoltage    || '',
          s.connDistKm     || '',
          s.notes    || '',
          s.nearestSub||'', s.nearestSubDist||'', s.connVoltage||'', s.connDistKm||''
        ]);

      } else {
        rows.push([
          index + 1,
          s.status   || 'Prospecting',
          s.name     || '',
          s.owner    || '',
          s.state    || '',
          s.city     || '',
          s.rental   || '',
          s.capacity || '',
          s.tenure   || '',
          s.flood    || '',
          s.terrain  || '',
          s.nearestSub     || '',
          s.nearestSubDist || '',
          s.connVoltage    || '',
          s.connDistKm     || '',
          s.notes    || '',
          metrics.areaAc.toFixed(4),
          metrics.areaM2.toFixed(0),
          s.nearestSub||'', s.nearestSubDist||'', s.connVoltage||'', s.connDistKm||''
        ]);
      }
    });

    // Use SheetJS (already loaded at top)
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, 'Site Boundaries');
    XLSX.writeFile(wb, 'site-boundaries-summary.xlsx');
  }

document.addEventListener('DOMContentLoaded', () => {
  const excelBtn = document.getElementById('exportExcelBtn');
  if (excelBtn) {
    excelBtn.addEventListener('click', exportSummaryExcel);
  }

  // KMZ export button
  const kmzBtn = document.getElementById('exportKmzBtn');
  if (kmzBtn) {
    kmzBtn.addEventListener('click', exportKMZ);
  }

  // KMZ import button + hidden file picker
  const kmzImportBtn = document.getElementById('importKmzBtn');
  const kmzFileInput = document.getElementById('kmzFileInput');
  if (kmzImportBtn && kmzFileInput) {
    kmzImportBtn.addEventListener('click', () => {
      if (!msmIsUnlocked()) {
        showToast('Unlock editing first to import KMZ.');
        return;
      }
      kmzFileInput.value = '';   // allow re-selecting the same file
      kmzFileInput.click();
    });
    kmzFileInput.addEventListener('change', (e) => {
      const f = e.target.files && e.target.files[0];
      if (f) importKmzFile(f);
    });
  }
});


  function importFromFile(file){
    const reader = new FileReader();
    reader.onload = () => {
      const text = reader.result;
      try{
        const fc = JSON.parse(text);
        if (fc && fc.type === 'FeatureCollection') {
          loadFromGeoJSON(fc);
          showToast('Imported GeoJSON.');
          return;
        }
      }catch{}
      if (text.includes('<kml')){
        const rings = [...text.matchAll(/<coordinates>([\s\S]*?)<\/coordinates>/g)].map(m=>m[1]);
        if (rings.length){
          internalClearShapes({ clearStorage: false });
          rings.forEach(coordsStr=>{
            const path = coordsStr.trim().split(/\s+/).map(s=>{
              const [lng,lat] = s.split(',').map(Number);
              return {lat, lng};
            });
            addPolygonFromPath(path);
          });
          showToast('Imported KML polygons.');
          return;
        }
      }
      showToast('Unsupported file. Provide GeoJSON (.geojson/.json) or KML (.kml).');
    };
    reader.readAsText(file);
  }

  (function addBoundaryPanelInSidebar(){
    const panel = document.createElement('div');
    panel.className = 'filter-section boundary';
    panel.id = 'siteBoundaryPanel';
    panel.innerHTML = `
      <div class="msm-panel-header">
        <h3><i class="fas fa-draw-polygon"></i> Site Boundary</h3>
      </div>
    
      <div class="boundary-fields">
        <div class="boundary-field">
          <label for="shapeNameInput">Land name</label>
          <input id="shapeNameInput"
                 type="text"
                 placeholder=" ">
        </div>
    
        <div class="boundary-field">
          <label for="shapeOwnerInput">Land owner</label>
          <input id="shapeOwnerInput"
                 type="text"
                 placeholder=" ">
        </div>
    
        <div class="boundary-field">
          <label for="shapeRentInput">Rental fee (RM)</label>
          <input id="shapeRentInput"
                 type="number"
                 min="0"
                 step="0.01"
                 placeholder=" ">
        </div>

        <div class="boundary-field">
          <label for="shapeStatusInput">Status</label>
          <select id="shapeStatusInput">
            <option value="Prospecting">Prospecting</option>
            <option value="Shortlisted">Shortlisted</option>
            <option value="Submitted">Submitted</option>
            <option value="Awarded">Awarded</option>
            <option value="Rejected">Rejected</option>
          </select>
        </div>
        <div class="boundary-field">
          <label for="shapeCapacityInput">Proposed capacity (MWp)</label>
          <input id="shapeCapacityInput" type="number" min="0" step="0.1" placeholder=" ">
        </div>
        <div class="boundary-field">
          <label for="shapeTenureInput">Land tenure</label>
          <select id="shapeTenureInput">
            <option value="">Unknown</option>
            <option value="Freehold">Freehold</option>
            <option value="Leasehold">Leasehold</option>
            <option value="Government">Government</option>
          </select>
        </div>
        <div class="boundary-field">
          <label for="shapeFloodInput">Flood zone</label>
          <select id="shapeFloodInput">
            <option value="Unknown">Unknown</option>
            <option value="No">No</option>
            <option value="Yes">Yes</option>
          </select>
        </div>
        <div class="boundary-field">
          <label for="shapeTerrainInput">Terrain <button type="button" id="detectTerrainBtn" style="margin-left:6px;font-size:10px;padding:2px 7px;border-radius:999px;border:1px solid var(--line);background:#f0fdf4;color:var(--brand-deep);cursor:pointer;font-family:inherit;transition:all .15s;" title="Auto-detect from elevation data"><i class="fas fa-mountain"></i> Auto</button></label>
          <select id="shapeTerrainInput">
            <option value="Unknown">Unknown</option>
            <option value="Flat">Flat</option>
            <option value="Gentle Slope">Gentle Slope</option>
            <option value="Steep">Steep</option>
          </select>
        </div>
        <div class="boundary-field">
          <label for="shapeNotesInput">Notes / remarks</label>
          <textarea id="shapeNotesInput" rows="2" placeholder=" "></textarea>
        </div>
      </div>

      <div id="shapeStatus" class="msm-boundary-status">
        No shape selected.
      </div>

      <!-- Solar scorecard (shown when shape selected) -->
      <div id="solarScorecard" class="solar-scorecard" style="display:none;"></div>

      <div class="msm-boundary-controls" style="grid-template-columns:repeat(3,minmax(0,1fr));">
        <button class="filter-btn" id="drawPolygonBtn">
          <i class="fas fa-draw-polygon"></i> Polygon
        </button>
        <button class="filter-btn" id="drawRectangleBtn">
          <i class="fas fa-vector-square"></i> Rectangle
        </button>
        <button class="filter-btn" id="stopDrawingBtn" title="Cancel active drawing">
          <i class="fas fa-times-circle"></i> Cancel Draw
        </button>
        <button class="filter-btn" id="saveBoundaryBtn">
          <i class="fas fa-save"></i> Save
        </button>
        <button class="filter-btn secondary" id="loadBoundaryBtn" title="Reload all saved lands from WordPress server">
          <i class="fas fa-cloud-download-alt"></i> Reload
        </button>
        <button class="filter-btn" id="clearBoundaryBtn">
          <i class="fas fa-eraser"></i> Clear
        </button>
      </div>
      <div style="margin-top:6px;">
        <button class="filter-btn" id="findConnectionBtn" style="width:100%;background:linear-gradient(135deg,#0d47a1,#1565c0);">
          <i class="fas fa-bolt"></i> Find Nearest Grid Connection
        </button>
      </div>
    `;


    const toolsTab = document.getElementById('tab-tools');
    if (toolsTab) {
      toolsTab.appendChild(panel);
    } else {
      const sidebar = document.querySelector('.sidebar');
      const routeCard = [...sidebar.querySelectorAll('.filter-section h3')]
        .find(h => /Measure via Road/i.test(h.textContent))?.parentElement;
      if (routeCard) {
        routeCard.after(panel);
      } else {
        const firstFilter = sidebar.querySelector('.filter-section');
        (firstFilter ? firstFilter.before(panel) : sidebar.appendChild(panel));
      }
    }

    const polyBtn  = panel.querySelector('#drawPolygonBtn');
    const rectBtn  = panel.querySelector('#drawRectangleBtn');
    const stopBtn  = panel.querySelector('#stopDrawingBtn');
    const saveBtn  = panel.querySelector('#saveBoundaryBtn');
    const loadBtn  = panel.querySelector('#loadBoundaryBtn');
    const clearBtn = panel.querySelector('#clearBoundaryBtn');
    
    if (polyBtn)  polyBtn.addEventListener('click', () => startDraw('polygon'));
    if (rectBtn)  rectBtn.addEventListener('click', () => startDraw('rectangle'));
    if (stopBtn)  stopBtn.addEventListener('click', () => {
      if (drawingManager) drawingManager.setDrawingMode(null);
    });
    
    if (saveBtn)  saveBtn.addEventListener('click', saveToWordPress);
    if (loadBtn)  loadBtn.addEventListener('click', loadFromWordPress);
    if (clearBtn) clearBtn.addEventListener('click', clearShapes);

    const findConnBtn    = panel.querySelector('#findConnectionBtn');
    if (findConnBtn) findConnBtn.addEventListener('click', msmFindNearestConnection);

    const detectTerrBtn = panel.querySelector('#detectTerrainBtn');
    if (detectTerrBtn) detectTerrBtn.addEventListener('click', () => {
      if (!activeShape) { showToast('Select a land parcel first.'); return; }
      showToast('Detecting terrain from elevation data…');
      autoDetectTerrain(activeShape);
    });

    // Solar field change handlers (event delegation on panel)
    panel.addEventListener('change', e => {
      if (!activeShape) return;
      const t = e.target;
      if (t.id === 'shapeNameInput'){
        activeShape.name = t.value.trim();
        refreshBoundarySummary();
        updateShapeMetrics();
      } else if (t.id === 'shapeOwnerInput'){
        activeShape.owner = t.value.trim();
        refreshBoundarySummary();
      } else if (t.id === 'shapeRentInput'){
        activeShape.rental = t.value.trim();
        refreshBoundarySummary();
        msmUpdateSolarPanel(activeShape); // refresh rental/acre in scorecard
      } else if (t.id === 'shapeStatusInput'){
        activeShape.status = t.value;
        msmApplyStatusStyle(activeShape);
        refreshBoundarySummary();
      } else if (t.id === 'shapeCapacityInput'){
        activeShape.capacity = t.value;
        refreshBoundarySummary();
      } else if (t.id === 'shapeTenureInput'){
        activeShape.tenure = t.value;
      } else if (t.id === 'shapeFloodInput'){
        activeShape.flood = t.value;
      } else if (t.id === 'shapeTerrainInput'){
        activeShape.terrain = t.value;
      }
    });
    panel.addEventListener('input', e => {
      if (!activeShape) return;
      const t = e.target;
      if (t.id === 'shapeNotesInput')  activeShape.notes  = t.value;
      if (t.id === 'shapeNameInput')   activeShape.name   = t.value.trim();
      if (t.id === 'shapeOwnerInput')  activeShape.owner  = t.value.trim();
      if (t.id === 'shapeRentInput')   activeShape.rental = t.value.trim();
    });

  })();

  // ---------- Map ----------
  function initMap(){
    infoWindow = new google.maps.InfoWindow();
    shapeGeocoder = new google.maps.Geocoder();
    window.msmElevationService = new google.maps.ElevationService();
    window.closePMUInfo = () => infoWindow.close();
    map = new google.maps.Map(document.getElementById('map'),{
      zoom:6,
      center:{lat:4.2105,lng:101.9758},
      mapTypeId: google.maps.MapTypeId.HYBRID,
      mapTypeControl: true,
      mapTypeControlOptions: {
        style: google.maps.MapTypeControlStyle.HORIZONTAL_BAR,
        position: google.maps.ControlPosition.TOP_RIGHT,
        mapTypeIds: [
          google.maps.MapTypeId.ROADMAP,
          google.maps.MapTypeId.SATELLITE,
          google.maps.MapTypeId.HYBRID,
          google.maps.MapTypeId.TERRAIN
        ]
      },
      zoomControl: true,
      streetViewControl: true,
      fullscreenControl: true,
      minZoom: 5,
      maxZoom: 21
    });

    map.setOptions({
      zoomControlOptions:       { position: google.maps.ControlPosition.RIGHT_TOP },
      streetViewControlOptions: { position: google.maps.ControlPosition.RIGHT_TOP },
      fullscreenControlOptions: { position: google.maps.ControlPosition.RIGHT_TOP },
      clickableIcons: true,
      gestureHandling: "greedy",
      mapId: "DEMO_MAP_ID"
    });

    document.getElementById('selectAllBtn').addEventListener('click', selectAllStates);
    document.getElementById('clearAllBtn').addEventListener('click', clearAllStates);
    document.getElementById('searchInput').addEventListener('input', filterSubstations);
    document.getElementById('clearSavedSitesBtn')?.addEventListener('click', clearAllSavedSites);
    document.getElementById('loadFromUrl').addEventListener('click', ()=>{
      const url = document.getElementById('mediaUrl').value.trim();
      handleUrl(url);
    });

    attachPlacesSearch();
    attachLocateMe();
    loadSavedSites();

    // ----- Coordinate bar on map -----
    const coordText = document.getElementById('coordText');
    const zoomText = document.getElementById('zoomText');
    if (coordText) {
      map.addListener('mousemove', (e) => {
        coordText.textContent = `${e.latLng.lat().toFixed(6)}, ${e.latLng.lng().toFixed(6)}`;
      });
    }
    if (zoomText) {
      map.addListener('zoom_changed', () => {
        zoomText.textContent = `Zoom: ${map.getZoom()}`;
      });
    }

    // ----- Measure distance via route -----
    let routeStart = null;
    let routeEnd = null;
    let routePickMode = null;
    const routeSummary = document.getElementById('routeSummary');
    const directionsService = new google.maps.DirectionsService();
    const directionsRenderer = new google.maps.DirectionsRenderer({
      map,
      polylineOptions: { strokeWeight: 5, strokeColor: '#0DFF7A' },
      suppressMarkers: true
    });

    function maybeRenderRoute(){
      if (!routeStart || !routeEnd) return;
      const request = {
        origin: routeStart,
        destination: routeEnd,
        travelMode: google.maps.TravelMode.DRIVING
      };
      directionsService.route(request, (result, status) => {
        if (status === 'OK') {
          directionsRenderer.setDirections(result);
          const leg = result.routes[0].legs[0];
          const km = (leg.distance.value / 1000).toFixed(2);
          const mins = Math.round(leg.duration.value / 60);
          routeSummary.textContent = `Distance: ${km} km | Time: ${mins} min`;
          showToast(`Route: ${km} km, ${mins} min`);
        } else {
          showToast('Could not calculate route.');
        }
      });
    }

    let startMarker = null;
    let endMarker = null;

    document.getElementById('routeStartBtn').addEventListener('click', () => {
      routePickMode = 'start';
      showToast('Click the map to pick START point');
    });
    document.getElementById('routeEndBtn').addEventListener('click', () => {
      routePickMode = 'end';
      showToast('Click the map to pick DESTINATION');
    });

map.addListener('click', (e) => {
  // 0) Custom PMU pin-drop mode
  if (_pinDropActive) {
    const latLng = e.latLng;
    _customPinLatLng = latLng;

    // Remove old temp marker
    if (_customPinMarker) _customPinMarker.setMap(null);

    // Place yellow PMU-style marker
    _customPinMarker = new google.maps.Marker({
      position: latLng,
      map,
      title: 'Custom PMU',
      draggable: true,
      icon: {
        path: google.maps.SymbolPath.CIRCLE,
        scale: 9,
        fillColor: '#F9A825',
        fillOpacity: 1,
        strokeColor: '#fff',
        strokeWeight: 2.5
      }
    });

    _customPinMarker.addListener('dragend', ev => { _customPinLatLng = ev.latLng; });

    // Ask for name then confirm
    const pmuName = (document.getElementById('customSubSearchInput')?.value || '').trim()
                    || 'Custom PMU';

    const shape = _pinDropShapeRef || activeShape;
    if (shape) {
      const metrics = polygonMetrics(shape.gobj);
      if (metrics) {
        const distKm = msmHaversineKm(
          metrics.centroid.lat, metrics.centroid.lng,
          latLng.lat(), latLng.lng()
        );
        shape.nearestSub     = pmuName;
        shape.nearestSubDist = distKm.toFixed(2);
        msmUpdateSolarPanel(shape);
        refreshBoundarySummary();
        showToast('Custom PMU set: ' + pmuName + ' (' + distKm.toFixed(2) + ' km) — click Save to persist.');
      }
    }

    msmEndPinDrop();
    return;
  }

  // 1) Normal mode: no route picking active → treat click as "select point"
  if (!routePickMode) {
    const latLng = e.latLng;
    setSearchLandmark(latLng, 'Selected point');
    if (typeof msmClearConnLines === 'function') msmClearConnLines();
    return;
  }

  // 2) Route picking mode (existing behaviour)
  if (routePickMode === 'start') {
    routeStart = e.latLng;
    if (startMarker) startMarker.setMap(null);
    startMarker = new google.maps.Marker({
      position: e.latLng,
      map,
      title: 'Start point',
      draggable: true,
      icon: {
        path: google.maps.SymbolPath.CIRCLE,
        scale: 7,
        fillColor: '#2E7D32',
        fillOpacity: 1,
        strokeColor: '#fff',
        strokeWeight: 2
      }
    });
    startMarker.addListener('dragend', (ev) => {
      routeStart = ev.latLng;
      maybeRenderRoute();
    });
    showToast('Start point selected — now pick destination');
    routePickMode = 'end';
    maybeRenderRoute();
  } else if (routePickMode === 'end') {
    routeEnd = e.latLng;
    if (endMarker) endMarker.setMap(null);
    endMarker = new google.maps.Marker({
      position: e.latLng,
      map,
      title: 'Destination',
      draggable: true,
      icon: {
        path: google.maps.SymbolPath.CIRCLE,
        scale: 7,
        fillColor: '#1565C0',
        fillOpacity: 1,
        strokeColor: '#fff',
        strokeWeight: 2
      }
    });
    endMarker.addListener('dragend', (ev) => {
      routeEnd = ev.latLng;
      maybeRenderRoute();
    });
    showToast('Destination selected');
    maybeRenderRoute();
    routePickMode = null;
  }
});


    document.getElementById('clearRouteBtn').addEventListener('click', () => {
      directionsRenderer.set('directions', null);
      if (startMarker) startMarker.setMap(null);
      if (endMarker) endMarker.setMap(null);
      startMarker = endMarker = null;
      routeStart = routeEnd = null;
      routeSummary.textContent = 'Distance: – | Time: –';
      clearSearchLandmark();
    });

    (function attachPlacesSearchFS(){
      const inputId = 'mapSearchInputFS';
      const input = document.getElementById(inputId);
      if(!input) return;
      const geocoder = new google.maps.Geocoder();
      function flyTo(latLng, zoom=18, title='Searched location'){
        map.setCenter(latLng);
        map.setZoom(zoom);
        setSearchLandmark(latLng, title);
      }
      function geocodeAndGo(query){
        if(!query) return;
    
        // 1) Try lat,lng first for fullscreen search
        const coord = parseCoordinate(query);
        if (
          coord &&
          coord.lat >= -90 && coord.lat <= 90 &&
          coord.lng >= -180 && coord.lng <= 180
        ) {
          const latLng = new google.maps.LatLng(coord.lat, coord.lng);
          flyTo(
            latLng,
            18,
            `Lat ${coord.lat.toFixed(6)}, Lng ${coord.lng.toFixed(6)}`
          );
          return;
        }
    
        // 2) Otherwise, use address geocoding (your existing logic)
        geocoder.geocode({ address: query }, (results, status) => {
          if(status === 'OK' && results && results[0]){
            const loc = results[0].geometry.location;
            flyTo(
              loc,
              results[0].geometry.viewport ? map.getZoom() : 18,
              results[0].formatted_address
            );
            if(results[0].geometry.viewport){
              map.fitBounds(results[0].geometry.viewport);
              google.maps.event.addListenerOnce(map, 'idle', () => {
                if(map.getZoom() < 16) map.setZoom(16);
              });
            }
          }else{
            showToast('Could not find that address.');
          }
        });
      }

      // Coordinate shortcut — capture phase fires BEFORE Places Autocomplete
  input.addEventListener('keydown', function coordCapture(e){
    if (e.key !== 'Enter') return;
    const coord = parseCoordinate(input.value.trim());
    if (coord && coord.lat >= -90 && coord.lat <= 90 && coord.lng >= -180 && coord.lng <= 180){
      e.preventDefault();
      e.stopImmediatePropagation();
      flyTo(new google.maps.LatLng(coord.lat, coord.lng), 18,
        'Lat ' + coord.lat.toFixed(6) + ', Lng ' + coord.lng.toFixed(6));
    }
  }, true); // capture = true → fires before Places Autocomplete

  const ac = new google.maps.places.Autocomplete(input, {
        fields: ['geometry', 'name', 'formatted_address'],
      });
      ac.bindTo('bounds', map);
      ac.addListener('place_changed', ()=>{
        const place = ac.getPlace();
        if(place && place.geometry){
          if(place.geometry.viewport){
            map.fitBounds(place.geometry.viewport);
            google.maps.event.addListenerOnce(map, 'idle', () => {
              if(map.getZoom() < 16) map.setZoom(16);
            });
          }else{
            flyTo(place.geometry.location, 18, place.formatted_address || place.name);
          }
        }else{
          geocodeAndGo(input.value.trim());
        }
      });
      input.addEventListener('keydown', (e)=>{ if(e.key==='Enter'){ e.preventDefault(); geocodeAndGo(input.value.trim()); } });
        document.getElementById('locateMeBtnFS')?.addEventListener('click', () => {
          const input = document.getElementById('mapSearchInputFS');
          if (!input) return;
        
          const val = input.value.trim();
          if (!val) {
            showToast('Please type a place, address, or "lat,lng" first.');
            input.focus();
            return;
          }
        
          // Same trick: pretend user pressed Enter in the FS search box
          const evt = new KeyboardEvent('keydown', { key: 'Enter' });
          input.dispatchEvent(evt);
        });

    })();

    document.getElementById('routeStartBtnFS')?.addEventListener('click', ()=> document.getElementById('routeStartBtn').click());
    document.getElementById('routeEndBtnFS')  ?.addEventListener('click', ()=> document.getElementById('routeEndBtn').click());
    document.getElementById('clearRouteBtnFS')?.addEventListener('click', ()=> document.getElementById('clearRouteBtn').click());
    clearSearchLandmark();

    const routeSummaryEl = document.getElementById('routeSummary');
    const routeSummaryFS = document.getElementById('routeSummaryFS');
    const mo = new MutationObserver(()=> routeSummaryFS.textContent = routeSummaryEl.textContent);
    mo.observe(routeSummaryEl, { childList:true, characterData:true, subtree:true });

    document.getElementById('clearLinesBtn').addEventListener('click', clearAllLines);
    document.getElementById('selectAllLinesBtn').addEventListener('click', selectAllLines);
    ['toggle11','toggle33','toggle132','toggle275','toggle500','toggleKTM','toggleHwy']
      .forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('change', onLayerToggleChange);
      });

    document.getElementById('importExcelBtn').addEventListener('click', ()=>{
      document.getElementById('fileInput').click();
    });
    document.getElementById('fileInput').addEventListener('change', (e)=>{
      const f = e.target.files?.[0];
      if (f) handleFile(f);
    });

    const dz = document.getElementById('dropZone');
    ;['dragenter','dragover'].forEach(evt=>dz.addEventListener(evt, e=>{e.preventDefault();e.stopPropagation();dz.classList.add('drag');}));
    ;['dragleave','drop'].forEach(evt=>dz.addEventListener(evt, e=>{e.preventDefault();e.stopPropagation();dz.classList.remove('drag');}));
    dz.addEventListener('drop', e=>{
      const f = e.dataTransfer.files?.[0];
      if(f) handleFile(f);
    });

    (async () => {
      try {
        await handleUrl(`${MEDIA_FILE_URL}?t=${Date.now()}`);
      } catch {
        setTimeout(() => handleUrl(`${MEDIA_FILE_URL}?t=${Date.now()}`), 1200);
      }
    })();

    // Auto-load transmission lines on startup (checkboxes are pre-checked,
    // so user doesn't need to click "Select All" manually)
    loadTransmissionLines().catch(err => {
      console.warn('Transmission lines auto-load failed, will retry:', err);
      setTimeout(() => loadTransmissionLines().catch(()=>{}), 2000);
    });

    // Initialize DrawingManager AFTER map exists
    drawingManager = new google.maps.drawing.DrawingManager({
      drawingMode: null,
      drawingControl: false,
      polygonOptions: {
        clickable: true,
        editable: true,
        strokeColor: '#00c853',
        strokeOpacity: 1,
        strokeWeight: 3,
        fillColor: '#b9f6ca',
        fillOpacity: 0.45
      },
      rectangleOptions: {
        clickable: true,
        editable: true,
        strokeColor: '#00c853',
        strokeOpacity: 1,
        strokeWeight: 3,
        fillColor: '#b9f6ca',
        fillOpacity: 0.35
      }
    });
    drawingManager.setMap(map);

google.maps.event.addListener(drawingManager, 'overlaycomplete', (e) => {
  // Stop drawing mode after one shape
  drawingManager.setDrawingMode(null);

  const nameEl  = document.getElementById('shapeNameInput');
  const ownerEl = document.getElementById('shapeOwnerInput');
  const rentEl  = document.getElementById('shapeRentInput');

  const name   = (nameEl  && nameEl.value  || '').trim();
  const owner  = (ownerEl && ownerEl.value || '').trim();
  const rental = (rentEl  && rentEl.value  || '').trim();

  let path;

  if (e.type === google.maps.drawing.OverlayType.POLYGON) {
    const poly = e.overlay;
    path = poly.getPath().getArray().map(p => ({ lat: p.lat(), lng: p.lng() }));
    poly.setMap(null);

  } else if (e.type === google.maps.drawing.OverlayType.RECTANGLE) {
    const poly = rectToPolygon(e.overlay);
    path = poly.getPath().getArray().map(p => ({ lat: p.lat(), lng: p.lng() }));
    poly.setMap(null);
  }

  if (!path || !path.length) return;

  // Create shape with name + owner + rental
  addPolygonFromPath(path, name, owner, rental, '', '', 'Prospecting', '', '', 'Unknown', 'Unknown', '', '', '', '', '', true);

  // (optional) clear the small form fields
  if (nameEl)  nameEl.value = '';
  if (ownerEl) ownerEl.value = '';
  if (rentEl)  rentEl.value = '';

  // Auto-save every time a new shape is created
  saveToWordPress();
});



    // Panels stay inside sidebar tabs — no floating map overlay
  const tools = document.getElementById('mapToolsControl');
  if (tools) tools.style.display = 'block';


    // Collapsible cards restore state
    (function(){
      function toggleCard(card, forceCollapse){
        const btn = card.querySelector('.min-btn');
        const collapse = (forceCollapse !== undefined) ? forceCollapse
                        : !card.classList.contains('collapsed');
        card.classList.toggle('collapsed', collapse);
        if (btn){
          btn.setAttribute('aria-expanded', (!collapse).toString());
          btn.textContent = collapse ? '+' : '–';
        }
        if (card.id) localStorage.setItem(card.id + '_collapsed', String(collapse));
      }
      document.addEventListener('click', (e)=>{
        const btn = e.target.closest('.collapsible .min-btn');
        if (!btn) return;
        const card = btn.closest('.collapsible');
        if (card){ 
          e.preventDefault();
          e.stopPropagation();
          toggleCard(card);
        }
      });
      document.querySelectorAll('.collapsible').forEach(card=>{
        const saved = card.id ? localStorage.getItem(card.id + '_collapsed') : null;
        if (saved === 'true') toggleCard(card, true);
      });
      if (window.matchMedia('(max-width: 768px)').matches){
        document.querySelectorAll('.collapsible').forEach(card=>{
          if (card.id && localStorage.getItem(card.id + '_collapsed') === null){
            toggleCard(card, true);
          }
        });
      }
    })();

    // Auto-load shapes AFTER map + drawingManager are ready
    loadFromWordPress();   // fetch saved shapes from WordPress on page load
    
    applyShareLinkFromUrl();
    
  }


  // ===== SOLAR BID HELPER FUNCTIONS =====

  // Status colour map for polygons
  const MSM_STATUS_STYLES = {
    'Prospecting': { fill: '#b9f6ca', stroke: '#00c853' },
    'Shortlisted': { fill: '#fff9c4', stroke: '#f9a825' },
    'Submitted':   { fill: '#bbdefb', stroke: '#1565c0' },
    'Awarded':     { fill: '#a5d6a7', stroke: '#1b5e20' },
    'Rejected':    { fill: '#ffcdd2', stroke: '#c62828' },
  };

  function msmApplyStatusStyle(shape) {
    if (!shape || !shape.gobj) return;
    const st = shape.status || 'Prospecting';
    const style = MSM_STATUS_STYLES[st] || MSM_STATUS_STYLES['Prospecting'];
    shape.gobj.setOptions({
      fillColor: style.fill,
      strokeColor: style.stroke,
    });
  }

  function msmUpdateSolarPanel(shape) {
    const scorecard   = document.getElementById('solarScorecard');
    const nameInp     = document.getElementById('shapeNameInput');
    const ownerInp    = document.getElementById('shapeOwnerInput');
    const rentInp     = document.getElementById('shapeRentInput');
    const statusSel   = document.getElementById('shapeStatusInput');
    const capacityInp = document.getElementById('shapeCapacityInput');
    const tenureSel   = document.getElementById('shapeTenureInput');
    const floodSel    = document.getElementById('shapeFloodInput');
    const terrainSel  = document.getElementById('shapeTerrainInput');
    const notesInp    = document.getElementById('shapeNotesInput');

    if (!shape) {
      if (scorecard) scorecard.style.display = 'none';
      [nameInp, ownerInp, rentInp, capacityInp, notesInp].forEach(el => {
        if (el) el.value = '';
      });
      [statusSel, tenureSel, floodSel, terrainSel].forEach(el => {
        if (el) el.selectedIndex = 0;
      });
      return;
    }

    // Basic fields
    if (nameInp)  nameInp.value  = shape.name   || '';
    if (ownerInp) ownerInp.value = shape.owner  || '';
    if (rentInp)  rentInp.value  = shape.rental || '';

    // Solar fields
    if (statusSel)   statusSel.value   = shape.status   || 'Prospecting';
    if (capacityInp) capacityInp.value = shape.capacity || '';
    if (tenureSel)   tenureSel.value   = shape.tenure   || '';
    if (floodSel)    floodSel.value    = shape.flood    || 'Unknown';
    if (terrainSel)  terrainSel.value  = shape.terrain  || 'Unknown';
    if (notesInp)    notesInp.value    = shape.notes    || '';

    if (scorecard) {
      const metrics = polygonMetrics(shape.gobj);
      const area  = metrics ? metrics.areaAc.toFixed(2) + ' ac' : '—';
      const rentalPerAc = shape.rental
        ? 'RM ' + parseFloat(shape.rental).toLocaleString('en-MY', {minimumFractionDigits:2, maximumFractionDigits:2})
        : '—';

      scorecard.style.display = 'block';
      scorecard.innerHTML = `
        <h4><i class="fas fa-solar-panel"></i> Site scorecard</h4>
        <table>
          <tr><td>Status</td><td><span class="status-badge status-${shape.status||'Prospecting'}">${shape.status||'Prospecting'}</span></td></tr>
          <tr><td>Capacity</td><td>${shape.capacity ? shape.capacity+' MWp' : '—'}</td></tr>
          <tr><td>Area</td><td>${area}</td></tr>
          <tr><td>Total rental</td><td>${rentalPerAc}</td></tr>
          <tr><td>Tenure</td><td>${shape.tenure || '—'}</td></tr>
          <tr><td>Flood zone</td><td>${shape.flood || '—'}</td></tr>
          <tr><td>Terrain</td><td>${shape.terrain || '—'}</td></tr>
        </table>
        <hr class="sc-divider">
        <div class="sc-section"><i class="fas fa-plug"></i> Grid connection</div>
        <table>
          <tr><td>Nearest sub</td><td>${shape.nearestSub ? shape.nearestSub + ' (' + shape.nearestSubDist + ' km)' : '—'}</td></tr>
          <tr><td>Nearest line</td><td>${shape.connVoltage ? shape.connVoltage + ' (' + shape.connDistKm + ' km)' : '—'}</td></tr>
        </table>
      `;
    }
  }

  // Haversine distance in km
  function msmHaversineKm(lat1, lng1, lat2, lng2) {
    const R = 6371;
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLng = (lng2 - lng1) * Math.PI / 180;
    const a = Math.sin(dLat/2)**2
      + Math.cos(lat1*Math.PI/180) * Math.cos(lat2*Math.PI/180) * Math.sin(dLng/2)**2;
    return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  }

  // Point-to-segment distance in km (flat-earth approx, fine for <100 km)
  function msmPtSegDistKm(plat, plng, alat, alng, blat, blng) {
    const deg2km = 111.32;
    const cosLat = Math.cos(plat * Math.PI / 180);
    const px = (plng - alng) * deg2km * cosLat;
    const py = (plat - alat) * deg2km;
    const dx = (blng - alng) * deg2km * cosLat;
    const dy = (blat - alat) * deg2km;
    const len2 = dx*dx + dy*dy;
    const t = len2 > 0 ? Math.max(0, Math.min(1, (px*dx + py*dy) / len2)) : 0;
    return Math.sqrt((px - t*dx)**2 + (py - t*dy)**2);
  }

  // Stores the two candidates found, for modal confirmation
  let _gridCandidates = null;
  // Custom PMU pin state
  let _customPinMarker  = null;   // the temporary yellow marker on map
  let _customPinLatLng  = null;   // confirmed LatLng
  let _pinDropActive    = false;  // map is in pin-drop mode
  let _pinDropShapeRef  = null;   // which shape we're pinning for
  let _customSubPicked  = null;   // { name, distKm } chosen from search list

  async function msmFindNearestConnection() {
    if (!activeShape) { showToast('Select a land parcel first.'); return; }
    const metrics = polygonMetrics(activeShape.gobj);
    if (!metrics)    { showToast('Shape has no valid geometry.'); return; }
    const { lat, lng } = metrics.centroid;
    const origin = new google.maps.LatLng(lat, lng);

    showToast('Calculating road distances…');

    // ── Promise wrapper for DirectionsService ─────────────────
    const dirSvc = new google.maps.DirectionsService();
    function roadDistKm(destLat, destLng) {
      return new Promise(resolve => {
        dirSvc.route({
          origin,
          destination: new google.maps.LatLng(destLat, destLng),
          travelMode: google.maps.TravelMode.DRIVING
        }, (result, status) => {
          if (status === 'OK') {
            const km = (result.routes[0].legs[0].distance.value / 1000);
            resolve(parseFloat(km.toFixed(2)));
          } else {
            resolve(null); // null = road not found, will fall back to straight-line
          }
        });
      });
    }

    // ── Find top-4 substations by straight line (pre-filter) ──
    // Use top-4 to ensure we get best 2 after road distances are sorted
    const subPool = substations
      .filter(s => s.lat != null && s.lng != null)
      .map(s => ({ name: s.name, lat: s.lat, lng: s.lng,
                   straightKm: msmHaversineKm(lat, lng, s.lat, s.lng) }))
      .sort((a, b) => a.straightKm - b.straightKm)
      .slice(0, 4);

    // ── Find top-4 transmission lines by nearest geometric point ───────────
    // Grid Conn = distance to the actual line, not to a substation
    if (!tl_loaded) {
      showToast('Loading transmission lines first…');
      await loadTransmissionLines();
    }
    const tlGroups = [
      { arr: tl_500, kv: '500 kV' },
      { arr: tl_275, kv: '275 kV' },
      { arr: tl_132, kv: '132 kV' },
      { arr: tl_33,  kv: '33 kV'  },
      { arr: tl_11,  kv: '11 kV'  },
    ];

    const linePool = [];
    tlGroups.forEach(({ arr, kv }) => {
      if (!arr.length) return;
      // Find nearest geometric point on this voltage class's lines
      let bestDist = Infinity, bestLat = null, bestLng = null;
      arr.forEach(polyline => {
        const path = polyline.getPath().getArray();
        for (let i = 0; i < path.length - 1; i++) {
          const ax = path[i].lat(),   ay = path[i].lng();
          const bx = path[i+1].lat(), by = path[i+1].lng();
          const deg2km = 111.32;
          const cosLat = Math.cos(lat * Math.PI / 180);
          const px = (lng - ay) * deg2km * cosLat;
          const py = (lat - ax) * deg2km;
          const dx = (by - ay) * deg2km * cosLat;
          const dy = (bx - ax) * deg2km;
          const len2 = dx*dx + dy*dy;
          const t = len2 > 0 ? Math.max(0, Math.min(1, (px*dx + py*dy) / len2)) : 0;
          const nearLat = ax + t*(bx - ax);
          const nearLng = ay + t*(by - ay);
          const d = msmHaversineKm(lat, lng, nearLat, nearLng);
          if (d < bestDist){ bestDist = d; bestLat = nearLat; bestLng = nearLng; }
        }
      });
      if (bestLat !== null) {
        linePool.push({ kv, lat: bestLat, lng: bestLng, straightKm: bestDist });
      }
    });
    linePool.sort((a, b) => a.straightKm - b.straightKm);
    const top4Lines = linePool.slice(0, 4);

    if (!subPool.length && !top4Lines.length) {
      showToast('No substations or lines loaded. Check data.');
      return;
    }

    // ── Get real road distances for all candidates in parallel ─
    const [sub1Road, sub2Road, sub3Road, sub4Road,
           line1Road, line2Road, line3Road, line4Road] = await Promise.all([
      subPool[0] ? roadDistKm(subPool[0].lat, subPool[0].lng) : Promise.resolve(null),
      subPool[1] ? roadDistKm(subPool[1].lat, subPool[1].lng) : Promise.resolve(null),
      subPool[2] ? roadDistKm(subPool[2].lat, subPool[2].lng) : Promise.resolve(null),
      subPool[3] ? roadDistKm(subPool[3].lat, subPool[3].lng) : Promise.resolve(null),
      top4Lines[0] ? roadDistKm(top4Lines[0].lat, top4Lines[0].lng) : Promise.resolve(null),
      top4Lines[1] ? roadDistKm(top4Lines[1].lat, top4Lines[1].lng) : Promise.resolve(null),
      top4Lines[2] ? roadDistKm(top4Lines[2].lat, top4Lines[2].lng) : Promise.resolve(null),
      top4Lines[3] ? roadDistKm(top4Lines[3].lat, top4Lines[3].lng) : Promise.resolve(null),
    ]);

    // Attach road distances, fall back to straight-line if API failed
    const subRoads = [sub1Road, sub2Road, sub3Road, sub4Road];
    subPool.forEach((s, i) => {
      s.roadKm = subRoads[i] !== null ? subRoads[i] : s.straightKm;
      s.isStraight = subRoads[i] === null;
    });
    const lineRoads = [line1Road, line2Road, line3Road, line4Road];
    top4Lines.forEach((l, i) => {
      l.roadKm = lineRoads[i] !== null ? lineRoads[i] : l.straightKm;
      l.isStraight = lineRoads[i] === null;
    });

    // Re-sort by road distance, take best 2 of each
    subPool.sort((a, b) => a.roadKm - b.roadKm);
    top4Lines.sort((a, b) => a.roadKm - b.roadKm);
    const top2Subs  = subPool.slice(0, 2);
    const top2Lines = top4Lines.slice(0, 2);

    // ── Populate modal with road distances ─────────────────────
    _gridCandidates = { subs: top2Subs, lines: top2Lines };

    const s1 = top2Subs[0]  || null, s2 = top2Subs[1]  || null;
    const l1 = top2Lines[0] || null, l2 = top2Lines[1] || null;

    const setOpt = (nameId, distId, optId, label, cand, show) => {
      const el = document.getElementById(optId);
      if (el) el.style.display = show ? '' : 'none';
      const nm = document.getElementById(nameId);
      const ds = document.getElementById(distId);
      if (nm) nm.textContent = label || '—';
      if (ds && cand) {
        ds.textContent = cand.roadKm.toFixed(2) + ' km by road'
          + (cand.isStraight ? ' (est.)' : '');
      } else if (ds) {
        ds.textContent = '';
      }
    };

    setOpt('subOpt1Name',  'subOpt1Dist',  'subOpt1',  s1?.name,  s1,  !!s1);
    setOpt('subOpt2Name',  'subOpt2Dist',  'subOpt2',  s2?.name,  s2,  !!s2);
    setOpt('lineOpt1Kv',   'lineOpt1Dist', 'lineOpt1', l1 ? l1.kv : '', l1, !!l1);
    setOpt('lineOpt2Kv',   'lineOpt2Dist', 'lineOpt2', l2 ? l2.kv : '', l2, !!l2);

    // Default select first of each
    const r1 = document.querySelector('input[name="subChoice"][value="1"]');
    const r3 = document.querySelector('input[name="lineChoice"][value="1"]');
    if (r1) r1.checked = true;
    if (r3) r3.checked = true;

    // Highlight selected on click
    document.querySelectorAll('.msm-choice-option').forEach(opt => {
      opt.addEventListener('click', () => {
        const grp = opt.querySelector('input').name;
        document.querySelectorAll('.msm-choice-option').forEach(o => {
          if (o.querySelector('input').name === grp) o.classList.remove('selected');
        });
        opt.classList.add('selected');
      });
    });
    document.getElementById('subOpt1')?.classList.add('selected');
    document.getElementById('lineOpt1')?.classList.add('selected');

    setConnMode('pmu');
    document.getElementById('gridChoiceModal').classList.add('show');
  }

  // ── Custom pin helper functions ───────────────────────────
  function msmStartPinDrop(shapeRef) {
    _pinDropActive   = true;
    _pinDropShapeRef = shapeRef;
    _customPinLatLng = null;
    document.body.classList.add('msm-pin-drop-mode');
    document.getElementById('pinDropBanner').classList.add('show');
    document.getElementById('gridChoiceModal').classList.remove('show');
  }

  function msmEndPinDrop() {
    _pinDropActive = false;
    document.body.classList.remove('msm-pin-drop-mode');
    document.getElementById('pinDropBanner').classList.remove('show');
  }

  function msmClearCustomPin() {
    if (_customPinMarker) { _customPinMarker.setMap(null); _customPinMarker = null; }
    _customPinLatLng = null;
  }

  // ── Connection type switcher — PMU vs Grid Line (OUTER scope so msmFindNearestConnection can call it) ──
  let _connMode = 'pmu'; // 'pmu' or 'grid'
  function setConnMode(mode) {
    _connMode = mode;
    const pmuPanel  = document.getElementById('connPanelPmu');
    const gridPanel = document.getElementById('connPanelGrid');
    const pmuBtn    = document.getElementById('connTypePmu');
    const gridBtn   = document.getElementById('connTypeGrid');
    if (pmuPanel)  pmuPanel.style.display  = mode === 'pmu'  ? '' : 'none';
    if (gridPanel) gridPanel.style.display = mode === 'grid' ? '' : 'none';
    if (pmuBtn)  { pmuBtn.classList.toggle('secondary', mode !== 'pmu');
                   pmuBtn.classList.toggle('filter-btn', true); }
    if (gridBtn) { gridBtn.classList.toggle('secondary', mode !== 'grid');
                   gridBtn.classList.toggle('filter-btn', true); }
  }

  // ── Wire modal confirm/cancel + custom pin ─────────────────
  document.addEventListener('DOMContentLoaded', () => {

    document.getElementById('connTypePmu')?.addEventListener('click',  () => setConnMode('pmu'));
    document.getElementById('connTypeGrid')?.addEventListener('click', () => setConnMode('grid'));

    // Show/hide search box when custom radio is selected
    document.querySelectorAll('input[name="subChoice"]').forEach(radio => {
      radio.addEventListener('change', () => {
        const wrap = document.getElementById('customSubSearchWrap');
        if (wrap) wrap.classList.toggle('show', radio.value === 'custom' && radio.checked);
        if (radio.value === 'custom' && radio.checked) {
          setTimeout(() => document.getElementById('customSubSearchInput')?.focus(), 80);
        }
        // Reset picked state
        _customSubPicked = null;
        const sel = document.getElementById('customSubSelected');
        if (sel) sel.classList.remove('show');
      });
    });

    // Live search in substations array
    const searchInp = document.getElementById('customSubSearchInput');
    const dropdown  = document.getElementById('customSubDropdown');
    const selLabel  = document.getElementById('customSubSelected');
    const pinHint   = document.getElementById('customPinHint');

    function msmSubSearch(query) {
      if (!dropdown) return;
      dropdown.innerHTML = '';
      _customSubPicked = null;
      if (selLabel) selLabel.classList.remove('show');

      const q = query.trim().toLowerCase();
      if (!q) { dropdown.classList.remove('show'); return; }

      // Filter substations
      const matches = substations
        .filter(s => s.name && s.name.toLowerCase().includes(q))
        .slice(0, 12); // cap at 12 results

      if (!matches.length) {
        dropdown.innerHTML = '<div class="custom-sub-no-results">No match — use pin drop below</div>';
        dropdown.classList.add('show');
        if (pinHint) pinHint.classList.add('show');
        return;
      }

      if (pinHint) pinHint.classList.remove('show');

      // Compute distance for each match (need active shape centroid)
      const metrics = activeShape ? polygonMetrics(activeShape.gobj) : null;
      matches.forEach(sub => {
        const item = document.createElement('div');
        item.className = 'custom-sub-item';
        const distKm = metrics
          ? msmHaversineKm(metrics.centroid.lat, metrics.centroid.lng, sub.lat, sub.lng)
          : null;
        item.innerHTML = '<strong>' + escapeHtml(sub.name) + '</strong>'
          + (distKm != null ? '<small>' + distKm.toFixed(2) + ' km away</small>' : '');
        item.addEventListener('mousedown', e => {
          e.preventDefault(); // prevent blur before click fires
          _customSubPicked = { name: sub.name, distKm };
          if (searchInp) searchInp.value = sub.name;
          dropdown.classList.remove('show');
          if (selLabel) {
            selLabel.textContent = '✓ Selected: ' + sub.name
              + (distKm != null ? '  (' + distKm.toFixed(2) + ' km)' : '');
            selLabel.classList.add('show');
          }
          if (pinHint) pinHint.classList.remove('show');
        });
        dropdown.appendChild(item);
      });
      dropdown.classList.add('show');
    }

    if (searchInp) {
      searchInp.addEventListener('input', () => msmSubSearch(searchInp.value));
      searchInp.addEventListener('focus', () => { if (searchInp.value) msmSubSearch(searchInp.value); });
      searchInp.addEventListener('blur',  () => setTimeout(() => dropdown?.classList.remove('show'), 150));
    }

    // Switch to pin-drop link
    document.getElementById('switchToPinDrop')?.addEventListener('click', e => {
      e.preventDefault();
      msmStartPinDrop(activeShape);
      showToast('Click the map to place the PMU location.');
    });

    document.getElementById('gridChoiceConfirmBtn')?.addEventListener('click', () => {
      if (!activeShape || !_gridCandidates) return;
      if (!msmIsUnlocked()) { showToast('Unlock editing first.'); return; }

      if (_connMode === 'pmu') {
        // ── PMU / Substation connection ──────────────────────────
        // Clear any previous grid-line connection
        activeShape.connVoltage = '';
        activeShape.connDistKm  = '';

        const subVal = document.querySelector('input[name="subChoice"]:checked')?.value;

        if (subVal === 'custom') {
          if (_customSubPicked) {
            activeShape.nearestSub     = _customSubPicked.name;
            activeShape.nearestSubDist = _customSubPicked.distKm != null
              ? String(_customSubPicked.distKm) : '';
            document.getElementById('gridChoiceModal').classList.remove('show');
            msmUpdateSolarPanel(activeShape);
            refreshBoundarySummary();
            if (typeof msmDrawConnLines === 'function') msmDrawConnLines(activeShape);
            showToast('PMU connection saved — click Save to persist.');
          } else {
            msmStartPinDrop(activeShape);
            showToast('Click the map to place the PMU location.');
          }
          return;
        }

        const chosenSub = _gridCandidates.subs[subVal === '2' ? 1 : 0];
        activeShape.nearestSub     = chosenSub ? chosenSub.name           : '';
        activeShape.nearestSubDist = chosenSub ? String(chosenSub.roadKm) : '';

      } else {
        // ── Grid line connection ─────────────────────────────────
        // Clear any previous PMU connection
        activeShape.nearestSub     = '';
        activeShape.nearestSubDist = '';

        const lineVal = document.querySelector('input[name="lineChoice"]:checked')?.value;
        const chosenLine = _gridCandidates.lines[lineVal === '2' ? 1 : 0];
        activeShape.connVoltage = chosenLine ? chosenLine.kv             : '';
        activeShape.connDistKm  = chosenLine ? String(chosenLine.roadKm) : '';
      }

      document.getElementById('gridChoiceModal').classList.remove('show');
      msmUpdateSolarPanel(activeShape);
      refreshBoundarySummary();
      if (typeof msmDrawConnLines === 'function') msmDrawConnLines(activeShape);
      showToast('Connection saved — click Save to persist.');
    });

    document.getElementById('gridChoiceCancelBtn')?.addEventListener('click', () => {
      document.getElementById('gridChoiceModal').classList.remove('show');
    });

    // Cancel pin-drop
    document.getElementById('pinDropCancelBtn')?.addEventListener('click', () => {
      msmEndPinDrop();
      msmClearCustomPin();
      showToast('Custom pin cancelled.');
    });
  });

  // ── Intercept map click for pin-drop mode ─────────────────
  // Injected into initMap after map is created — see below


  // ===== CONNECTION LINE VISUALISER =====
  let _connLines    = [];   // active google.maps.Polyline objects
  let _connMarkers  = [];   // endpoint dot markers

  function msmClearConnLines(){
    _connLines.forEach(l => l.setMap(null));
    _connMarkers.forEach(m => m.setMap(null));
    if (typeof _connRenderers !== 'undefined'){
      _connRenderers.forEach(r => r.setMap(null));
      _connRenderers = [];
    }
    _connLines   = [];
    _connMarkers = [];
  }

  // Find nearest point (lat,lng) on any loaded polyline in lineArrays
  function msmNearestPointOnLines(lat, lng, lineArrays){
    let best = { distKm: Infinity, lat: null, lng: null, kv: null };
    lineArrays.forEach(({ arr, kv }) => {
      arr.forEach(polyline => {
        const path = polyline.getPath().getArray();
        for (let i = 0; i < path.length - 1; i++){
          const ax = path[i].lat(),   ay = path[i].lng();
          const bx = path[i+1].lat(), by = path[i+1].lng();
          const deg2km = 111.32;
          const cosLat = Math.cos(lat * Math.PI / 180);
          const px = (lng - ay) * deg2km * cosLat;
          const py = (lat - ax) * deg2km;
          const dx = (by - ay) * deg2km * cosLat;
          const dy = (bx - ax) * deg2km;
          const len2 = dx*dx + dy*dy;
          const t = len2 > 0 ? Math.max(0, Math.min(1, (px*dx + py*dy) / len2)) : 0;
          const nearLat = ax + t*(bx - ax);
          const nearLng = ay + t*(by - ay);
          const d = msmHaversineKm(lat, lng, nearLat, nearLng);
          if (d < best.distKm){
            best = { distKm: d, lat: nearLat, lng: nearLng, kv };
          }
        }
      });
    });
    return best.lat != null ? best : null;
  }

  // Each connection gets its own DirectionsRenderer so routes don't clash
  let _connRenderers = [];

  function msmDrawConnLines(shape){
    msmClearConnLines();
    if (!shape || !shape.gobj) return;
    const metrics = polygonMetrics(shape.gobj);
    if (!metrics) return;
    const { lat, lng } = metrics.centroid;
    const origin = new google.maps.LatLng(lat, lng);
    const svc = new google.maps.DirectionsService();

    // Helper: render one road route + endpoint marker + distance label
    // onRoadKm(km, mins) callback fires after route resolves — used to update table
    function drawRoute(destLatLng, color, labelPrefix, onRoadKm){
      const renderer = new google.maps.DirectionsRenderer({
        map,
        suppressMarkers: true,
        polylineOptions: {
          strokeColor: color,
          strokeWeight: 4,
          strokeOpacity: 0.85,
          zIndex: 500
        }
      });
      _connRenderers.push(renderer);

      // Endpoint marker
      const dot = new google.maps.Marker({
        position: destLatLng,
        map,
        icon: {
          path: google.maps.SymbolPath.CIRCLE,
          scale: 8,
          fillColor: color,
          fillOpacity: 1,
          strokeColor: '#fff',
          strokeWeight: 2.5
        },
        zIndex: 510
      });
      _connMarkers.push(dot);

      svc.route({
        origin,
        destination: destLatLng,
        travelMode: google.maps.TravelMode.DRIVING
      }, (result, status) => {
        if (status === 'OK') {
          renderer.setDirections(result);
          const leg  = result.routes[0].legs[0];
          const km   = (leg.distance.value / 1000).toFixed(2);
          const mins = Math.round(leg.duration.value / 60);

          // Update stored distance with real road km, refresh table
          if (typeof onRoadKm === 'function') onRoadKm(km, mins);

          // Label at route midpoint
          const steps  = leg.steps;
          const mid    = steps[Math.floor(steps.length / 2)];
          const midPos = mid ? mid.start_location : destLatLng;

          const labelMarker = new google.maps.Marker({
            position: midPos,
            map,
            icon: { path: google.maps.SymbolPath.CIRCLE, scale: 0 },
            label: {
              text: labelPrefix + '  ' + km + ' km (' + mins + ' min)',
              color: '#fff',
              fontSize: '10px',
              fontWeight: '700'
            },
            zIndex: 511
          });
          _connMarkers.push(labelMarker);
        } else {
          // Road route unavailable — straight dashed fallback
          const fallback = new google.maps.Polyline({
            path: [{ lat, lng }, { lat: destLatLng.lat(), lng: destLatLng.lng() }],
            map,
            geodesic: true,
            strokeColor: color,
            strokeOpacity: 0,
            icons: [{ icon: { path: 'M 0,-1 0,1', strokeOpacity: 0.9, scale: 3 },
                      offset: '0', repeat: '14px' }],
            strokeWeight: 2,
            zIndex: 500
          });
          _connLines.push(fallback);
          showToast('Road route unavailable — showing straight-line distance.');
        }
      });
    }

    // ── Route to substation ──────────────────────────────────────────────────
    if (shape.nearestSub) {
      const sub = substations.find(s =>
        s.name && s.name.toLowerCase() === shape.nearestSub.toLowerCase()
      );
      if (sub && sub.lat != null && sub.lng != null) {
        drawRoute(
          new google.maps.LatLng(sub.lat, sub.lng),
          '#00E5FF',
          shape.nearestSub.split(' ').slice(0, 4).join(' '),
          (km) => {
            // Update table distance with real road km
            shape.nearestSubDist = km;
            refreshBoundarySummary();
            msmUpdateSolarPanel(shape);
          }
        );
      }
    }

    // ── Route to nearest geometric point on the chosen transmission line ──────
    // Grid Conn = actual nearest point on the voltage line, not a substation
    if (shape.connVoltage && tl_loaded) {
      const kvColors = {
        '500 kV': '#C97125', '275 kV': '#00A8FF',
        '132 kV': '#FF1744', '33 kV':  '#FF9800', '11 kV': '#9C27B0'
      };
      const col = kvColors[shape.connVoltage] || '#FF6D00';

      const tlGroupsForDraw = [
        { arr: tl_500, kv: '500 kV' }, { arr: tl_275, kv: '275 kV' },
        { arr: tl_132, kv: '132 kV' }, { arr: tl_33,  kv: '33 kV'  },
        { arr: tl_11,  kv: '11 kV'  },
      ].filter(g => g.kv === shape.connVoltage);

      const nearest = msmNearestPointOnLines(lat, lng, tlGroupsForDraw);
      if (nearest) {
        drawRoute(
          new google.maps.LatLng(nearest.lat, nearest.lng),
          col,
          shape.connVoltage,
          (km) => {
            shape.connDistKm = km;
            refreshBoundarySummary();
            msmUpdateSolarPanel(shape);
          }
        );
      }
    }
  }

  window.initMap = initMap;

  // ===== Sidebar Tab Switching =====
  document.querySelectorAll('.sidebar-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.sidebar-tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      tab.classList.add('active');
      const target = document.getElementById(tab.dataset.tab);
      if (target) target.classList.add('active');
    });
  });

  // Mobile drawer controls
  const side = document.querySelector('.sidebar');
  const openBtn = document.getElementById('mobileMenuBtn');
  const closeBtn = document.getElementById('closeSidebarBtn');
  const backdrop = document.getElementById('backdrop');

  function openDrawer(){
    side.classList.add('open');
    backdrop && backdrop.classList.add('show');
  }
  function closeDrawer(){
    side.classList.remove('open');
    backdrop && backdrop.classList.remove('show');
  }
  openBtn?.addEventListener('click', openDrawer);
  closeBtn?.addEventListener('click', closeDrawer);
  backdrop?.addEventListener('click', closeDrawer);
  window.addEventListener('keydown', (e)=>{ if(e.key==='Escape') closeDrawer(); });
  
  
  // ===== Saved Lands header dropdown filters =====

// Build the dropdown options based on current table values
function msmBuildSavedLandFilters() {
  const tbody = document.getElementById('boundarySummaryBody');
  if (!tbody) return;

  const rows = Array.from(tbody.querySelectorAll('tr'));

  const names  = new Set();
  const owners = new Set();
  const states = new Set();
  const cities = new Set();
  const rents  = new Set();

  rows.forEach(row => {
    const nameInput  = row.querySelector('.land-name-input');
    const ownerInput = row.querySelector('.land-owner-input');
    const stateInput = row.querySelector('.land-state-input');
    const cityInput  = row.querySelector('.land-city-input');
    const rentInput  = row.querySelector('.land-rent-input');

    if (!nameInput && !ownerInput && !stateInput && !cityInput && !rentInput) {
      return; // skip "No lands drawn yet" row
    }

    const vName  = nameInput  ? nameInput.value.trim()  : '';
    const vOwner = ownerInput ? ownerInput.value.trim() : '';
    const vState = stateInput ? stateInput.value.trim() : '';
    const vCity  = cityInput  ? cityInput.value.trim()  : '';
    const vRent  = rentInput  ? rentInput.value.trim()  : '';

    if (vName)  names.add(vName);
    if (vOwner) owners.add(vOwner);
    if (vState) states.add(vState);
    if (vCity)  cities.add(vCity);
    if (vRent)  rents.add(vRent);
  });

  function fillSelect(id, values) {
    const sel = document.getElementById(id);
    if (!sel) return;

    const previous = sel.value;
    sel.innerHTML = '';

    const optAll = document.createElement('option');
    optAll.value = '';
    optAll.textContent = 'All';
    sel.appendChild(optAll);

    Array.from(values).sort((a, b) => a.localeCompare(b)).forEach(v => {
      const opt = document.createElement('option');
      opt.value = v;
      opt.textContent = v;
      sel.appendChild(opt);
    });

    if (previous && values.has(previous)) {
      sel.value = previous;
    }
  }

  fillSelect('filterLandName',  names);
  fillSelect('filterLandOwner', owners);
  fillSelect('filterState',     states);
  fillSelect('filterCity',      cities);
  fillSelect('filterRent',      rents);
}

// Apply filtering based on selected dropdown values
function msmApplySavedLandDropdownFilter() {
  const tbody = document.getElementById('boundarySummaryBody');
  if (!tbody) return;

  const valStatus = (document.getElementById('filterStatus')    || {}).value || '';
  const valName   = (document.getElementById('filterLandName')  || {}).value || '';
  const valOwner  = (document.getElementById('filterLandOwner') || {}).value || '';
  const valState  = (document.getElementById('filterState')     || {}).value || '';
  const valCity   = (document.getElementById('filterCity')      || {}).value || '';
  const valRent   = (document.getElementById('filterRent')      || {}).value || '';

  const anyFilter = valStatus || valName || valOwner || valState || valCity || valRent;

  Array.from(tbody.querySelectorAll('tr')).forEach(row => {
    const nameInput  = row.querySelector('.land-name-input');
    const ownerInput = row.querySelector('.land-owner-input');
    const stateInput = row.querySelector('.land-state-input');
    const cityInput  = row.querySelector('.land-city-input');
    const rentInput  = row.querySelector('.land-rent-input');

    if (!nameInput && !ownerInput && !stateInput && !cityInput && !rentInput) {
      row.style.display = anyFilter ? 'none' : '';
      return;
    }

    // Find shape by id to check status
    const id = row.dataset.shapeId;
    const shape = drawnShapes.find(s => s.id === id);
    const rowStatus = shape ? (shape.status || 'Prospecting') : '';

    let show = true;
    if (valStatus && rowStatus !== valStatus)                                 show = false;
    if (valName  && (!nameInput  || nameInput.value.trim()  !== valName))    show = false;
    if (valOwner && (!ownerInput || ownerInput.value.trim() !== valOwner))   show = false;
    if (valState && (!stateInput || stateInput.value.trim() !== valState))   show = false;
    if (valCity  && (!cityInput  || cityInput.value.trim()  !== valCity))    show = false;
    if (valRent  && (!rentInput  || rentInput.value.trim()  !== valRent))    show = false;

    row.style.display = show ? '' : 'none';
  });
}

// Attach filter change listeners when page is ready
(function () {
  document.addEventListener('DOMContentLoaded', function () {
    ['filterStatus','filterLandName','filterLandOwner','filterState','filterCity','filterRent']
      .forEach(id => {
        const sel = document.getElementById(id);
        if (sel) {
          sel.addEventListener('change', () => {
            sel.classList.toggle('filter-active', !!sel.value);
            msmApplySavedLandDropdownFilter();
          });
        }
      });
    const clearBtn = document.getElementById('clearFiltersBtn');
    if (clearBtn) clearBtn.addEventListener('click', () => {
      ['filterStatus','filterLandName','filterLandOwner','filterState','filterCity','filterRent']
        .forEach(id => {
          const sel = document.getElementById(id);
          if (sel) { sel.value = ''; sel.classList.remove('filter-active'); }
        });
      msmApplySavedLandDropdownFilter();
    });
  });
})();

  
</script>

<!-- Google Maps -->
<script async defer
  src="https://maps.googleapis.com/maps/api/js?key=AIzaSyBlKy1gqkaSRAYBrPMQB8V0mTDXeVI0V78&libraries=places,geometry,drawing&callback=initMap">
</script>
</body>
</html>
