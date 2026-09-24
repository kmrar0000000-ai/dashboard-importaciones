import{a as f,b as l}from"./chunk-P4PAYY2C.js";import{M as v,Rb as i,Vb as h,bc as x}from"./chunk-KINISEJQ.js";var P=(()=>{class c{imprimirIMP(n,o){let s=window.open("","_blank","width=900,height=700");if(!s){alert("El navegador bloque\xF3 la ventana emergente. Habilita las ventanas emergentes para este sitio.");return}s.document.write(this.htmlIMP(n,o)),s.document.close(),setTimeout(()=>s.print(),350)}excelIMP(n,o){let s=o.map(a=>({id_linea:a.id,OCI:a.oci,"C\xF3digo SAP":a.sap,"C\xF3digo producto":a.cprod,Descripci\u00F3n:a.desc,Proveedor:a.prov,Cantidad:a.cant,"Precio total":a.precio,Tipo:a.Tipo,Estado:a.st==="O"?"Abierta":"Cerrada","Estado f\xE1brica":a.sfab,"F. OCI":i(a.foci),"F. embarque":i(a.femb),"F. llegada puerto":i(a.fll),"F. estimada ingreso":i(a.fEst),"F. ingreso Interamsa":i(a.ing),"Plazo entrega":i(a.plazo),Facturas:a.facturas.join(" | "),Comentario:a.com,"Prioridad / proyecto":a.prior})),p=l.json_to_sheet(s),r=l.book_new(),e=n.replace(/[\\/?*[\]]/g,"-").slice(0,31);l.book_append_sheet(r,p,e),f(r,`${n.replace(/\s+/g,"_")}.xlsx`)}excelLineas(n,o){if(!n.length){alert("No hay l\xEDneas que exportar con los filtros actuales.");return}let s=n.map(e=>({id_linea:e.id,OCI:e.oci,"C\xF3digo SAP":e.sap,"C\xF3digo producto":e.cprod,Descripci\u00F3n:e.desc,Proveedor:e.prov,Cantidad:e.cant,"Precio total":e.precio,Tipo:e.Tipo,Estado:e.st==="O"?"Abierta":"Cerrada","Estado f\xE1brica":e.sfab,IMP:e.IMP,"F. OCI":i(e.foci),"F. embarque":i(e.femb),"F. llegada puerto":i(e.fll),"F. estimada ingreso":i(e.fEst),"F. ingreso Interamsa":i(e.ing),"Plazo entrega":i(e.plazo),Facturas:e.facturas.join(" | "),Comentario:e.com,"Prioridad / proyecto":e.prior})),p=l.json_to_sheet(s),r=l.book_new();l.book_append_sheet(r,p,"L\xEDneas"),f(r,`${o}.xlsx`)}htmlIMP(n,o){let s=h(o),p=x(o[0]),r=o.filter(t=>t.st==="O").length,e=o.filter(t=>t.st==="C").length,a=o.reduce((t,b)=>t+(b.cant||0),0),u=o.reduce((t,b)=>t+(b.precio||0),0),m=[...new Set(o.map(t=>t.oci))],w=[...new Set(o.map(t=>t.prov).filter(Boolean))],g=[...new Set(o.flatMap(t=>t.facturas))],d=t=>String(t??"\u2014").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"),y=o.map(t=>`<tr>
        <td class="m">${d(t.id)}</td>
        <td class="m">${d(t.oci)}</td>
        <td>${d(t.desc)}</td>
        <td class="m r">${(t.cant||0).toLocaleString("es-PE")}</td>
        <td>${t.st==="O"?"Abierta":"Cerrada"}</td>
        <td>${d(t.sfab)}</td>
        <td class="m">${i(t.fll)}</td>
        <td class="m">${i(t.fEst)}</td>
        <td>${d(t.com)}</td>
      </tr>`).join(""),$=p.map(t=>`<div class="tl-s">
        <div class="tl-d ${t.ok?"ok":t.warn?"wa":t.date?"pe":"na"}">${t.ok?"\u2713":t.warn?"!":"\u25CB"}</div>
        <div class="tl-l">${d(t.label)}</div>
        <div class="tl-f">${i(t.date)}</div>
      </div>`).join(""),k=s.level==="g"?"En orden":s.level==="y"?"Requiere atenci\xF3n":"Cr\xEDtico";return`<!doctype html><html lang="es"><head><meta charset="utf-8">
<title>${d(n)} \u2014 Interamsa</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:"Segoe UI",system-ui,sans-serif;color:#16232f;padding:28px 30px;font-size:11px}
  .hd{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:2px solid #0d1b2a;padding-bottom:12px;margin-bottom:16px}
  .hd h1{font-size:20px;letter-spacing:-.01em}
  .hd .co{font-size:11px;color:#61748a;margin-top:2px}
  .hd .fe{font-size:10px;color:#61748a;text-align:right}
  .est{display:inline-block;padding:3px 10px;border-radius:4px;font-size:10px;font-weight:700;margin-top:6px}
  .est.g{background:#d8f2e8;color:#0b7a5a}
  .est.y{background:#fdf0d8;color:#b26a00}
  .est.r{background:#fde8e8;color:#c02626}
  .kpis{display:flex;gap:22px;margin-bottom:16px;flex-wrap:wrap}
  .k .v{font-size:19px;font-weight:700;font-variant-numeric:tabular-nums}
  .k .l{font-size:9px;color:#61748a;text-transform:uppercase;letter-spacing:.06em;margin-top:1px}
  .meta{font-size:10px;color:#61748a;margin-bottom:16px;line-height:1.6}
  .meta b{color:#16232f}
  .tl{display:flex;align-items:flex-start;gap:6px;margin-bottom:18px;padding:12px;background:#f7f9fb;border-radius:6px}
  .tl-s{text-align:center;min-width:84px}
  .tl-d{width:24px;height:24px;border-radius:50%;margin:0 auto;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:11px;border:2px solid}
  .tl-d.ok{background:#d8f2e8;color:#0b7a5a;border-color:#0b7a5a}
  .tl-d.wa{background:#fde8e8;color:#c02626;border-color:#c02626}
  .tl-d.pe{background:#e6f0fa;color:#0b5ea8;border-color:#0b5ea8}
  .tl-d.na{background:#eef2f6;color:#8ea2b8;border-color:#b9cbe0}
  .tl-l{font-size:8px;text-transform:uppercase;letter-spacing:.05em;color:#61748a;font-weight:700;margin-top:5px}
  .tl-f{font-size:9px;margin-top:2px;font-variant-numeric:tabular-nums}
  h2{font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:#61748a;margin-bottom:8px;padding-bottom:5px;border-bottom:1px solid #dbe3ec}
  table{width:100%;border-collapse:collapse;font-size:9.5px}
  th{background:#f7f9fb;text-align:left;padding:6px 7px;font-size:8.5px;text-transform:uppercase;letter-spacing:.05em;color:#61748a;border-bottom:1px solid #dbe3ec}
  td{padding:5px 7px;border-bottom:1px solid #eef2f6;vertical-align:top}
  td.m{font-family:ui-monospace,monospace;font-size:9px;white-space:nowrap}
  td.r{text-align:right}
  tbody tr:nth-child(even){background:#fbfcfe}
  .ft{margin-top:20px;padding-top:10px;border-top:1px solid #dbe3ec;font-size:9px;color:#8ea2b8;display:flex;justify-content:space-between}
  @media print{ body{padding:0} thead{display:table-header-group} tr{page-break-inside:avoid} }
</style></head><body>
  <div class="hd">
    <div>
      <h1>${d(n)}</h1>
      <div class="co">Inversiones Interamericanas S.A. \xB7 Control de Importaciones</div>
      <span class="est ${s.level}">${k} \u2014 ${d(s.label)}</span>
    </div>
    <div class="fe">Emitido<br><b>${new Date().toLocaleString("es-PE")}</b></div>
  </div>

  <div class="kpis">
    <div class="k"><div class="v">${o.length}</div><div class="l">L\xEDneas</div></div>
    <div class="k"><div class="v">${m.length}</div><div class="l">OCIs</div></div>
    <div class="k"><div class="v" style="color:#0b5ea8">${r}</div><div class="l">Abiertas</div></div>
    <div class="k"><div class="v" style="color:#0b7a5a">${e}</div><div class="l">Cerradas</div></div>
    <div class="k"><div class="v">${a.toLocaleString("es-PE")}</div><div class="l">Unidades</div></div>
    <div class="k"><div class="v">${u.toLocaleString("es-PE",{style:"currency",currency:"USD",maximumFractionDigits:0})}</div><div class="l">Valor</div></div>
  </div>

  <div class="meta">
    <b>OCIs:</b> ${d(m.join(", "))}<br>
    <b>Proveedor(es):</b> ${d(w.join(" \xB7 "))}<br>
    <b>Factura(s):</b> ${d(g.length?g.join(" \xB7 "):"\u2014 sin factura asociada \u2014")}
  </div>

  <h2>Trazabilidad del embarque</h2>
  <div class="tl">${$}</div>

  <h2>Detalle de l\xEDneas (${o.length})</h2>
  <table>
    <thead><tr>
      <th>id_linea</th><th>OCI</th><th>Descripci\xF3n</th><th>Cant.</th>
      <th>Estado</th><th>F\xE1brica</th><th>Llegada</th><th>Est. ingreso</th><th>Comentario</th>
    </tr></thead>
    <tbody>${y}</tbody>
  </table>

  <div class="ft">
    <span>Documento generado autom\xE1ticamente desde el Panel de Importaciones.</span>
    <span>F. estimada = fecha de llegada + 7 d\xEDas</span>
  </div>
</body></html>`}static \u0275fac=function(o){return new(o||c)};static \u0275prov=v({token:c,factory:c.\u0275fac,providedIn:"root"})}return c})();export{P as a};
