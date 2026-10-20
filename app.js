// Agenda Planilla · app
(() => {
"use strict";
const $ = s => document.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const pad = n => String(n).padStart(2, "0");
const iso = d => `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
const fromIso = s => { const [y,m,d] = s.split("-").map(Number); return new Date(y, m-1, d); };
const addDays = (d, n) => { const x = new Date(d); x.setDate(x.getDate()+n); return x; };
const TODAY = iso(new Date());
const MESES = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
const DIAS = ["dom","lun","mar","mié","jue","vie","sáb"];
const DIAS_L = ["domingos","lunes","martes","miércoles","jueves","viernes","sábados"];
const fmtDay = s => { const d = fromIso(s); return `${DIAS[d.getDay()]} ${d.getDate()} ${MESES[d.getMonth()].slice(0,3)}`; };
const norm = s => String(s ?? "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
const local = {
  get(k, def){ try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : def; } catch { return def; } },
  set(k, v){ try { localStorage.setItem(k, JSON.stringify(v)); } catch {} }
};
const TZ = "America/Argentina/Buenos_Aires";

// ---------- Supabase ----------
const CFG = window.APP_CONFIG || {};
if (!window.supabase || !CFG.SUPABASE_URL || /TU-PROYECTO/.test(CFG.SUPABASE_URL)) {
  $("#loading").textContent = "Falta configurar la conexión (js/config.js).";
  return;
}
const sb = window.supabase.createClient(CFG.SUPABASE_URL, CFG.SUPABASE_ANON_KEY);

// ---------- Estado ----------
let user = null, perfil = null;
let workbook = null, rows = [], headers = [], fileName = "";
let tasks = [], occ = [], issues = [];
let done = {};
let viewMonth = new Date(); viewMonth.setDate(1);
let sideDay = null;
let collapsed = new Set(local.get("ap.collapsed", []));

// ---------- Ingreso con código ----------
let loginEmail = "";
function loginMsg(text, cls = ""){ $("#loginMsg").innerHTML = text ? `<div class="alert ${cls}">${text}</div>` : ""; }
$("#emailForm").addEventListener("submit", async e => {
  e.preventDefault();
  loginEmail = $("#loginEmail").value.trim().toLowerCase();
  const b = $("#sendCode"); b.disabled = true; b.textContent = "Enviando…";
  const { error } = await sb.auth.signInWithOtp({ email: loginEmail, options: { shouldCreateUser: true } });
  b.disabled = false; b.textContent = "Enviarme un código";
  if (error) { loginMsg(/rate|seconds/i.test(error.message) ? "Ya te mandamos un código hace poco. Esperá un minuto y probá de nuevo." : "No pudimos enviar el código. Revisá el correo y probá de nuevo.", "err"); return; }
  $("#emailForm").hidden = true; $("#codeForm").hidden = false;
  $("#codeLbl").textContent = `Código que te llegó a ${loginEmail}`;
  loginMsg("Revisá tu correo (y la carpeta de spam).", "ok");
  $("#loginCode").focus();
});
$("#codeForm").addEventListener("submit", async e => {
  e.preventDefault();
  const token = $("#loginCode").value.replace(/\D/g, "");
  const b = $("#verifyCode"); b.disabled = true; b.textContent = "Verificando…";
  const { error } = await sb.auth.verifyOtp({ email: loginEmail, token, type: "email" });
  b.disabled = false; b.textContent = "Ingresar";
  if (error) loginMsg("El código no es correcto o venció. Pedí uno nuevo si hace falta.", "err");
});
$("#backEmail").addEventListener("click", () => { $("#codeForm").hidden = true; $("#emailForm").hidden = false; loginMsg(""); });
$("#logout").addEventListener("click", async () => { await sb.auth.signOut(); });

sb.auth.onAuthStateChange((_ev, session) => {
  const u = session && session.user;
  if (u && (!user || user.id !== u.id)) { user = u; setTimeout(start, 0); }  // fuera del callback, como pide Supabase
  else if (!u) { user = null; showLogin(); }
});
function showLogin(){
  $("#loading").hidden = true; $("#app").hidden = true; $("#login").hidden = false;
  $("#emailForm").hidden = false; $("#codeForm").hidden = true; $("#loginCode").value = ""; loginMsg("");
}

async function start(){
  $("#login").hidden = true; $("#loading").hidden = false; $("#app").hidden = true;
  await sb.from("perfiles").upsert({ user_id: user.id, email: user.email }, { onConflict: "user_id", ignoreDuplicates: true });
  const [p, t, d] = await Promise.all([
    sb.from("perfiles").select("*").eq("user_id", user.id).single(),
    sb.from("tareas").select("*").order("orden"),
    sb.from("tildes").select("tarea_id, fecha")
  ]);
  if (p.error || t.error || d.error) { $("#loading").textContent = "No pudimos cargar tu agenda. Recargá la página."; return; }
  perfil = p.data;
  tasks = (t.data || []).map(fromRow);
  done = {}; for (const x of d.data || []) done[x.tarea_id + "@" + x.fecha] = 1;
  fileName = perfil.archivo || "";
  issues = local.get("ap.issues." + user.id, []);
  $("#linkTxt").textContent = user.email; $("#linkPill").classList.add("on");
  $("#remEmail").textContent = user.email;
  $("#rDay").checked = !!perfil.aviso_dia_anterior; $("#rSame").checked = !!perfil.aviso_mismo_dia;
  $("#loading").hidden = true; $("#app").hidden = false;
  setTab(local.get("ap.tab", "agenda") === "rem" ? "rem" : "agenda");
  finish();
  if (!tasks.length) openSetup(true, false);
}
function fromRow(r){
  return { id: r.id, titulo: r.titulo, inicio: r.inicio, hora: r.hora || "", responsable: r.responsable || "", categoria: r.categoria || "",
    prioridad: r.prioridad || "", notas: r.notas || "", repetir: Object.assign({ tipo: "ninguna", cada: 1, dias_semana: [], dia_mes: null, solo_habiles: false, hasta: null }, r.repetir || {}) };
}

// ---------- Carga del Excel ----------
const drop = $("#drop"), fileInput = $("#file");
drop.addEventListener("dragover", e => { e.preventDefault(); drop.classList.add("over"); });
drop.addEventListener("dragleave", () => drop.classList.remove("over"));
drop.addEventListener("drop", e => { e.preventDefault(); drop.classList.remove("over"); if (e.dataTransfer.files[0]) readFile(e.dataTransfer.files[0]); });
fileInput.addEventListener("change", () => { if (fileInput.files[0]) readFile(fileInput.files[0]); fileInput.value = ""; });
function openSetup(open, canCancel){ $("#setup").hidden = !open; $("#closeSetup").hidden = !canCancel; }
function reject(html, cls = "err"){ $("#rejectBox").innerHTML = html ? `<div class="alert ${cls}">${html}</div>` : ""; }
$("#changeBtn").addEventListener("click", () => { openSetup(true, true); $("#preview").hidden = true; reject(""); $("#setup").scrollIntoView({ behavior: "smooth", block: "start" }); });
$("#closeSetup").addEventListener("click", () => openSetup(false));

function readFile(f){
  if (typeof XLSX === "undefined") { reject("No se pudo cargar el lector de Excel. Recargá la página."); return; }
  if (f.size > 5 * 1024 * 1024) { reject("El archivo pesa más de 5 MB. Dejá solo la hoja de tareas y volvé a subirlo."); return; }
  const r = new FileReader();
  r.onload = () => {
    try {
      workbook = XLSX.read(new Uint8Array(r.result), { type: "array", cellNF: true, raw: true });
      fileName = f.name;
      const sel = $("#sheet");
      sel.innerHTML = workbook.SheetNames.map(n => `<option>${esc(n)}</option>`).join("");
      sel.hidden = workbook.SheetNames.length < 2;
      loadSheet(workbook.SheetNames[0]);
    } catch { openSetup(true, tasks.length > 0); reject("No pude abrir el archivo. Guardalo como .xlsx y volvé a subirlo."); }
  };
  r.readAsArrayBuffer(f);
}
$("#sheet").addEventListener("change", e => loadSheet(e.target.value));
// Lee cada celda desde su número de serie de Excel (evita los segundos de desfase de las fechas)
function cellToText(cell){
  if (!cell || cell.v == null) return "";
  if (cell.t === "n" && cell.z && XLSX.SSF.is_date(cell.z)) {
    const p = XLSX.SSF.parse_date_code(cell.v);
    if (!p) return String(cell.v);
    const hm = pad(p.H) + ":" + pad(p.M);
    if (cell.v < 1) return hm;
    return `${p.y}-${pad(p.m)}-${pad(p.d)}` + ((p.H || p.M) ? " " + hm : "");
  }
  if (cell.v instanceof Date) return iso(cell.v);
  return String(cell.v).trim();
}
function sheetRows(ws){
  if (!ws || !ws["!ref"]) return [];
  const R = XLSX.utils.decode_range(ws["!ref"]), out = [];
  const maxR = Math.min(R.e.r, R.s.r + 5000), maxC = Math.min(R.e.c, R.s.c + 40);
  for (let r = R.s.r; r <= maxR; r++) {
    const row = [];
    for (let c = R.s.c; c <= maxC; c++) row.push(cellToText(ws[XLSX.utils.encode_cell({ r, c })]));
    out.push(row);
  }
  return out;
}
function loadSheet(name){
  const aoa = sheetRows(workbook.Sheets[name]).filter(r => r.some(c => c));
  const hIdx = aoa.findIndex(r => r.filter(c => c !== "").length >= 2);
  if (hIdx < 0) { openSetup(true, tasks.length > 0); reject("Esta hoja está vacía o no tiene la fila de títulos de la plantilla."); return; }
  headers = aoa[hIdx].map((h, i) => h || `Columna ${i+1}`);
  rows = aoa.slice(hIdx+1).map(r => headers.map((_, i) => r[i] || ""));
  afterLoad();
}
function showPreview(){
  const p = $("#preview"); p.hidden = false;
  p.innerHTML = `<table><thead><tr>${headers.map(h => `<th>${esc(h)}</th>`).join("")}</tr></thead><tbody>${
    rows.slice(0, 6).map(r => `<tr>${r.map(c => `<td title="${esc(c)}">${esc(c)}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
}
function goAgenda(){ openSetup(false); setTab("agenda"); $("#view-agenda").scrollIntoView({ behavior: "smooth", block: "start" }); }
async function afterLoad(){
  const map = detectColumns();
  const missing = [map.tarea == null && "Tarea", map.fecha == null && "Fecha inicio"].filter(Boolean);
  openSetup(true, tasks.length > 0);
  if (missing.length) {
    showPreview();
    reject(`<b>Este archivo no sigue la plantilla.</b> Falta la columna ${missing.map(m => "«" + m + "»").join(" y ")}. Descargá la plantilla, pegá tus tareas y volvé a subirlo.`);
    return;
  }
  const parsed = readRows(map);
  if (!parsed.tasks.length) { showPreview(); reject("La planilla sigue la plantilla pero no tiene tareas cargadas."); return; }
  reject("Guardando tus tareas…", "");
  const payload = parsed.tasks.map((t, i) => ({ ...t, orden: i }));
  const { error } = await sb.rpc("reemplazar_tareas", { p_tareas: payload, p_archivo: fileName });
  if (error) { reject("No pudimos guardar la planilla: " + esc(error.message)); return; }
  tasks = parsed.tasks; issues = parsed.issues;
  local.set("ap.issues." + user.id, issues);
  perfil.archivo = fileName;
  reject("");
  finish(); goAgenda();
}

// ---------- Lectura con la plantilla ----------
const COLS = {
  tarea: ["tarea","actividad","descripcion","titulo"],
  fecha: ["fecha","inicio","vence","vencimiento"],
  hora: ["hora","horario"],
  frecuencia: ["frecuencia","periodicidad","repeticion","recurrencia"],
  responsable: ["responsable","encargado","asignado"],
  categoria: ["area","categoria","sector"],
  prioridad: ["prioridad","urgencia"],
  notas: ["nota","notas","observacion","observaciones","comentario"]
};
function detectColumns(){
  const map = {}, used = new Set();
  for (const [role, keys] of Object.entries(COLS)) {
    const i = headers.findIndex((h, idx) => !used.has(idx) && keys.some(k => norm(h).includes(k)));
    if (i >= 0) { map[role] = i; used.add(i); }
  }
  return map;
}
function parseDate(s){
  s = String(s || "").trim(); if (!s) return null;
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})(?:[\/\-.](\d{2,4}))?/);
  if (m) { let y = m[3] ? +m[3] : new Date().getFullYear(); if (y < 100) y += 2000; const d = new Date(y, +m[2]-1, +m[1]); if (!isNaN(d) && d.getMonth() === +m[2]-1) return iso(d); }
  return null;
}
const DOW = { domingo:0, lunes:1, martes:2, miercoles:3, jueves:4, viernes:5, sabado:6 };
function parseFreq(s){
  const t = norm(s); const r = { tipo:"ninguna", cada:1, dias_semana:[], dia_mes:null, solo_habiles:false, hasta:null, ok:true };
  const days = Object.keys(DOW).filter(d => t.includes(d)).map(d => DOW[d]);
  if (!t || /unica|una vez|puntual/.test(t)) return r;
  if (/habil/.test(t)) { r.tipo = "diaria"; r.solo_habiles = true; return r; }
  if (/diari|todos los dias|cada dia/.test(t)) { r.tipo = "diaria"; return r; }
  if (/quincen/.test(t)) { r.tipo = "semanal"; r.cada = 2; r.dias_semana = days; return r; }
  if (/semest/.test(t)) { r.tipo = "mensual"; r.cada = 6; return r; }
  if (/trimest/.test(t)) { r.tipo = "mensual"; r.cada = 3; return r; }
  if (/bimest/.test(t)) { r.tipo = "mensual"; r.cada = 2; return r; }
  if (/mensual|cada mes|todos los meses/.test(t)) { r.tipo = "mensual"; return r; }
  if (/anual|cada ano/.test(t)) { r.tipo = "anual"; return r; }
  if (/semanal|cada semana|todas las semanas/.test(t) || days.length) { r.tipo = "semanal"; r.dias_semana = days; return r; }
  r.ok = false; return r;
}
function hash(s){ let h = 0; for (const c of s) h = (h * 31 + c.charCodeAt(0)) | 0; return (h >>> 0).toString(36); }
function readRows(map){
  const get = (r, k) => map[k] != null ? r[map[k]] : "";
  const out = [], warn = [];
  rows.forEach((r, i) => {
    const fila = i + 2;
    const titulo = String(get(r, "tarea")).trim().slice(0, 300);
    if (!titulo || /^ejemplo:/i.test(titulo)) return;
    const fechaTxt = get(r, "fecha");
    let inicio = parseDate(fechaTxt);
    if (!inicio) { inicio = TODAY; warn.push(`Fila ${fila} («${titulo}»): ${fechaTxt ? `la fecha «${fechaTxt}» no se entiende` : "no tiene fecha"}; se tomó hoy.`); }
    const freqTxt = get(r, "frecuencia");
    const rep = parseFreq(freqTxt);
    if (!rep.ok) warn.push(`Fila ${fila} («${titulo}»): la frecuencia «${freqTxt}» no está en la lista; se tomó como única vez.`);
    const hm = String(get(r, "hora") || (String(fechaTxt).match(/(\d{1,2}:\d{2})/) || [])[1] || "").match(/(\d{1,2}):(\d{2})/);
    const prio = norm(get(r, "prioridad"));
    out.push({
      id: "t" + hash(titulo + "|" + inicio + "|" + i),
      titulo, inicio,
      hora: hm && +hm[1] < 24 ? pad(+hm[1]) + ":" + hm[2] : "",
      responsable: String(get(r, "responsable")).trim().slice(0, 120),
      categoria: String(get(r, "categoria")).trim().slice(0, 120),
      prioridad: ["alta","media","baja"].includes(prio) ? prio : "",
      notas: String(get(r, "notas")).trim().slice(0, 1000),
      repetir: { tipo: rep.tipo, cada: rep.cada, dias_semana: rep.dias_semana, dia_mes: rep.dia_mes, solo_habiles: rep.solo_habiles, hasta: rep.hasta }
    });
  });
  return { tasks: out, issues: warn };
}

// ---------- Repeticiones ----------
function occurrences(t, end, floor){
  const r = t.repetir, out = [], start = fromIso(t.inicio);
  const until = r.hasta ? fromIso(r.hasta) : end; const stop = until < end ? until : end;
  const ok = x => x >= start && x >= floor && x <= stop && !(r.solo_habiles && (x.getDay() === 0 || x.getDay() === 6));
  if (r.tipo === "ninguna") return [t.inicio];
  let g = 0;
  if (r.tipo === "diaria") { for (let d = start; d <= stop && g < 2000; d = addDays(d, r.cada), g++) if (ok(d)) out.push(iso(d)); }
  else if (r.tipo === "semanal") {
    const dws = r.dias_semana.length ? r.dias_semana : [start.getDay()];
    for (let wk = addDays(start, -start.getDay()); wk <= stop && g < 600; wk = addDays(wk, 7 * r.cada), g++)
      for (const dw of [...dws].sort()) { const x = addDays(wk, dw); if (ok(x)) out.push(iso(x)); }
  } else if (r.tipo === "mensual") {
    const dom = r.dia_mes || start.getDate();
    for (let k = 0; g < 400; k += r.cada, g++) {
      const base = new Date(start.getFullYear(), start.getMonth() + k, 1);
      const x = new Date(base.getFullYear(), base.getMonth(), Math.min(dom, new Date(base.getFullYear(), base.getMonth() + 1, 0).getDate()));
      if (x > stop) break; if (ok(x)) out.push(iso(x));
    }
  } else if (r.tipo === "anual") {
    for (let k = 0; g < 50; k += r.cada, g++) { const y = start.getFullYear() + k, last = new Date(y, start.getMonth() + 1, 0).getDate(); const x = new Date(y, start.getMonth(), Math.min(start.getDate(), last)); if (x > stop) break; if (ok(x)) out.push(iso(x)); }
  }
  return out;
}
function expand(){
  const end = new Date(); end.setMonth(end.getMonth() + +$("#horizon").value);
  const floor = addDays(new Date(), -90);
  occ = [];
  for (const t of tasks) for (const f of occurrences(t, end, floor)) occ.push({ ...t, key: t.id + "@" + f, fecha: f });
  occ.sort((a, b) => (a.fecha + a.hora).localeCompare(b.fecha + b.hora));
}
$("#horizon").addEventListener("change", () => { expand(); render(); });
function repText(r){
  if (r.tipo === "ninguna") return "Una vez";
  if (r.tipo === "diaria") return r.solo_habiles ? "Días hábiles" : r.cada > 1 ? `Cada ${r.cada} días` : "Todos los días";
  if (r.tipo === "semanal") { const d = r.dias_semana.length ? " · " + r.dias_semana.map(x => DIAS_L[x]).join(", ") : ""; return (r.cada === 2 ? "Quincenal" : r.cada > 1 ? `Cada ${r.cada} semanas` : "Semanal") + d; }
  if (r.tipo === "mensual") return ({1:"Mensual",2:"Bimestral",3:"Trimestral",6:"Semestral"})[r.cada] || `Cada ${r.cada} meses`;
  return r.cada > 1 ? `Cada ${r.cada} años` : "Anual";
}
const catColor = {};
function colorOf(cat){ const k = cat || "Sin área"; if (!catColor[k]) catColor[k] = `var(--c${(Object.keys(catColor).length % 6) + 1})`; return catColor[k]; }

function finish(){
  for (const k in catColor) delete catColor[k];
  [...new Set(tasks.map(t => t.categoria || "Sin área"))].forEach(colorOf);
  expand();
  const whos = [...new Set(tasks.map(t => t.responsable).filter(Boolean))].sort();
  $("#fWho").innerHTML = `<option value="">Todos</option>` + whos.map(w => `<option>${esc(w)}</option>`).join("");
  const first = occ.find(o => o.fecha >= TODAY) || occ[0];
  if (first) { const d = fromIso(first.fecha); viewMonth = new Date(d.getFullYear(), d.getMonth(), 1); }
  sideDay = null;
  renderFilebar();
  render();
}
function renderFilebar(){
  $("#filebar").hidden = !tasks.length;
  $("#fbName").textContent = fileName || "Mi planilla";
  $("#fbSub").textContent = `${tasks.length} ${tasks.length === 1 ? "tarea" : "tareas"}`;
  $("#issues").hidden = !issues.length;
  $("#issuesSum").textContent = issues.length === 1 ? "1 fila con un dato a revisar" : `${issues.length} filas con datos a revisar`;
  $("#issuesList").innerHTML = issues.map(x => `<li>${esc(x)}</li>`).join("");
}

// ---------- Render ----------
async function toggle(key, val){
  const [tarea_id, fecha] = key.split("@");
  const o = key;
  if (val) done[o] = 1; else delete done[o];
  render();
  const q = val ? sb.from("tildes").insert({ user_id: user.id, tarea_id, fecha })
                : sb.from("tildes").delete().eq("user_id", user.id).eq("tarea_id", tarea_id).eq("fecha", fecha);
  const { error } = await q;
  if (error && !/duplicate/i.test(error.message)) {
    if (val) delete done[o]; else done[o] = 1;
    render();
    alertBar("No se pudo guardar el cambio. Revisá tu conexión.");
  }
}
function alertBar(text){
  $("#issues").hidden = false; $("#issuesSum").textContent = text; $("#issuesList").innerHTML = "";
  setTimeout(renderFilebar, 4000);
}
function render(){ renderSummary(); renderCal(); renderList(); renderRem(); }
function renderSummary(){
  const late = occ.filter(o => o.fecha < TODAY && !done[o.key]).length;
  const today = occ.filter(o => o.fecha === TODAY && !done[o.key]).length;
  const total = occ.length, d = occ.filter(o => done[o.key]).length;
  $("#sDone").textContent = d; $("#sTotal").textContent = total;
  $("#sLate").textContent = late; $("#sToday").textContent = today;
  $("#sBar").style.width = total ? (d / total * 100) + "%" : "0";
  $("#pendCount").textContent = late + today;
}
function renderCal(){
  const y = viewMonth.getFullYear(), m = viewMonth.getMonth();
  $("#monthLbl").textContent = `${MESES[m]} ${y}`;
  const first = new Date(y, m, 1), start = addDays(first, -((first.getDay() + 6) % 7));
  const byDay = {}; for (const o of occ) (byDay[o.fecha] ||= []).push(o);
  let html = ["lun","mar","mié","jue","vie","sáb","dom"].map(d => `<div class="dow">${d}</div>`).join("");
  for (let i = 0; i < 42; i++) {
    const d = addDays(start, i), k = iso(d), list = byDay[k] || [];
    if (i >= 35 && d.getMonth() !== m) break;
    const cls = ["day", d.getMonth() !== m && "out", k === TODAY && "today", k === sideDay && "sel"].filter(Boolean).join(" ");
    html += `<button class="${cls}" data-day="${k}" type="button" aria-label="${fmtDay(k)}, ${list.length} tareas">
      <span class="dnum"><span>${d.getDate()}</span>${list.length ? `<span class="c">${list.filter(o => done[o.key]).length}/${list.length}</span>` : ""}</span>
      ${list.slice(0, 3).map(o => `<span class="ev${done[o.key] ? " done" : ""}" style="--cc:${colorOf(o.categoria)}"><span>${o.hora ? o.hora + " " : ""}${esc(o.titulo)}</span></span>`).join("")}
      ${list.length > 3 ? `<span class="more">+${list.length - 3} más</span>` : ""}
    </button>`;
  }
  $("#cal").innerHTML = html;
  $("#legend").innerHTML = Object.entries(catColor).map(([k, c]) => `<span style="--cc:${c}"><i></i>${esc(k)}</span>`).join("");
}
$("#cal").addEventListener("click", e => {
  const b = e.target.closest("[data-day]"); if (!b) return;
  sideDay = sideDay === b.dataset.day ? null : b.dataset.day;
  renderCal(); renderList();
  if (window.matchMedia("(max-width:980px)").matches) $(".side").scrollIntoView({ behavior: "smooth", block: "start" });
});
$("#prev").addEventListener("click", () => { viewMonth.setMonth(viewMonth.getMonth() - 1); renderCal(); });
$("#next").addEventListener("click", () => { viewMonth.setMonth(viewMonth.getMonth() + 1); renderCal(); });
$("#todayBtn").addEventListener("click", () => { viewMonth = new Date(); viewMonth.setDate(1); sideDay = TODAY; renderCal(); renderList(); });

function itemHTML(o, showDate){
  const late = o.fecha < TODAY && !done[o.key];
  const rep = o.repetir.tipo !== "ninguna" ? repText(o.repetir) : "";
  const id = "chk-" + o.key.replace(/[^\w-]/g, "_");
  return `<div class="item${done[o.key] ? " done" : ""}">
    <input type="checkbox" id="${id}" data-key="${esc(o.key)}" ${done[o.key] ? "checked" : ""}>
    <div style="min-width:0">
      <label class="t" for="${id}">${esc(o.titulo)}</label>
      <div class="meta">
        ${o.prioridad ? `<span class="prio ${o.prioridad}">${o.prioridad}</span>` : ""}
        ${o.categoria ? `<span class="tag" style="--cc:${colorOf(o.categoria)}"><i></i>${esc(o.categoria)}</span>` : ""}
        ${o.responsable ? `<span>${esc(o.responsable)}</span>` : ""}
        ${rep ? `<span>↻ ${esc(rep)}</span>` : ""}
      </div>
      ${o.notas ? `<div class="note" style="margin-top:3px">${esc(o.notas)}</div>` : ""}
    </div>
    <div class="when${late ? " late" : ""}">${showDate ? fmtDay(o.fecha) : ""}${o.hora ? (showDate ? "<br>" : "") + o.hora : ""}</div>
  </div>`;
}
function renderList(){
  $("#dayChip").innerHTML = sideDay ? `<span class="daychip">${fmtDay(sideDay)}<button type="button" id="clearDay" aria-label="Ver toda la lista">✕</button></span>` : "";
  $("#fRange").disabled = !!sideDay;
  if (!occ.length) { $("#list").innerHTML = `<div class="empty">Todavía no hay tareas. Descargá la plantilla, completala y subila.</div>`; return; }
  const q = norm($("#q").value), who = $("#fWho").value, range = $("#fRange").value, hide = $("#fHide").checked;
  const base = occ.filter(o => (!who || o.responsable === who) && (!hide || !done[o.key])
    && (!q || norm(o.titulo + " " + o.responsable + " " + o.categoria + " " + o.notas).includes(q)));
  if (sideDay) {
    const l = base.filter(o => o.fecha === sideDay);
    $("#list").innerHTML = l.length ? `<div class="group"><h3>${l.length} ${l.length === 1 ? "tarea" : "tareas"} este día</h3>${l.map(o => itemHTML(o, false)).join("")}</div>`
      : `<div class="empty">Nada programado este día.</div>`;
    return;
  }
  const lim = range === "all" ? "9999-12-31" : iso(addDays(new Date(), +range));
  const weekEnd = iso(addDays(new Date(), 7));
  const items = base.filter(o => o.fecha <= lim);
  const groups = [
    ["Vencidas", "late", items.filter(o => o.fecha < TODAY && !done[o.key])],
    ["Hoy", "", items.filter(o => o.fecha === TODAY)],
    ["Próximos 7 días", "", items.filter(o => o.fecha > TODAY && o.fecha <= weekEnd)],
    ["Más adelante", "", items.filter(o => o.fecha > weekEnd)],
    ["Hechas", "", items.filter(o => o.fecha < TODAY && done[o.key])]
  ].filter(g => g[2].length);
  $("#list").innerHTML = groups.length
    ? groups.map(([t, cls, l]) => { const c = collapsed.has(t); return `<div class="group${c ? " closed" : ""}"><h3 class="${cls}"><button class="ghead" type="button" data-group="${t}" aria-expanded="${!c}"><svg class="chev" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6 9l6 6 6-6"/></svg>${t} <span class="count">${l.length}</span></button></h3>${c ? "" : l.map(o => itemHTML(o, true)).join("")}</div>`; }).join("")
    : `<div class="empty">No hay tareas con estos filtros.</div>`;
}
$("#list").addEventListener("click", e => {
  const b = e.target.closest("[data-group]"); if (!b) return;
  const g = b.dataset.group;
  collapsed.has(g) ? collapsed.delete(g) : collapsed.add(g);
  local.set("ap.collapsed", [...collapsed]);
  renderList();
  const again = document.querySelector(`[data-group="${g}"]`); if (again) again.focus();
});
$("#dayChip").addEventListener("click", e => { if (e.target.closest("#clearDay")) { sideDay = null; renderCal(); renderList(); } });
["#q", "#fWho", "#fRange", "#fHide"].forEach(s => $(s).addEventListener("input", renderList));
document.addEventListener("change", e => { if (e.target.matches("input[data-key]")) toggle(e.target.dataset.key, e.target.checked); });

// ---------- Pestañas ----------
function setTab(which){
  for (const k of ["agenda", "rem"]) {
    $("#tab-" + k).setAttribute("aria-selected", k === which);
    $("#view-" + k).hidden = k !== which;
  }
  local.set("ap.tab", which);
}
for (const k of ["agenda", "rem"]) $("#tab-" + k).addEventListener("click", () => setTab(k));
$("#linkPill").addEventListener("click", () => setTab("rem"));

// ---------- Recordatorios ----------
function remMsg(text, cls = ""){ $("#remMsg").innerHTML = text ? `<div class="alert ${cls}">${text}</div>` : ""; }
async function saveAvisos(){
  const patch = { aviso_dia_anterior: $("#rDay").checked, aviso_mismo_dia: $("#rSame").checked, actualizado: new Date().toISOString() };
  const { error } = await sb.from("perfiles").update(patch).eq("user_id", user.id);
  if (error) { remMsg("No se pudo guardar. Probá de nuevo.", "err"); return; }
  Object.assign(perfil, patch);
  remMsg(patch.aviso_dia_anterior || patch.aviso_mismo_dia ? "Guardado. El próximo resumen sale a las 8:00." : "Guardado. No vas a recibir avisos por correo.", "ok");
}
$("#rDay").addEventListener("change", saveAvisos);
$("#rSame").addEventListener("change", saveAvisos);

function firstDate(t){ const end = new Date(); end.setFullYear(end.getFullYear() + 3); return occurrences(t, end, fromIso(t.inicio))[0] || t.inicio; }
const BYDAY = ["SU","MO","TU","WE","TH","FR","SA"];
function rrule(t, startIso){
  const r = t.repetir; if (r.tipo === "ninguna") return null;
  const parts = [];
  if (r.tipo === "diaria" && r.solo_habiles) parts.push("FREQ=WEEKLY", "BYDAY=MO,TU,WE,TH,FR");
  else if (r.tipo === "diaria") parts.push("FREQ=DAILY", `INTERVAL=${r.cada}`);
  else if (r.tipo === "semanal") { const d = r.dias_semana.length ? r.dias_semana : [fromIso(startIso).getDay()]; parts.push("FREQ=WEEKLY", `INTERVAL=${r.cada}`, "BYDAY=" + [...d].sort().map(x => BYDAY[x]).join(",")); }
  else if (r.tipo === "mensual") {
    const dom = r.dia_mes || fromIso(t.inicio).getDate();
    parts.push("FREQ=MONTHLY", `INTERVAL=${r.cada}`);
    if (dom >= 31) parts.push("BYMONTHDAY=-1");
    else if (dom >= 29) parts.push("BYMONTHDAY=" + Array.from({ length: dom - 27 }, (_, i) => 28 + i).join(","), "BYSETPOS=-1");
    else parts.push("BYMONTHDAY=" + dom);
  }
  else if (r.tipo === "anual") parts.push("FREQ=YEARLY", `INTERVAL=${r.cada}`);
  if (r.hasta) parts.push("UNTIL=" + r.hasta.replace(/-/g, "") + "T235959Z");
  return "RRULE:" + parts.join(";");
}
function describe(t){
  return [t.notas, t.responsable && "Responsable: " + t.responsable, t.categoria && "Área: " + t.categoria,
    t.prioridad && "Prioridad: " + t.prioridad, "Creado desde Agenda Planilla"].filter(Boolean).join("\n");
}
function slot(t){
  const day = firstDate(t).replace(/-/g, ""), hora = (t.hora || $("#defHour").value || "09:00").replace(":", "");
  const endM = +hora.slice(0, 2) * 60 + +hora.slice(2) + 30;
  return { day, start: `${day}T${hora}00`, end: `${day}T${endM >= 1440 ? "2359" : pad(Math.floor(endM / 60)) + pad(endM % 60)}00` };
}
function gcalLink(t){
  const s = slot(t), rr = rrule(t, firstDate(t));
  const p = new URLSearchParams({ action: "TEMPLATE", text: t.titulo, dates: `${s.start}/${s.end}`, details: describe(t), ctz: TZ });
  if (rr) p.set("recur", rr);
  return "https://calendar.google.com/calendar/render?" + p.toString();
}
function renderRem(){
  if (!tasks.length) { $("#remList").innerHTML = `<div class="empty">Cuando subas una planilla, acá vas a poder pasar cada tarea a tu Google Calendar.</div>`; return; }
  const sorted = [...tasks].sort((a, b) => firstDate(a).localeCompare(firstDate(b)));
  $("#remList").innerHTML = sorted.map(t => `<div class="rrow" style="grid-template-columns:1fr auto">
      <div style="min-width:0"><div class="t">${esc(t.titulo)}</div>
        <div class="meta"><span>${esc(repText(t.repetir))}</span><span>Desde ${fmtDay(firstDate(t))}${t.hora ? " · " + t.hora : ""}</span>${t.responsable ? `<span>${esc(t.responsable)}</span>` : ""}</div></div>
      <a class="btn quiet sm gcal" href="${esc(gcalLink(t))}" target="_blank" rel="noopener">Agregar a Google Calendar ↗</a>
    </div>`).join("");
}
$("#defHour").addEventListener("change", renderRem);

function icsEsc(s){ return String(s).replace(/\\/g, "\\\\").replace(/\n/g, "\\n").replace(/[,;]/g, m => "\\" + m); }
function fold(line){ const out = []; while (line.length > 74) { out.push(line.slice(0, 74)); line = " " + line.slice(74); } out.push(line); return out.join("\r\n"); }
function buildIcs(){
  const stamp = new Date().toISOString().replace(/[-:]/g, "").replace(/\.\d+/, "");
  const L = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Agenda Planilla//ES", "CALSCALE:GREGORIAN",
    "BEGIN:VTIMEZONE", "TZID:" + TZ, "BEGIN:STANDARD", "DTSTART:19700101T000000", "TZOFFSETFROM:-0300", "TZOFFSETTO:-0300", "END:STANDARD", "END:VTIMEZONE"];
  for (const t of tasks) {
    const s = slot(t);
    L.push("BEGIN:VEVENT", `UID:${t.id}-${user.id}@agenda-planilla`, `DTSTAMP:${stamp}`,
      `DTSTART;TZID=${TZ}:${s.start}`, `DTEND;TZID=${TZ}:${s.end}`,
      fold("SUMMARY:" + icsEsc(t.titulo)), fold("DESCRIPTION:" + icsEsc(describe(t))));
    const rr = rrule(t, firstDate(t)); if (rr) L.push(rr);
    L.push("BEGIN:VALARM", "ACTION:DISPLAY", fold("DESCRIPTION:" + icsEsc(t.titulo)), "TRIGGER:-PT15M", "END:VALARM");
    L.push("END:VEVENT");
  }
  L.push("END:VCALENDAR");
  return L.join("\r\n");
}
$("#icsBtn").addEventListener("click", () => {
  if (!tasks.length) { remMsg("Primero subí una planilla.", "warn"); return; }
  const url = URL.createObjectURL(new Blob([buildIcs()], { type: "text/calendar" }));
  const a = document.createElement("a"); a.href = url; a.download = "agenda-planilla.ics";
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
});

// Arranque: si no hay sesión, muestra el ingreso
sb.auth.getSession().then(({ data }) => { if (!data.session) showLogin(); });
})();
