let excelData = [];
const urlParams = new URLSearchParams(window.location.search);
const carreraKey = (urlParams.get('carrera') || 'REDES').toUpperCase();

let ultimoMotivoFallo = "";

// --- 1. CONFIGURACIÓN ---
const esCarreraSemi = carreraKey.includes('_SEMI') || carreraKey.includes('SEMIPRESENCIAL');
let modalidadActiva = esCarreraSemi ? 'SEMIPRESENCIAL' : 'SEMANAL';

let STORAGE_KEY = `upds_v25_${carreraKey}_${modalidadActiva}`;
let estados = JSON.parse(localStorage.getItem(STORAGE_KEY)) || {};

const tooltip = document.createElement('div');
Object.assign(tooltip.style, {
    position: 'fixed', padding: '12px 16px', background: 'rgba(0, 26, 53, 0.95)',
    color: 'white', borderRadius: '10px', fontSize: '12px', pointerEvents: 'none',
    display: 'none', zIndex: '9999', boxShadow: '0 10px 25px rgba(0,0,0,0.4)',
    borderLeft: '4px solid #ffcc00', lineHeight: '1.5', maxWidth: '280px', backdropFilter: 'blur(5px)'
});
document.body.appendChild(tooltip);

const moduloAMes = {
    '1º MÓDULO': 'Feb', '2º MÓDULO': 'Mar', '3º MÓDULO': 'Abr',
    '4º MÓDULO': 'May', '5º MÓDULO': 'Jun', '6º MÓDULO': 'Jul'
};

const mesesOrden = ['Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Invierno', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic', 'Verano'];
const modulosCabecera = Object.keys(moduloAMes);

// --- 2. SINCRONIZACIÓN ---
async function sincronizarOferta() {
    const msg = document.getElementById('msg-clash');
    try {
        const response = await fetch('data/oferta.xlsx?v=' + Date.now());
        if (!response.ok) throw new Error("No se encuentra data/oferta.xlsx");
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), {type:'array'});
        excelData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        if(msg) { msg.innerText = `✅ Sincronizado (${excelData.length} filas)`; msg.style.color = "#2e7d32"; }
        render();
    } catch (e) {
        if(msg) { msg.innerText = "🚨 Error leyendo Excel"; msg.style.color = "red"; }
        render();
    }
}

function limpiarTexto(t) {
    if (!t) return "";
    return String(t).toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Z0-9 ]/g, " ").replace(/\s+/g, " ").trim();
}

function mapearTurno(tExcel) {
    const t = limpiarTexto(tExcel);
    if (t.includes("MANANA")) return "Mañana";
    if (t.includes("MEDIO")) return "Medio Día";
    if (t.includes("TARDE")) return "Tarde";
    if (t.includes("NOCHE")) return "Noche";
    return "Turno";
}

// --- 3. LÓGICA DE EXTRACCIÓN ---

function extraerNombreRealElectiva(rawExcel, nomMalla) {
    if (!nomMalla.toUpperCase().includes("ELECTIVA")) return rawExcel;
    return rawExcel.replace(/\(ELECTIVA\s*\d+\)/gi, "").replace(/ELECTIVA\s*\d+/gi, "").trim();
}

function obtenerOpcionesOferta(nomMalla) {
    if (excelData.length === 0) return [];
    const nMatBusqueda = limpiarTexto(nomMalla);
    const carreraBase = limpiarTexto(carreraKey).replace('_SEMI', '').replace('_NUEVA', '').trim();
    const listaPermitida = (typeof ELECTIVAS_POR_CARRERA !== 'undefined') ? (ELECTIVAS_POR_CARRERA[carreraKey] || []) : [];
    
    let opcionesMap = new Map();
    const electivasTomadas = Object.values(estados).filter(v => v.nombreReal && (v.st === 'A' || v.st === 'P')).map(v => v.nombreReal);

    excelData.forEach(r => {
        const modExcel = limpiarTexto(r['MODALIDAD'] || "");
        let coincideModalidad = (modalidadActiva === "SEMANAL") ? 
            (modExcel.includes("PRESENCIAL") || modExcel.includes("SEMANAL")) && !modExcel.includes("SEMI") : 
            (modExcel.includes("SEMIPRESENCIAL") || modExcel.includes("SEMI"));
        
        if (!coincideModalidad) return;

        const turnoActual = mapearTurno(r['TURNO']);
        const carreraActual = r['CARRERA'] || "OTRA";

        modulosCabecera.forEach(col => {
            const vExcelRaw = r[col] || "";
            const vExcelLimpio = limpiarTexto(vExcelRaw);
            if (vExcelLimpio === nMatBusqueda || (nomMalla.toUpperCase().includes("ELECTIVA") && vExcelRaw.toUpperCase().includes(nomMalla.toUpperCase()))) {
                const nombreRealMateria = extraerNombreRealElectiva(vExcelRaw, nomMalla);
                let esValidaParaCarrera = true;
                if (nomMalla.toUpperCase().includes("ELECTIVA") && listaPermitida.length > 0) {
                    esValidaParaCarrera = listaPermitida.some(p => limpiarTexto(nombreRealMateria) === limpiarTexto(p));
                }

                if (esValidaParaCarrera) {
                    const mesActual = moduloAMes[col];
                    const yaCursada = electivasTomadas.includes(nombreRealMateria);
                    const llaveUnica = `${mesActual}-${turnoActual}-${carreraActual}-${nombreRealMateria}`;
                    if (!opcionesMap.has(llaveUnica)) {
                        opcionesMap.set(llaveUnica, { mes: mesActual, turno: turnoActual, carrera: carreraActual, nombreReal: nombreRealMateria, yaCursada: yaCursada, esPropia: limpiarTexto(carreraActual).includes(carreraBase) });
                    }
                }
            }
        });
    });
    return Array.from(opcionesMap.values()).sort((a, b) => mesesOrden.indexOf(a.mes) - mesesOrden.indexOf(b.mes));
}

// --- 4. REGLAS DE NEGOCIO ---

function contarAprobadas() {
    return Object.values(estados).filter(v => v.st === 'A').length;
}

function contarRegulares() {
    const mesesReg = ['Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
    return Object.values(estados).filter(v => v.st === 'P' && mesesReg.includes(v.mes)).length;
}

function contarTemporada(temp) {
    return Object.values(estados).filter(v => v.st === 'P' && v.mes === temp).length;
}

function verificarChoque(id, mes, turno) {
    for(const [k,v] of Object.entries(estados)) {
        if(k!==id && v.st==='P' && v.mes===mes && v.turno===turno) return true;
    }
    return false;
}

function validarLimiteSemestre(idMateria) {
    const mat = (typeof materias !== 'undefined') ? materias.find(m => m.id === idMateria) : null;
    if (!mat) return { ok: true };
    const aprobadas = contarAprobadas();
    const semestreBase = Math.floor(aprobadas / 6) + 1;
    const limite = semestreBase + 3;
    if (mat.sem > limite) return { ok: false, msg: `NIVELACIÓN: Máximo Sem. ${limite} (Estás en el ${semestreBase}).` };
    return { ok: true };
}

function validarHabilitacionElectiva(materiaId, mesElegido) {
    if (typeof materias === 'undefined') return { ok: true };
    const mat = materias.find(m => m.id === materiaId);
    if (!mat || mat.tipo !== 'electiva') return { ok: true };
    const materiasBase = materias.filter(m => m.sem <= 5);
    const idxNuevo = mesesOrden.indexOf(mesElegido);
    for (let mBase of materiasBase) {
        const est = estados[mBase.id];
        if (!est) return { ok: false, msg: `REQUISITO: Falta completar S1-S5.` };
        if (est.st === 'P') {
            const idxBase = mesesOrden.indexOf(est.mes);
            if (idxBase >= idxNuevo) return { ok: false, msg: `CRONOLOGÍA: ${mBase.nom} antes que electiva.` };
        }
    }
    return { ok: true };
}

function validarCronologiaPre(mId, mesElegido) {
    if (typeof materias === 'undefined') return { ok: true };
    const mat = materias.find(m => m.id === mId);
    if (!mat || mat.pre.length === 0) return { ok: true };
    const idxNuevo = mesesOrden.indexOf(mesElegido);
    for (let pId of mat.pre) {
        const estPre = estados[pId];
        if (!estPre) return { ok: false, msg: "Prerrequisito no definido." };
        if (estPre.st === 'A') continue; 
        if (estPre.st === 'P') {
            const idxPre = mesesOrden.indexOf(estPre.mes);
            if (idxPre >= idxNuevo) return { ok: false, msg: `PRERREQUISITO: Proyectado en ${estPre.mes}.` };
        }
    }
    return { ok: true };
}

// --- 5. RESALTADO ---
function highlightRelaciones(materiaId) {
    if (typeof materias === 'undefined') return;
    const mat = materias.find(m => m.id === materiaId);
    if (!mat) return;
    document.querySelectorAll('.materia').forEach(div => div.style.opacity = '0.2');
    const targetDiv = document.getElementById(materiaId);
    if (targetDiv) targetDiv.style.opacity = '1';
    mat.pre.forEach(preId => {
        const preDiv = document.getElementById(preId);
        if (preDiv) { preDiv.style.opacity = '1'; preDiv.style.border = '2px solid #f43f5e'; preDiv.style.boxShadow = '0 0 15px rgba(244, 63, 94, 0.6)'; }
    });
    materias.filter(m => m.pre.includes(materiaId)).forEach(post => {
        const postDiv = document.getElementById(post.id);
        if (postDiv) { postDiv.style.opacity = '1'; postDiv.style.border = '2px solid #06b6d4'; postDiv.style.boxShadow = '0 0 15px rgba(6, 182, 212, 0.6)'; }
    });
}
function removeHighlight() { document.querySelectorAll('.materia').forEach(div => { div.style.opacity = '1'; div.style.border = ''; div.style.boxShadow = ''; }); }

// --- 6. RENDERIZADO Y SELECCIÓN ---

function seleccionarMateria(id, mes, turno, externa, carreraOri, nombreReal = null) {
    const esMesRegular = !['Invierno', 'Verano'].includes(mes);
    if (esMesRegular && contarRegulares() >= 6) { alert("⛔ LÍMITE: Máximo 6 regulares."); return; }
    if (mes === 'Invierno' && contarTemporada('Verano') >= 1) { alert("⛔ GESTIÓN: Ya tienes Verano."); return; }
    if (mes === 'Verano' && contarTemporada('Invierno') >= 1) { alert("⛔ GESTIÓN: Ya tienes Invierno."); return; }
    if (!esMesRegular && contarTemporada(mes) >= 1) { alert(`⛔ LÍMITE: Solo 1 en ${mes}.`); return; }

    const adelanto = validarLimiteSemestre(id);
    if (!adelanto.ok) { alert(adelanto.msg); return; }

    if (!validarCronologiaPre(id, mes).ok) { alert("⛔ BLOQUEO: " + validarCronologiaPre(id, mes).msg); return; }
    
    const checkElec = validarHabilitacionElectiva(id, mes);
    if (!checkElec.ok) { alert(checkElec.msg); return; }
    
    const header = document.getElementById('materia-activa-display');
    if(header) header.innerText = "";

    estados[id] = { st: 'P', mes: mes, turno: turno, externa: externa, carreraOri: carreraOri, nombreReal: nombreReal };
    save(); render();
}

function render() {
    const body = document.getElementById('malla-body');
    const headerDisplay = document.getElementById('materia-activa-display');
    if (!body || typeof materias === 'undefined') return;
    body.innerHTML = '';
    const materiaList = materias;

    const s5Cerrado = materiaList.filter(m => m.sem <= 5).every(x => estados[x.id] !== undefined);
    const tieneInvierno = Object.values(estados).some(v => v.mes === 'Invierno');
    const tieneVerano = Object.values(estados).some(v => v.mes === 'Verano');

    for (let s = 1; s <= Math.max(...materiaList.map(m => m.sem)); s++) {
        const col = document.createElement('div');
        col.className = 'semestre-col';
        col.innerHTML = `<div class="sem-header">Sem. ${s}<button class="btn-mass-approve" onclick="aprobarSemestre(${s})">✓</button></div>`;

        materiaList.filter(m => m.sem === s).forEach(m => {
            const div = document.createElement('div');
            const d = estados[m.id];
            
            let habilitadaParaClick = m.pre.every(pId => estados[pId] !== undefined);
            if (!validarLimiteSemestre(m.id).ok) habilitadaParaClick = false;
            if (m.tipo === 'electiva' && !s5Cerrado) habilitadaParaClick = false;

            let claseEstado = "bloqueada";
            if (d?.st === 'A') claseEstado = "aprobada";
            else if (d?.st === 'P') claseEstado = "programada";
            else if (habilitadaParaClick) claseEstado = ""; 

            div.id = m.id;
            div.className = `materia ${claseEstado}`;
            
            div.onmouseenter = () => { highlightRelaciones(m.id); const infoPre = m.pre.map(pId => materiaList.find(x => x.id === pId)?.nom || pId); tooltip.innerHTML = `<strong>${m.nom}</strong><br><small>PRE-REQUISITOS:</small><br>• ${infoPre.length > 0 ? infoPre.join("<br>• ") : "Ninguno"}`; tooltip.style.display = 'block'; };
            div.onmousemove = (e) => { tooltip.style.left = (e.clientX + 15) + 'px'; tooltip.style.top = (e.clientY + 15) + 'px'; };
            div.onmouseleave = () => { removeHighlight(); tooltip.style.display = 'none'; };

            let nombreDisplay = m.nom;
            if (d?.st === 'P' && d.nombreReal) nombreDisplay = `${m.nom}:<br><small style="color:#fbbf24">${d.nombreReal}</small>`;

            div.innerHTML = `<strong class="mat-title">${nombreDisplay}</strong>` + 
                (d?.st === 'A' ? '<div class="status-label">APROBADA</div>' : 
                (d?.st === 'P' ? `<div class="info-programacion"><div class="prog-display">${d.mes} - ${d.turno}</div></div>` : 
                (habilitadaParaClick ? '<span class="status-ok">✅ DISPONIBLE</span>' : '<span class="status-lock">🔒 BLOQUEADA</span>')));

            div.onclick = (e) => {
                if(estados[m.id]?.st === 'P') { 
                    delete estados[m.id]; 
                    if(headerDisplay) headerDisplay.innerText = "";
                    save(); render(); return; 
                }
                if(!estados[m.id]) {
                    if (habilitadaParaClick) { estados[m.id] = { st: 'A' }; save(); render(); }
                    else alert(m.tipo === 'electiva' && !s5Cerrado ? "⛔ Debes completar S1-S5." : "⛔ Requisitos pendientes.");
                    return;
                }
                if(estados[m.id]?.st === 'A') {
                    if(headerDisplay) headerDisplay.innerText = `PROGRAMANDO: ${m.nom} - ${carreraKey}`;

                    let opciones = obtenerOpcionesOferta(m.nom);
                    const turnosTemporada = ['Mañana', 'Medio Día', 'Tarde', 'Noche'];
                    
                    if (!tieneVerano && !opciones.some(o => o.mes === 'Invierno')) {
                        turnosTemporada.forEach(t => opciones.push({ mes: 'Invierno', turno: t, carrera: 'TEMPORADA', esPropia: true, nombreReal: "TEMPORADA" }));
                    }
                    if (!tieneInvierno && !opciones.some(o => o.mes === 'Verano')) {
                        turnosTemporada.forEach(t => opciones.push({ mes: 'Verano', turno: t, carrera: 'TEMPORADA', esPropia: true, nombreReal: "TEMPORADA" }));
                    }

                    const menu = document.createElement('div');
                    menu.className = 'menu-seleccion-oferta';
                    const menuWidth = 240;
                    const menuHeightEstimate = Math.min(450, opciones.length * 60 + 50); 
                    let posX = e.pageX; let posY = e.pageY;
                    if (posX + menuWidth > window.innerWidth) posX = posX - menuWidth;
                    if (posY + menuHeightEstimate > window.innerHeight + window.scrollY) posY = posY - menuHeightEstimate;
                    if (posY < window.scrollY) posY = window.scrollY + 10;

                    Object.assign(menu.style, { position: 'absolute', top: posY + 'px', left: posX + 'px', background: '#001a35', border: '1px solid #ffcc00', borderRadius: '8px', padding: '12px', zIndex: '10000', color: 'white', minWidth: menuWidth + 'px', boxShadow: '0 10px 30px rgba(0,0,0,0.5)' });
                    
                    let h = `<div style="font-weight:bold; color:#ffcc00; font-size:11px; margin-bottom:10px;">SELECCIONAR HORARIO:</div>`;
                    
                    // --- REPARACIÓN: QUITAR APROBADA ---
                    h += `<div style="padding:8px; cursor:pointer; color:#f43f5e; border-bottom:1px solid #334155; font-size:10px;" 
                               onclick="event.stopPropagation(); delete estados['${m.id}']; if(document.getElementById('materia-activa-display')) document.getElementById('materia-activa-display').innerText=''; save(); render();">✖ QUITAR APROBADA</div>`;

                    opciones.forEach(opt => {
                        const crono = validarCronologiaPre(m.id, opt.mes);
                        const hElec = validarHabilitacionElectiva(m.id, opt.mes);
                        const esMesReg = !['Invierno', 'Verano'].includes(opt.mes);
                        const gBloq = (opt.mes === 'Invierno' && tieneVerano) || (opt.mes === 'Verano' && tieneInvierno);
                        const disabled = !crono.ok || !hElec.ok || opt.yaCursada || verificarChoque(m.id, opt.mes, opt.turno) || (esMesReg && contarRegulares() >= 6) || (!esMesReg && contarTemporada(opt.mes) >= 1) || gBloq;

                        h += `<div style="padding:8px; cursor:${disabled ? 'not-allowed' : 'pointer'}; border-bottom:1px solid #334155; opacity:${disabled ? 0.4 : 1}"
                                 onclick="${disabled ? '' : `event.stopPropagation(); seleccionarMateria('${m.id}','${opt.mes}','${opt.turno}',${!opt.esPropia},'${opt.carrera}','${opt.nombreReal}'); document.querySelector('.menu-seleccion-oferta').remove();`}">
                                <span style="color:${esMesReg ? '#ffcc00' : '#38bdf8'}; font-weight:bold;">${opt.mes}</span> - ${opt.turno}
                                <div style="font-size:10px; color:#cbd5e1;">${opt.nombreReal || ""}</div>
                                ${opt.yaCursada ? `<small style="color:#f43f5e">[YA CURSADA]</small>` : ''}
                                ${!hElec.ok ? `<br><small style="color:#f43f5e">${hElec.msg}</small>` : ''}
                                ${!crono.ok ? `<br><small style="color:#f43f5e">${crono.msg}</small>` : ''}
                                <div style="font-size:9px; color:#94a3b8; font-style:italic; margin-top:2px;">Carrera: ${opt.carrera}</div>
                            </div>`;
                    });
                    menu.innerHTML = h; document.body.appendChild(menu);
                    
                    setTimeout(() => { 
                        window.onclick = () => { 
                            if(document.querySelector('.menu-seleccion-oferta')) {
                                document.querySelector('.menu-seleccion-oferta').remove();
                                if(headerDisplay) headerDisplay.innerText = "";
                            }
                            window.onclick = null; 
                        }; 
                    }, 100);
                }
            };
            col.appendChild(div);
        });
        body.appendChild(col);
    }
}

// --- 7. UTILIDADES ---
function aprobarSemestre(n) { 
    if (typeof materias === 'undefined') return;
    const materiasSem = materias.filter(m => m.sem === n);
    const yaHayAprobadas = materiasSem.some(m => estados[m.id]?.st === 'A');
    if (yaHayAprobadas) materiasSem.forEach(m => { delete estados[m.id]; });
    else materiasSem.forEach(m => { estados[m.id] = { st: 'A' }; });
    save(); render(); 
}
function cambiarModalidad(n) { modalidadActiva=n; STORAGE_KEY=`upds_v25_${carreraKey}_${modalidadActiva}`; estados=JSON.parse(localStorage.getItem(STORAGE_KEY))||{}; render(); }
function save() { localStorage.setItem(STORAGE_KEY, JSON.stringify(estados)); }
function reiniciarMalla() { if(confirm("¿Limpiar todo?")) { estados={}; save(); render(); } }
function enviarWhatsApp() {
    if (typeof materias === 'undefined') return;
    const t = document.getElementById('student-phone')?.value;
    if (!t) { alert("Ingresa un número."); return; }
    let msg = `*UPDS PROYECCIÓN*\n*Carrera:* ${carreraKey}\n\n`;
    Object.keys(estados).filter(id => estados[id].st === 'P').forEach((id, i) => {
        const m = materias.find(x => x.id === id);
        if(m) {
            const nomFinal = estados[id].nombreReal ? `${m.nom} (${estados[id].nombreReal})` : m.nom;
            msg += `${i+1}. *${nomFinal}* (${estados[id].mes} - ${estados[id].turno})\n`;
        }
    });
    window.open(`https://api.whatsapp.com/send?phone=591${t}&text=${encodeURIComponent(msg)}`, '_blank');
}

window.onload = () => { 
    const tituloCarrera = document.getElementById('carrera-actual-titulo');
    if (tituloCarrera) {
        tituloCarrera.innerText = carreraKey.replace('_', ' '); 
    }
    sincronizarOferta(); 
};