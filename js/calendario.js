// --- CONFIGURACIÓN E INSTANCIAS ---
const CALENDAR_STORAGE_KEY = 'upds_calendar_v25_final';
let currentDate = new Date();
let eventsData = JSON.parse(localStorage.getItem(CALENDAR_STORAGE_KEY)) || [];
let rawExcelEvents = []; 

// 1. INYECTOR DE INTERFAZ (DISEÑO PREMIUM)
function injectCalendarUI() {
    if(document.getElementById('calendar-widget')) return;

    const btn = document.createElement('button');
    btn.id = 'btn-calendar-toggle';
    btn.innerHTML = '<span class="cal-icon">📅</span>';
    btn.onclick = toggleCalendar;
    document.body.appendChild(btn);

    const widget = document.createElement('div');
    widget.className = 'calendar-widget-v3'; 
    widget.id = 'calendar-widget';
    widget.innerHTML = `
        <div class="cal-app-container">
            <div class="cal-main-panel">
                <div class="cal-header-v3">
                    <button class="nav-btn" onclick="changeMonth(-1)">❮</button>
                    <div class="month-display">
                        <h3 id="cal-month-year"></h3>
                        <span class="btn-today" onclick="goToToday()">HOY</span>
                    </div>
                    <button class="nav-btn" onclick="changeMonth(1)">❯</button>
                </div>
                <div class="cal-grid-v3" id="cal-grid"></div>
                <div class="cal-footer-actions">
                    <button onclick="descargarPlantillaExcel()" class="f-action">Descargar Plantilla</button>
                    <button onclick="borrarTodoElCalendario()" class="f-action danger">Limpiar Notas</button>
                </div>
            </div>
            <div class="cal-side-panel">
                <div id="cal-details-header" class="side-header">
                    <span class="selected-date-text" id="side-date-display">Selecciona un día</span>
                    <button class="add-note-btn" onclick="showAddForm()">+</button>
                </div>
                <div id="cal-events-list" class="side-events-list">
                    <div class="empty-state-msg">Sin actividades</div>
                </div>
                <div id="cal-editor" class="mini-editor" style="display:none;">
                    <input type="text" id="edit-title" placeholder="¿Qué tienes planeado?">
                    <div class="editor-actions">
                        <button onclick="guardarCambios()" class="save-btn">Guardar</button>
                        <button onclick="cerrarEditor()" class="cancel-btn">X</button>
                    </div>
                </div>
            </div>
        </div>
    `;
    document.body.appendChild(widget);
}

// 2. MOTOR DE SINCRONIZACIÓN
async function sincronizarArchivoData() {
    if (typeof XLSX === 'undefined') return; 

    try {
        const response = await fetch('data/calendario_upds.xlsx?v=' + Date.now());
        if (!response.ok) return;
        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: "" });

        rawExcelEvents = rows.map(row => {
            let fRaw = ""; let evRaw = ""; let descRaw = "";
            Object.keys(row).forEach(k => {
                const key = k.trim().toUpperCase();
                if(key === "FECHA") fRaw = row[k];
                if(key === "EVENTO") evRaw = row[k];
                if(key === "DESCRIPCION" || key === "DETALLE") descRaw = row[k];
            });

            if (fRaw && evRaw) {
                let fFinal = "";
                if (typeof fRaw === 'string' && fRaw.includes("/")) {
                    const p = fRaw.split("/");
                    if(p.length === 3) fFinal = `${p[2]}-${p[1].padStart(2,'0')}-${p[0].padStart(2,'0')}`;
                } else if (!isNaN(fRaw)) {
                    const d = new Date((fRaw - 25569) * 86400 * 1000);
                    if(!isNaN(d)) fFinal = d.toISOString().split('T')[0];
                }
                return { 
                    fecha: fFinal, 
                    evento: String(evRaw).trim(), 
                    descripcion: descRaw || "Evento programado en el calendario académico." 
                };
            }
            return null;
        }).filter(x => x !== null);
        
        renderCalendar();
        // --- ACTIVACIÓN DEL BRIEFING DIARIO ---
        revisarEventosDelDia(); 
    } catch (e) { console.warn("Sync falló"); }
}

// 3. BUSCADOR DE EVENTOS
function obtenerEventosDelDia(dateStr) {
    let encontrados = [];
    rawExcelEvents.forEach(ev => {
        if (ev.fecha === dateStr) {
            encontrados.push({ 
                id: 'ex-'+Math.random(), 
                title: ev.evento, 
                desc: ev.descripcion,
                type: 'academic', 
                isAuto: true 
            });
        }
    });
    const personales = eventsData.filter(e => e.date === dateStr);
    return encontrados.concat(personales);
}

// 4. RENDERIZADO
function renderCalendar() {
    const year = currentDate.getFullYear();
    const month = currentDate.getMonth();
    const monthName = new Date(year, month).toLocaleString('es-ES', { month: 'long', year: 'numeric' }).toUpperCase();
    document.getElementById('cal-month-year').innerText = monthName;

    const grid = document.getElementById('cal-grid');
    grid.innerHTML = '';
    ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'].forEach(d => grid.innerHTML += `<div class="day-name-label">${d}</div>`);

    const firstDay = new Date(year, month, 1).getDay();
    for(let i=0; i<firstDay; i++) grid.innerHTML += `<div class="day-cell empty"></div>`;

    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const todayStr = new Date().toISOString().split('T')[0];

    for(let i=1; i<=daysInMonth; i++) {
        const loopDate = `${year}-${String(month+1).padStart(2,'0')}-${String(i).padStart(2,'0')}`;
        const evs = obtenerEventosDelDia(loopDate);
        const dayDiv = document.createElement('div');
        dayDiv.className = `day-cell ${loopDate === todayStr ? 'is-today' : ''} ${evs.length > 0 ? 'has-content' : ''}`;
        
        let indicators = "";
        if(evs.some(e => e.isAuto)) indicators += `<span class="dot-acad"></span>`;
        if(evs.some(e => !e.isAuto)) indicators += `<span class="dot-pers"></span>`;

        dayDiv.innerHTML = `<span class="day-number">${i}</span><div class="dots-box">${indicators}</div>`;
        dayDiv.onclick = () => {
            document.querySelectorAll('.day-cell').forEach(d => d.classList.remove('is-selected'));
            dayDiv.classList.add('is-selected');
            showEventsForDate(loopDate);
        };
        grid.appendChild(dayDiv);
    }
}

function showEventsForDate(dateStr) {
    const list = document.getElementById('cal-events-list');
    const displayDate = document.getElementById('side-date-display');
    const evs = obtenerEventosDelDia(dateStr);
    list.setAttribute('data-selected', dateStr);
    
    const dateObj = new Date(dateStr + "T00:00:00");
    const datePretty = dateObj.toLocaleDateString('es-ES', { day: 'numeric', month: 'short' });
    displayDate.innerText = datePretty;
    
    if(evs.length === 0) list.innerHTML = `<div class="empty-state-msg">Sin actividades</div>`;
    else {
        list.innerHTML = evs.map((e) => `
            <div class="event-item-v3 ${e.isAuto ? 'institutional' : 'manual'} ${e.completed ? 'is-done' : ''}" 
                 onclick="lanzarModalEvento('${e.title}', '${datePretty}', '${e.desc || 'Nota personal.'}', '${e.isAuto ? 'ACADÉMICO' : 'PERSONAL'}')">
                <div class="ev-check" onclick="event.stopPropagation(); toggleEstado('${e.id}', '${dateStr}')">
                    ${e.completed ? '●' : '○'}
                </div>
                <div class="ev-content">
                    <span class="ev-title-text">${e.title}</span>
                </div>
                ${!e.isAuto ? `<button class="ev-delete-btn" onclick="event.stopPropagation(); eliminarEvento('${e.id}')">✕</button>` : ''}
            </div>
        `).join('');
    }
}

// --- LÓGICA DE ALERTA EMERGENTE (MODAL) ---
function lanzarModalEvento(titulo, fecha, descripcion, categoria) {
    const modal = document.getElementById('event-modal');
    if(!modal) return;

    document.getElementById('modal-event-title').innerText = titulo;
    document.getElementById('modal-event-date').innerText = fecha;
    document.getElementById('modal-event-description').innerText = descripcion;
    document.getElementById('modal-event-category').innerText = categoria;

    const badge = document.getElementById('modal-event-category');
    if(categoria === 'PERSONAL') {
        badge.style.background = '#06b6d4';
        badge.style.color = 'white';
    } else {
        badge.style.background = '#ffcc00';
        badge.style.color = '#001a35';
    }

    modal.style.display = 'flex';
}

function cerrarAlertaEvento() {
    const modal = document.getElementById('event-modal');
    if(modal) modal.style.display = 'none';
}

// NUEVO: Vigilante que detecta eventos al iniciar el día
function revisarEventosDelDia() {
    // Obtenemos la fecha de hoy en formato local YYYY-MM-DD
    const t = new Date();
    const todayStr = `${t.getFullYear()}-${String(t.getMonth()+1).padStart(2,'0')}-${String(t.getDate()).padStart(2,'0')}`;
    const evs = obtenerEventosDelDia(todayStr);

    if (evs.length > 0) {
        const principal = evs[0];
        const datePretty = t.toLocaleDateString('es-ES', { day: 'numeric', month: 'long' });
        
        // Esperamos 1.5 segundos para que la carga sea fluida
        setTimeout(() => {
            lanzarModalEvento(
                "📢 ACTIVIDAD PARA HOY", 
                datePretty, 
                `Tienes ${evs.length} actividad(es) registrada(s) para hoy: \n\n${evs.map(e => "• " + e.title).join('\n')}`,
                principal.isAuto ? 'ACADÉMICO' : 'PERSONAL'
            );
        }, 1500); 
    }
}

// 5. GESTIÓN Y PERSISTENCIA
function toggleEstado(id, date) {
    const ev = eventsData.find(e => e.id === id);
    if(ev) { ev.completed = !ev.completed; actualizarTodo(date); }
}

function eliminarEvento(id) {
    if(!confirm("¿Borrar nota?")) return;
    const date = eventsData.find(e => e.id === id)?.date;
    eventsData = eventsData.filter(e => e.id !== id);
    actualizarTodo(date);
}

function guardarCambios() {
    const title = document.getElementById('edit-title').value.trim();
    const date = document.getElementById('cal-events-list').getAttribute('data-selected');
    if(!title || !date) return;
    eventsData.push({ 
        id: 'm-'+Date.now(), 
        date: date, 
        title: title, 
        desc: "Nota personal creada por el usuario.",
        completed: false 
    });
    actualizarTodo(date);
}

function descargarPlantillaExcel() {
    if (typeof XLSX === 'undefined') return alert("Librería no cargada");
    const data = [{ "FECHA": "24/03/2026", "EVENTO": "Examen Final", "DESCRIPCION": "Llevar carnet" }];
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Fechas");
    XLSX.writeFile(wb, "Plantilla_UPDS.xlsx");
}

function actualizarTodo(d) {
    localStorage.setItem(CALENDAR_STORAGE_KEY, JSON.stringify(eventsData));
    renderCalendar();
    if(d) showEventsForDate(d);
    cerrarEditor();
}

function showAddForm() { document.getElementById('cal-editor').style.display = 'block'; }
function cerrarEditor() { document.getElementById('cal-editor').style.display = 'none'; document.getElementById('edit-title').value = ""; }
function goToToday() { currentDate = new Date(); renderCalendar(); }
function toggleCalendar() { document.getElementById('calendar-widget').classList.toggle('active'); }
function changeMonth(delta) { currentDate.setMonth(currentDate.getMonth() + delta); renderCalendar(); }
function borrarTodoElCalendario() { if(confirm("¿Borrar notas?")) { eventsData = []; actualizarTodo(); } }

// INICIALIZACIÓN SEGURA
window.addEventListener('load', () => {
    injectCalendarUI();
    
    const btnCerrar = document.getElementById('btn-cerrar-modal');
    const btnEntendido = document.getElementById('btn-entendido-modal');
    const modalOverlay = document.getElementById('event-modal');

    if(btnCerrar) btnCerrar.onclick = cerrarAlertaEvento;
    if(btnEntendido) btnEntendido.onclick = cerrarAlertaEvento;
    if(modalOverlay) {
        modalOverlay.onclick = (e) => { if(e.target === modalOverlay) cerrarAlertaEvento(); };
    }

    const checkXLSX = setInterval(() => {
        if (typeof XLSX !== 'undefined') {
            clearInterval(checkXLSX);
            sincronizarArchivoData();
        }
    }, 100);
});