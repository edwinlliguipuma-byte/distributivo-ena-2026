// Carga de datos base integrada
let rawData = [];

// Elementos del DOM
const tableBody = document.getElementById('table-body');
const searchInput = document.getElementById('search-input');
const filterDocente = document.getElementById('filter-docente');
const filterCurso = document.getElementById('filter-curso');
const filterMateria = document.getElementById('filter-materia');
const btnReset = document.getElementById('btn-reset');
const resultsCount = document.getElementById('results-count');
const fileInput = document.getElementById('excel-file-input');

document.addEventListener('DOMContentLoaded', async () => {
    // Intenta cargar desde archivo local si existe
    await loadInitialData();

    // Eventos de Filtros
    searchInput.addEventListener('input', applyFilters);
    filterDocente.addEventListener('change', applyFilters);
    filterCurso.addEventListener('change', applyFilters);
    filterMateria.addEventListener('change', applyFilters);
    btnReset.addEventListener('click', resetFilters);
    
    // Subida manual de archivo Excel
    fileInput.addEventListener('change', handleFileUpload);

    document.getElementById('close-modal').addEventListener('click', () => {
        document.getElementById('modal-docente').style.display = 'none';
    });
});

async function loadInitialData() {
    try {
        const response = await fetch('03 marzo-Distributivo-ENA 2026.xlsx');
        if (response.ok) {
            const arrayBuffer = await response.arrayBuffer();
            parseExcelBuffer(arrayBuffer);
            return;
        }
    } catch (e) {
        console.log("Apertura directa detectada (file://). Active la carga manual si desea actualizar datos.");
    }
    
    // Si no se puede hacer fetch por restricciones file://, mostrar mensaje para subir el archivo
    if (rawData.length === 0) {
        resultsCount.textContent = "Haga clic en 'Subir Otro Excel' para cargar el archivo .xlsx";
        tableBody.innerHTML = `<tr><td colspan="7" style="text-align:center; padding: 30px;">
            <p style="margin-bottom:10px;">Para visualizar los datos abiertos desde el explorador de archivos:</p>
            <button class="btn-upload" onclick="document.getElementById('excel-file-input').click()">
                <i class="fa-solid fa-file-excel"></i> Seleccionar archivo "03 marzo-Distributivo-ENA 2026.xlsx"
            </button>
        </td></tr>`;
    }
}

function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(evt) {
        const arrayBuffer = evt.target.result;
        parseExcelBuffer(arrayBuffer);
    };
    reader.readAsArrayBuffer(file);
}

function parseExcelBuffer(arrayBuffer) {
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheet = workbook.Sheets['AREAS'] || workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    let parsedData = [];
    let currentDocente = null;

    for (let i = 3; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0) continue;

        const nomina = row[1] ? String(row[1]).trim() : '';
        const curso = row[2] ? String(row[2]).trim() : '';
        const especialidad = row[3] ? String(row[3]).trim() : '';
        const periodo = row[4] ? String(row[4]).trim() : '';
        const horasVal = row[5];
        const asignatura = row[6] ? String(row[6]).trim() : '';

        if (asignatura.toUpperCase().includes('TOTAL') || String(horasVal).toUpperCase().includes('TOTAL')) {
            continue;
        }

        if (nomina !== '' && nomina.toLowerCase() !== 'tutor' && isNaN(nomina)) {
            currentDocente = nomina;
        }

        if (currentDocente && asignatura !== '' && !asignatura.toUpperCase().includes('TOTAL')) {
            const h = parseInt(horasVal);
            const cursoCompleto = `${curso} ${especialidad}`.trim();
            
            parsedData.push({
                docente: currentDocente,
                curso: cursoCompleto,
                periodo: (periodo && periodo !== 'nan') ? periodo : 'N/A',
                horas: isNaN(h) ? 0 : h,
                asignatura: asignatura
            });
        }
    }

    rawData = parsedData;
    populateFilters();
    updateDashboard(rawData);
}

function populateFilters() {
    filterDocente.innerHTML = '<option value="">Todos los docentes</option>';
    filterCurso.innerHTML = '<option value="">Todos los cursos</option>';
    filterMateria.innerHTML = '<option value="">Todas las asignaturas</option>';

    const docentes = [...new Set(rawData.map(item => item.docente))].sort();
    const cursos = [...new Set(rawData.map(item => item.curso))].sort();
    const materias = [...new Set(rawData.map(item => item.asignatura))].sort();

    docentes.forEach(d => filterDocente.add(new Option(d, d)));
    cursos.forEach(c => filterCurso.add(new Option(c, c)));
    materias.forEach(m => filterMateria.add(new Option(m, m)));
}

function updateDashboard(data) {
    renderTable(data);
    updateStats(data);
}

function renderTable(data) {
    tableBody.innerHTML = '';
    resultsCount.textContent = `Mostrando ${data.length} registros`;

    if (data.length === 0) {
        tableBody.innerHTML = `<tr><td colspan="7" style="text-align:center; padding: 20px;">No se encontraron coincidencias.</td></tr>`;
        return;
    }

    data.forEach((item, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${index + 1}</td>
            <td><strong>${item.docente}</strong></td>
            <td>${item.curso}</td>
            <td>${item.asignatura}</td>
            <td><span class="badge badge-period">${item.periodo}</span></td>
            <td><span class="badge badge-hours">${item.horas} hrs</span></td>
            <td><button class="btn-view" onclick="verFichaDocente('${item.docente.replace(/'/g, "\\'")}')"><i class="fa-solid fa-eye"></i> Ficha</button></td>
        `;
        tableBody.appendChild(row);
    });
}

function updateStats(data) {
    const totalHoras = data.reduce((sum, item) => sum + item.horas, 0);
    const docentesUnicos = new Set(data.map(item => item.docente)).size;
    const materiasUnicas = new Set(data.map(item => item.asignatura)).size;
    const cursosUnicos = new Set(data.map(item => item.curso)).size;

    document.getElementById('stat-horas').textContent = totalHoras;
    document.getElementById('stat-docentes').textContent = docentesUnicos;
    document.getElementById('stat-materias').textContent = materiasUnicas;
    document.getElementById('stat-cursos').textContent = cursosUnicos;
}

function applyFilters() {
    const searchVal = searchInput.value.toLowerCase();
    const docenteVal = filterDocente.value;
    const cursoVal = filterCurso.value;
    const materiaVal = filterMateria.value;

    const filtered = rawData.filter(item => {
        const matchSearch = item.docente.toLowerCase().includes(searchVal) || 
                            item.asignatura.toLowerCase().includes(searchVal) || 
                            item.curso.toLowerCase().includes(searchVal);
        const matchDocente = !docenteVal || item.docente === docenteVal;
        const matchCurso = !cursoVal || item.curso === cursoVal;
        const matchMateria = !materiaVal || item.asignatura === materiaVal;

        return matchSearch && matchDocente && matchCurso && matchMateria;
    });

    updateDashboard(filtered);
}

function resetFilters() {
    searchInput.value = '';
    filterDocente.value = '';
    filterCurso.value = '';
    filterMateria.value = '';
    updateDashboard(rawData);
}

function verFichaDocente(nombreDocente) {
    const items = rawData.filter(d => d.docente === nombreDocente);
    const totalHoras = items.reduce((sum, i) => sum + i.horas, 0);
    
    let html = `
        <h2 style="color:#1e3a8a; margin-bottom:10px;"><i class="fa-solid fa-user-check"></i> Ficha de Carga Horaria</h2>
        <p style="margin-bottom: 5px; font-size:1.1rem;"><strong>Docente:</strong> ${nombreDocente}</p>
        <p style="margin-bottom: 15px;"><strong>Total Horas Asignadas:</strong> <span class="badge badge-hours">${totalHoras} Horas Semanales</span></p>
        
        <table style="width:100%; border-collapse:collapse; margin-top:10px;">
            <thead>
                <tr style="background:#f1f5f9; text-align:left;">
                    <th style="padding:8px; border-bottom:2px solid #ddd;">Curso y Paralelo</th>
                    <th style="padding:8px; border-bottom:2px solid #ddd;">Asignatura</th>
                    <th style="padding:8px; border-bottom:2px solid #ddd;">Período</th>
                    <th style="padding:8px; border-bottom:2px solid #ddd;">Horas</th>
                </tr>
            </thead>
            <tbody>
    `;

    items.forEach(i => {
        html += `
            <tr>
                <td style="padding:8px; border-bottom:1px solid #eee;">${i.curso}</td>
                <td style="padding:8px; border-bottom:1px solid #eee;">${i.asignatura}</td>
                <td style="padding:8px; border-bottom:1px solid #eee;"><span class="badge badge-period">${i.periodo}</span></td>
                <td style="padding:8px; border-bottom:1px solid #eee;"><strong>${i.horas} hrs</strong></td>
            </tr>
        `;
    });

    html += `</tbody></table>`;
    
    document.getElementById('modal-body').innerHTML = html;
    document.getElementById('modal-docente').style.display = 'flex';
}