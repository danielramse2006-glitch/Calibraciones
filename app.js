// Variables globales
let equiposData = [];
let filteredData = [];
let currentEditIndex = -1;
let currentDeleteIndex = -1;

// Definir las columnas según el Excel
const COLUMNS_SHEET2 = [
    'No',
    'ID',
    'NOMBRE DEL EQUIPO',
    'Modelo',
    'select',
    'No. SERIE',
    'FABRICANTE',
    'RANGO',
    'UBICACION',
    'RESPONSIBLE',
    'Fecha de calibracion',
    'VENCIMIENTO CALIBRACIÓN',
    'Precio $',
    'VENCIMIENTO CALIBRACIÓN A 2 ANOS',
    'Etiqueta',
    'Certificado',
    'PRP5',
    'Interno / Externo',
    'Notas'
];

// Columnas para la hoja Extraviados
const COLUMNS_EXTRAVIADOS = [
    'No',
    'ID',
    'NOMBRE DEL EQUIPO',
    'Modelo',
    'No. SERIE',
    'FABRICANTE',
    'UBICACION',
    'RESPONSIBLE',
    'Fecha de calibracion',
    'VENCIMIENTO CALIBRACIÓN',
    'VENCIMIENTO CALIBRACIÓN A 2 ANOS',
    'Etiqueta',
    'Certificado',
    'PRP5'
];

// Inicializar cuando carga la página
document.addEventListener('DOMContentLoaded', function() {
    loadExcelFromRepo();
    
    // Event listener para búsqueda en tiempo real
    document.getElementById('searchInput').addEventListener('input', function(e) {
        const searchTerm = e.target.value.toLowerCase();
        if (searchTerm === '') {
            filteredData = [...equiposData];
        } else {
            filteredData = equiposData.filter(equipo => {
                return (
                    (equipo.ID || '').toString().toLowerCase().includes(searchTerm) ||
                    (equipo['NOMBRE DEL EQUIPO'] || '').toLowerCase().includes(searchTerm) ||
                    (equipo['No. SERIE'] || '').toLowerCase().includes(searchTerm) ||
                    (equipo.FABRICANTE || '').toLowerCase().includes(searchTerm) ||
                    (equipo.Modelo || '').toLowerCase().includes(searchTerm) ||
                    (equipo.UBICACION || '').toLowerCase().includes(searchTerm) ||
                    (equipo.RESPONSIBLE || '').toLowerCase().includes(searchTerm)
                );
            });
        }
        renderTable();
    });
});

// Cargar archivo Excel desde el repositorio
async function loadExcelFromRepo() {
    try {
        // Cargar el archivo Excel
        const response = await fetch('Lista_Master_de_equipos_de_calibracion_2025.xlsx');
        if (!response.ok) {
            throw new Error('No se pudo cargar el archivo Excel');
        }
        
        const data = await response.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Limpiar datos anteriores
        equiposData = [];
        
        // Procesar la hoja "Sheet2" (hoja principal)
        if (workbook.SheetNames.includes('Sheet2')) {
            const sheet = workbook.Sheets['Sheet2'];
            // Convertir a JSON, empezando desde la fila 4
            const jsonData = XLSX.utils.sheet_to_json(sheet, {
                range: 3,
                header: COLUMNS_SHEET2,
                defval: ''
            });
            
            // Procesar cada registro de Sheet2
            jsonData.forEach((item, index) => {
                if (item.ID || item['NOMBRE DEL EQUIPO']) {
                    // Crear un objeto limpio con todas las propiedades necesarias
                    const equipo = {
                        No: item.No || '',
                        ID: item.ID || '',
                        'NOMBRE DEL EQUIPO': item['NOMBRE DEL EQUIPO'] || '',
                        Modelo: item.Modelo || '',
                        'No. SERIE': item['No. SERIE'] || '',
                        FABRICANTE: item.FABRICANTE || '',
                        RANGO: item.RANGO || '',
                        UBICACION: item.UBICACION || '',
                        RESPONSIBLE: item.RESPONSIBLE || '',
                        'Fecha de calibracion': item['Fecha de calibracion'] || '',
                        'VENCIMIENTO CALIBRACIÓN': item['VENCIMIENTO CALIBRACIÓN'] || '',
                        'Precio $': item['Precio $'] || '',
                        'VENCIMIENTO CALIBRACIÓN A 2 ANOS': item['VENCIMIENTO CALIBRACIÓN A 2 ANOS'] || '',
                        Etiqueta: item.Etiqueta || '',
                        Certificado: item.Certificado || '',
                        PRP5: item.PRP5 || '',
                        'Interno / Externo': item['Interno / Externo'] || '',
                        Notas: item.Notas || '',
                        Hoja: 'Sheet2'
                    };
                    
                    // Solo agregar si tiene ID o Nombre
                    if (equipo.ID || equipo['NOMBRE DEL EQUIPO']) {
                        equiposData.push(equipo);
                    }
                }
            });
        }
        
        // Procesar la hoja "Extraviados"
        if (workbook.SheetNames.includes('Extraviados')) {
            const sheet = workbook.Sheets['Extraviados'];
            const jsonData = XLSX.utils.sheet_to_json(sheet, {
                range: 2,
                header: COLUMNS_EXTRAVIADOS,
                defval: ''
            });
            
            // Procesar cada registro de Extraviados
            jsonData.forEach((item, index) => {
                if (item.ID || item['NOMBRE DEL EQUIPO']) {
                    const equipo = {
                        No: item.No || '',
                        ID: item.ID || '',
                        'NOMBRE DEL EQUIPO': item['NOMBRE DEL EQUIPO'] || '',
                        Modelo: item.Modelo || '',
                        'No. SERIE': item['No. SERIE'] || '',
                        FABRICANTE: item.FABRICANTE || '',
                        RANGO: '',
                        UBICACION: item.UBICACION || '',
                        RESPONSIBLE: item.RESPONSIBLE || '',
                        'Fecha de calibracion': item['Fecha de calibracion'] || '',
                        'VENCIMIENTO CALIBRACIÓN': item['VENCIMIENTO CALIBRACIÓN'] || '',
                        'Precio $': '',
                        'VENCIMIENTO CALIBRACIÓN A 2 ANOS': item['VENCIMIENTO CALIBRACIÓN A 2 ANOS'] || '',
                        Etiqueta: item.Etiqueta || '',
                        Certificado: item.Certificado || '',
                        PRP5: item.PRP5 || '',
                        'Interno / Externo': '',
                        Notas: '',
                        Hoja: 'Extraviados'
                    };
                    
                    // Solo agregar si tiene ID o Nombre
                    if (equipo.ID || equipo['NOMBRE DEL EQUIPO']) {
                        equiposData.push(equipo);
                    }
                }
            });
        }
        
        // Asignar números consecutivos
        equiposData.forEach((item, index) => {
            item.No = index + 1;
        });
        
        filteredData = [...equiposData];
        
        // Guardar en localStorage
        saveToLocalStorage();
        
        // Actualizar la interfaz
        renderTable();
        updateStats();
        populateFilterOptions();
        
        console.log(`✅ Archivo cargado exitosamente: ${equiposData.length} registros`);
        
    } catch (error) {
        console.error('Error al cargar el archivo:', error);
        alert('Error al cargar el archivo Excel. Verifica que el archivo esté en la misma carpeta.');
        loadFromLocalStorage();
    }
}

// Guardar en localStorage
function saveToLocalStorage() {
    try {
        localStorage.setItem('equiposCalibration', JSON.stringify(equiposData));
        console.log('💾 Datos guardados en localStorage');
    } catch (error) {
        console.error('Error al guardar en localStorage:', error);
    }
}

// Cargar desde localStorage
function loadFromLocalStorage() {
    try {
        const data = localStorage.getItem('equiposCalibration');
        if (data) {
            equiposData = JSON.parse(data);
            filteredData = [...equiposData];
            renderTable();
            updateStats();
            populateFilterOptions();
            console.log('📂 Datos cargados desde localStorage');
        } else {
            console.log('ℹ️ No hay datos en localStorage');
        }
    } catch (error) {
        console.error('Error al cargar desde localStorage:', error);
    }
}

// Renderizar tabla
function renderTable() {
    const tbody = document.getElementById('tableBody');
    
    if (filteredData.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="19" class="empty-state">
                    <div>
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                        </svg>
                        <h3>No se encontraron resultados</h3>
                        <p>Intenta con otro criterio de búsqueda</p>
                    </div>
                </td>
            </tr>
        `;
        return;
    }
    
    tbody.innerHTML = '';
    
    filteredData.forEach((equipo, index) => {
        const row = document.createElement('tr');
        
        // Calcular estado de calibración
        const estado = calcularEstado(equipo['VENCIMIENTO CALIBRACIÓN']);
        
        row.innerHTML = `
            <td>${equipo.No || ''}</td>
            <td>${equipo.ID || ''}</td>
            <td>${equipo['NOMBRE DEL EQUIPO'] || ''}</td>
            <td>${equipo.Modelo || ''}</td>
            <td>${equipo['No. SERIE'] || ''}</td>
            <td>${equipo.FABRICANTE || ''}</td>
            <td>${equipo.RANGO || ''}</td>
            <td>${equipo.UBICACION || ''}</td>
            <td>${equipo.RESPONSIBLE || ''}</td>
            <td>${formatDate(equipo['Fecha de calibracion'])}</td>
            <td>${formatDate(equipo['VENCIMIENTO CALIBRACIÓN'])}</td>
            <td>${formatCurrency(equipo['Precio $'])}</td>
            <td>${formatDate(equipo['VENCIMIENTO CALIBRACIÓN A 2 ANOS'])}</td>
            <td>${formatSiNo(equipo.Etiqueta)}</td>
            <td>${formatSiNo(equipo.Certificado)}</td>
            <td>${equipo.PRP5 || ''}</td>
            <td>${equipo['Interno / Externo'] || ''}</td>
            <td>${equipo.Notas || ''}</td>
            <td><span class="status-badge status-${estado}">${getEstadoTexto(estado)}</span></td>
        `;
        
        // Agregar evento click para selección
        row.addEventListener('click', function() {
            // Remover selección de otras filas
            document.querySelectorAll('#tableBody tr').forEach(r => r.classList.remove('selected'));
            // Agregar selección a esta fila
            this.classList.add('selected');
        });
        
        tbody.appendChild(row);
    });
}

// Calcular estado de calibración
function calcularEstado(fechaVencimiento) {
    if (!fechaVencimiento || fechaVencimiento === '00:00:00') return 'sin-fecha';
    
    const hoy = new Date();
    const vencimiento = new Date(fechaVencimiento);
    
    // Si la fecha no es válida
    if (isNaN(vencimiento.getTime())) return 'sin-fecha';
    
    const diffTime = vencimiento - hoy;
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    if (diffDays < 0) return 'vencido';
    if (diffDays <= 30) return 'proximo';
    return 'vigente';
}

// Obtener texto del estado
function getEstadoTexto(estado) {
    const estados = {
        'vigente': 'VIGENTE',
        'vencido': 'VENCIDO',
        'proximo': 'POR VENCER',
        'sin-fecha': 'SIN FECHA'
    };
    return estados[estado] || 'SIN FECHA';
}

// Formatear fecha
function formatDate(date) {
    if (!date || date === '00:00:00') return '';
    
    try {
        const d = new Date(date);
        if (isNaN(d.getTime())) return date.toString();
        
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    } catch (error) {
        return date.toString();
    }
}

// Formatear moneda
function formatCurrency(value) {
    if (!value) return '';
    if (typeof value === 'number') {
        return '$' + value.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    }
    if (typeof value === 'string' && value.trim() !== '') {
        const num = parseFloat(value);
        if (!isNaN(num)) {
            return '$' + num.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
        }
    }
    return value;
}

// Formatear SI/NO
function formatSiNo(value) {
    if (!value) return '';
    const val = value.toString().toUpperCase().trim();
    if (val === 'SI') return '✅ SI';
    if (val === 'NO') return '❌ NO';
    if (val === 'NOK') return '❌ NOK';
    if (val === 'PD') return '⚠️ PD';
    return val;
}

// Actualizar estadísticas
function updateStats() {
    const total = equiposData.length;
    let vigentes = 0;
    let proximos = 0;
    let vencidos = 0;
    let sinFecha = 0;
    
    equiposData.forEach(equipo => {
        const estado = calcularEstado(equipo['VENCIMIENTO CALIBRACIÓN']);
        if (estado === 'vigente') vigentes++;
        else if (estado === 'proximo') proximos++;
        else if (estado === 'vencido') vencidos++;
        else if (estado === 'sin-fecha') sinFecha++;
    });
    
    document.getElementById('statTotal').textContent = total;
    document.getElementById('statVigente').textContent = vigentes;
    document.getElementById('statProximo').textContent = proximos;
    document.getElementById('statVencido').textContent = vencidos;
}

// Poblar opciones de filtros
function populateFilterOptions() {
    // Ubicaciones únicas
    const ubicaciones = [...new Set(equiposData.map(e => e.UBICACION).filter(Boolean))];
    const selectUbicacion = document.getElementById('filterUbicacion');
    selectUbicacion.innerHTML = '<option value="">Todas</option>';
    ubicaciones.sort().forEach(ub => {
        selectUbicacion.innerHTML += `<option value="${ub}">${ub}</option>`;
    });
    
    // PRP5 únicos
    const prp5s = [...new Set(equiposData.map(e => e.PRP5).filter(Boolean))];
    const selectPRP5 = document.getElementById('filterPRP5');
    selectPRP5.innerHTML = '<option value="">Todos</option>';
    prp5s.sort().forEach(prp => {
        selectPRP5.innerHTML += `<option value="${prp}">${prp}</option>`;
    });
}

// Toggle filtros
function toggleFilters() {
    const container = document.getElementById('filtersContainer');
    container.style.display = container.style.display === 'none' ? 'block' : 'none';
}

// Aplicar filtros
function applyFilters() {
    const ubicacion = document.getElementById('filterUbicacion').value;
    const prp5 = document.getElementById('filterPRP5').value;
    const tipo = document.getElementById('filterTipo').value;
    const estado = document.getElementById('filterEstado').value;
    
    filteredData = equiposData.filter(equipo => {
        let match = true;
        
        if (ubicacion && equipo.UBICACION !== ubicacion) match = false;
        if (prp5 && equipo.PRP5 !== prp5) match = false;
        if (tipo && equipo['Interno / Externo'] !== tipo) match = false;
        if (estado) {
            const estadoEquipo = calcularEstado(equipo['VENCIMIENTO CALIBRACIÓN']);
            if (estadoEquipo !== estado) match = false;
        }
        
        return match;
    });
    
    renderTable();
}

// Generar formulario
function generateForm(containerId, data = {}) {
    const container = document.getElementById(containerId);
    container.innerHTML = '';
    
    const fields = [
        { name: 'ID', type: 'text', required: true },
        { name: 'NOMBRE DEL EQUIPO', type: 'text', required: true },
        { name: 'Modelo', type: 'text' },
        { name: 'No. SERIE', type: 'text' },
        { name: 'FABRICANTE', type: 'text' },
        { name: 'RANGO', type: 'text' },
        { name: 'UBICACION', type: 'text' },
        { name: 'RESPONSIBLE', type: 'text' },
        { name: 'Fecha de calibracion', type: 'date' },
        { name: 'VENCIMIENTO CALIBRACIÓN', type: 'date' },
        { name: 'Precio $', type: 'number' },
        { name: 'VENCIMIENTO CALIBRACIÓN A 2 ANOS', type: 'date' },
        { name: 'Etiqueta', type: 'select', options: ['', 'SI', 'NO', 'NOK', 'PD'] },
        { name: 'Certificado', type: 'select', options: ['', 'SI', 'NO', 'NOK', 'PD'] },
        { name: 'PRP5', type: 'text' },
        { name: 'Interno / Externo', type: 'select', options: ['', 'Interno', 'Externo'] },
        { name: 'Notas', type: 'textarea', fullWidth: true }
    ];
    
    fields.forEach(field => {
        const formGroup = document.createElement('div');
        formGroup.className = field.fullWidth ? 'form-group full-width' : 'form-group';
        
        const label = document.createElement('label');
        label.textContent = field.name + (field.required ? ' *' : '');
        formGroup.appendChild(label);
        
        let input;
        if (field.type === 'textarea') {
            input = document.createElement('textarea');
            input.rows = 3;
        } else if (field.type === 'select') {
            input = document.createElement('select');
            field.options.forEach(opt => {
                const option = document.createElement('option');
                option.value = opt;
                option.textContent = opt;
                input.appendChild(option);
            });
        } else {
            input = document.createElement('input');
            input.type = field.type;
        }
        
        input.id = `field_${field.name.replace(/[^a-zA-Z0-9]/g, '_')}`;
        
        // Establecer valor, manejando fechas
        if (field.type === 'date' && data[field.name]) {
            const date = new Date(data[field.name]);
            if (!isNaN(date.getTime())) {
                input.value = date.toISOString().split('T')[0];
            } else {
                input.value = '';
            }
        } else {
            input.value = data[field.name] || '';
        }
        
        if (field.required) input.required = true;
        
        formGroup.appendChild(input);
        container.appendChild(formGroup);
    });
}

// Abrir modal nuevo
function openNewModal() {
    generateForm('formNew');
    document.getElementById('modalNew').style.display = 'block';
}

// Guardar nuevo
function saveNew() {
    const newEquipo = {};
    
    const fields = [
        'ID', 'NOMBRE DEL EQUIPO', 'Modelo', 'No. SERIE', 'FABRICANTE', 'RANGO',
        'UBICACION', 'RESPONSIBLE', 'Fecha de calibracion', 'VENCIMIENTO CALIBRACIÓN',
        'Precio $', 'VENCIMIENTO CALIBRACIÓN A 2 ANOS', 'Etiqueta', 'Certificado',
        'PRP5', 'Interno / Externo', 'Notas'
    ];
    
    fields.forEach(field => {
        const fieldId = `field_${field.replace(/[^a-zA-Z0-9]/g, '_')}`;
        const element = document.getElementById(fieldId);
        if (element) {
            newEquipo[field] = element.value;
        }
    });
    
    // Validar campos requeridos
    if (!newEquipo.ID || !newEquipo['NOMBRE DEL EQUIPO']) {
        alert('⚠️ Por favor completa los campos obligatorios (ID y Nombre del Equipo)');
        return;
    }
    
    // Asignar número consecutivo
    newEquipo.No = equiposData.length + 1;
    newEquipo.Hoja = 'Sheet2'; // Por defecto se agrega a Sheet2
    
    // Agregar a la lista
    equiposData.push(newEquipo);
    filteredData = [...equiposData];
    
    // Guardar y actualizar
    saveToLocalStorage();
    renderTable();
    updateStats();
    populateFilterOptions();
    
    closeModal('modalNew');
    alert('✅ Equipo agregado exitosamente');
}

// Abrir modal actualizar
function openUpdateModal() {
    document.getElementById('updateFormContainer').style.display = 'none';
    document.getElementById('btnUpdate').style.display = 'none';
    document.getElementById('updateSearch').value = '';
    document.getElementById('modalUpdate').style.display = 'block';
}

// Buscar para actualizar
function searchForUpdate() {
    const searchTerm = document.getElementById('updateSearch').value.toLowerCase().trim();
    if (!searchTerm) {
        alert('⚠️ Ingresa un ID o Nombre para buscar');
        return;
    }
    
    const index = equiposData.findIndex(e => 
        (e.ID || '').toString().toLowerCase() === searchTerm ||
        (e['NOMBRE DEL EQUIPO'] || '').toLowerCase().includes(searchTerm)
    );
    
    if (index === -1) {
        alert('❌ No se encontró el equipo');
        return;
    }
    
    currentEditIndex = index;
    generateForm('formUpdate', equiposData[index]);
    document.getElementById('updateFormContainer').style.display = 'block';
    document.getElementById('btnUpdate').style.display = 'block';
}

// Guardar actualización
function saveUpdate() {
    if (currentEditIndex === -1) return;
    
    const updatedEquipo = { ...equiposData[currentEditIndex] };
    
    const fields = [
        'ID', 'NOMBRE DEL EQUIPO', 'Modelo', 'No. SERIE', 'FABRICANTE', 'RANGO',
        'UBICACION', 'RESPONSIBLE', 'Fecha de calibracion', 'VENCIMIENTO CALIBRACIÓN',
        'Precio $', 'VENCIMIENTO CALIBRACIÓN A 2 ANOS', 'Etiqueta', 'Certificado',
        'PRP5', 'Interno / Externo', 'Notas'
    ];
    
    fields.forEach(field => {
        const fieldId = `field_${field.replace(/[^a-zA-Z0-9]/g, '_')}`;
        const element = document.getElementById(fieldId);
        if (element) {
            updatedEquipo[field] = element.value;
        }
    });
    
    // Validar campos requeridos
    if (!updatedEquipo.ID || !updatedEquipo['NOMBRE DEL EQUIPO']) {
        alert('⚠️ Por favor completa los campos obligatorios');
        return;
    }
    
    equiposData[currentEditIndex] = updatedEquipo;
    filteredData = [...equiposData];
    
    saveToLocalStorage();
    renderTable();
    updateStats();
    populateFilterOptions();
    
    closeModal('modalUpdate');
    alert('✅ Equipo actualizado exitosamente');
}

// Abrir modal eliminar
function openDeleteModal() {
    document.getElementById('deleteInfo').style.display = 'none';
    document.getElementById('btnDelete').style.display = 'none';
    document.getElementById('deleteSearch').value = '';
    document.getElementById('modalDelete').style.display = 'block';
}

// Buscar para eliminar
function searchForDelete() {
    const searchTerm = document.getElementById('deleteSearch').value.toLowerCase().trim();
    if (!searchTerm) {
        alert('⚠️ Ingresa un ID o Nombre para buscar');
        return;
    }
    
    const index = equiposData.findIndex(e => 
        (e.ID || '').toString().toLowerCase() === searchTerm ||
        (e['NOMBRE DEL EQUIPO'] || '').toLowerCase().includes(searchTerm)
    );
    
    if (index === -1) {
        alert('❌ No se encontró el equipo');
        return;
    }
    
    currentDeleteIndex = index;
    const equipo = equiposData[index];
    
    document.getElementById('deleteInfo').innerHTML = `
        <h3 style="color: #c0392b; margin-bottom: 15px;">⚠️ ¿Confirmar eliminación?</h3>
        <p><strong>ID:</strong> ${equipo.ID || 'N/A'}</p>
        <p><strong>Nombre:</strong> ${equipo['NOMBRE DEL EQUIPO'] || 'N/A'}</p>
        <p><strong>Modelo:</strong> ${equipo.Modelo || 'N/A'}</p>
        <p><strong>Ubicación:</strong> ${equipo.UBICACION || 'N/A'}</p>
        <p><strong>Responsable:</strong> ${equipo.RESPONSIBLE || 'N/A'}</p>
        <p style="margin-top: 15px; color: #c0392b;"><strong>⚠️ Esta acción no se puede deshacer</strong></p>
    `;
    
    document.getElementById('deleteInfo').style.display = 'block';
    document.getElementById('btnDelete').style.display = 'block';
}

// Confirmar eliminación
function confirmDelete() {
    if (currentDeleteIndex === -1) return;
    
    if (!confirm('¿Estás 100% seguro de eliminar este equipo? Esta acción es permanente.')) {
        return;
    }
    
    equiposData.splice(currentDeleteIndex, 1);
    
    // Reajustar números consecutivos
    equiposData.forEach((eq, index) => {
        eq.No = index + 1;
    });
    
    filteredData = [...equiposData];
    
    saveToLocalStorage();
    renderTable();
    updateStats();
    populateFilterOptions();
    
    closeModal('modalDelete');
    alert('✅ Equipo eliminado exitosamente');
}

// Cerrar modal
function closeModal(modalId) {
    document.getElementById(modalId).style.display = 'none';
    currentEditIndex = -1;
    currentDeleteIndex = -1;
}

// Cerrar modal al hacer clic fuera
window.onclick = function(event) {
    if (event.target.classList.contains('modal')) {
        event.target.style.display = 'none';
        currentEditIndex = -1;
        currentDeleteIndex = -1;
    }
}

// Descargar Excel
function downloadExcel() {
    if (equiposData.length === 0) {
        alert('⚠️ No hay datos para exportar');
        return;
    }
    
    try {
        // Crear workbook
        const wb = XLSX.utils.book_new();
        
        // Separar datos por hoja
        const datosSheet2 = equiposData.filter(e => e.Hoja === 'Sheet2');
        const datosExtraviados = equiposData.filter(e => e.Hoja === 'Extraviados');
        
        // Preparar datos para Sheet2
        const ws2Data = [
            ['Listado de calibracion de equipos'],
            [],
            COLUMNS_SHEET2,
            ...datosSheet2.map(equipo => COLUMNS_SHEET2.map(col => {
                if (col === 'No') return equipo.No;
                if (col === 'select') return ''; // Columna vacía
                return equipo[col] || '';
            }))
        ];
        
        // Preparar datos para Extraviados
        const wsExtraData = [
            ['Equipos Extraviados'],
            [],
            COLUMNS_EXTRAVIADOS,
            ...datosExtraviados.map(equipo => COLUMNS_EXTRAVIADOS.map(col => {
                if (col === 'No') return equipo.No;
                return equipo[col] || '';
            }))
        ];
        
        const ws2 = XLSX.utils.aoa_to_sheet(ws2Data);
        const wsExtra = XLSX.utils.aoa_to_sheet(wsExtraData);
        
        // Agregar las hojas al workbook
        XLSX.utils.book_append_sheet(wb, ws2, 'Calibraciones');
        XLSX.utils.book_append_sheet(wb, wsExtra, 'Extraviados');
        
        // Generar y descargar archivo
        const fecha = new Date().toISOString().split('T')[0].replace(/-/g, '');
        const hora = new Date().toTimeString().split(' ')[0].replace(/:/g, '');
        XLSX.writeFile(wb, `Calibraciones_${fecha}_${hora}.xlsx`);
        
        alert('✅ Archivo Excel descargado exitosamente');
    } catch (error) {
        console.error('Error al descargar Excel:', error);
        alert('❌ Error al generar el archivo Excel');
    }
}

// Función para recargar datos desde el archivo
function reloadFromExcel() {
    if (confirm('¿Recargar datos desde el archivo Excel? Se perderán los cambios no guardados.')) {
        loadExcelFromRepo();
    }
}

// Agregar botón de recarga al DOM
document.addEventListener('DOMContentLoaded', function() {
    // Agregar botón de recarga después de los otros botones
    const topBar = document.querySelector('.top-bar');
    const reloadBtn = document.createElement('button');
    reloadBtn.className = 'btn btn-info';
    reloadBtn.innerHTML = '🔄 Recargar';
    reloadBtn.onclick = reloadFromExcel;
    topBar.appendChild(reloadBtn);
});
