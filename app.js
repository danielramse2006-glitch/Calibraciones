// Variables globales
let equiposData = [];
let filteredData = [];
let currentEditIndex = -1;
let currentDeleteIndex = -1;

// Definir las columnas según el Excel (para Sheet1 - hoja principal)
const COLUMNS = [
    'No',
    'ID',
    'NOMBRE DEL EQUIPO',
    'Modelo',
    'select', // Columna adicional que existe en el Excel
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
        console.log('📂 Intentando cargar archivo Excel...');
        
        // Cargar el archivo Excel
        const response = await fetch('Lista_Master_de_equipos_de_calibracion_2025.xlsx');
        if (!response.ok) {
            throw new Error(`No se pudo cargar el archivo Excel (status: ${response.status})`);
        }
        
        const data = await response.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        
        console.log('📊 Hojas disponibles:', workbook.SheetNames);
        
        // Limpiar datos anteriores
        equiposData = [];
        
        // Procesar la hoja "Sheet1" (hoja principal)
        if (workbook.SheetNames.includes('Sheet1')) {
            console.log('✅ Encontrada hoja "Sheet1"');
            const sheet = workbook.Sheets['Sheet1'];
            
            // Obtener el rango de la hoja para debug
            const range = XLSX.utils.decode_range(sheet['!ref']);
            console.log(`📈 Rango de datos: ${range.s.r}-${range.e.r} filas, ${range.s.c}-${range.e.c} columnas`);
            
            // Convertir a JSON, empezando desde la fila 4 (donde están los headers en el Excel)
            const jsonData = XLSX.utils.sheet_to_json(sheet, {
                range: 3, // Empezar desde la fila 4 (0-indexed, por eso es 3)
                header: COLUMNS,
                defval: ''
            });
            
            console.log(`📋 Registros leídos: ${jsonData.length}`);
            
            // Procesar cada registro
            jsonData.forEach((item, index) => {
                // Solo procesar si tiene ID o Nombre del Equipo
                if (item.ID || item['NOMBRE DEL EQUIPO']) {
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
                        Notas: item.Notas || ''
                    };
                    
                    equiposData.push(equipo);
                }
            });
            
            console.log(`✅ Registros procesados: ${equiposData.length}`);
            
        } else {
            console.error('❌ No se encontró la hoja "Sheet1"');
            console.log('Hojas disponibles:', workbook.SheetNames);
            
            // Intentar con la primera hoja disponible
            const firstSheetName = workbook.SheetNames[0];
            console.log(`Intentando con la primera hoja: "${firstSheetName}"`);
            
            const sheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet, {
                range: 3,
                header: COLUMNS,
                defval: ''
            });
            
            jsonData.forEach((item, index) => {
                if (item.ID || item['NOMBRE DEL EQUIPO']) {
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
                        Notas: item.Notas || ''
                    };
                    
                    equiposData.push(equipo);
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
        
        console.log(`🎉 Carga completada: ${equiposData.length} registros cargados`);
        
    } catch (error) {
        console.error('❌ Error al cargar el archivo:', error);
        alert(`Error al cargar el archivo Excel: ${error.message}\n\nVerifica que:\n1. El archivo esté en la misma carpeta\n2. Se llame "Lista_Master_de_equipos_de_calibracion_2025.xlsx"\n3. El archivo no esté corrupto`);
        
        // Si falla, intentar cargar desde localStorage
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
                        <p>Intenta con otro criterio de búsqueda o recarga los datos</p>
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
            <td><strong>${equipo.ID || ''}</strong></td>
            <td>${equipo['NOMBRE DEL EQUIPO'] || ''}</td>
            <td>${equipo.Modelo || ''}</td>
            <td>${equipo['No. SERIE'] || ''}</td>
            <td>${equipo.FABRICANTE || ''}</td>
            <td>${equipo.RANGO || ''}</td>
            <td>${equipo.UBICACION || ''}</td>
            <td>${equipo.RESPONSIBLE || ''}</td>
            <td>${formatDate(equipo['Fecha de calibracion'])}</td>
            <td><strong>${formatDate(equipo['VENCIMIENTO CALIBRACIÓN'])}</strong></td>
            <td>${formatCurrency(equipo['Precio $'])}</td>
            <td>${formatDate(equipo['VENCIMIENTO CALIBRACIÓN A 2 ANOS'])}</td>
            <td>${formatSiNo(equipo.Etiqueta)}</td>
            <td>${formatSiNo(equipo.Certificado)}</td>
            <td><span class="badge-prp5">${equipo.PRP5 || ''}</span></td>
            <td>${formatTipo(equipo['Interno / Externo'])}</td>
            <td class="notas-cell">${equipo.Notas || ''}</td>
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
    
    try {
        const hoy = new Date();
        const vencimiento = new Date(fechaVencimiento);
        
        // Si la fecha no es válida
        if (isNaN(vencimiento.getTime())) return 'sin-fecha';
        
        const diffTime = vencimiento - hoy;
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        
        if (diffDays < 0) return 'vencido';
        if (diffDays <= 30) return 'proximo';
        return 'vigente';
    } catch (error) {
        return 'sin-fecha';
    }
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
    if (!date || date === '00:00:00' || date === '') return '';
    
    try {
        const d = new Date(date);
        if (isNaN(d.getTime())) return date.toString();
        
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${day}/${month}/${year}`;
    } catch (error) {
        return date.toString();
    }
}

// Formatear moneda
function formatCurrency(value) {
    if (!value || value === '') return '';
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
    if (!value || value === '') return '';
    const val = value.toString().toUpperCase().trim();
    if (val === 'SI') return '✅ SI';
    if (val === 'NO') return '❌ NO';
    if (val === 'NOK') return '❌ NOK';
    if (val === 'PD') return '⚠️ PD';
    return val;
}

// Formatear tipo (Interno/Externo)
function formatTipo(value) {
    if (!value || value === '') return '';
    const val = value.toString().toUpperCase().trim();
    if (val === 'INTERNO') return '🏢 Interno';
    if (val === 'EXTERNO') return '🏭 Externo';
    return value;
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
    
    // Actualizar título con el total
    document.querySelector('header h1').innerHTML = `🔧 Sistema de Calibraciones <span style="font-size: 0.6em; opacity: 0.8;">(${total} equipos)</span>`;
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
    
    console.log(`📍 ${ubicaciones.length} ubicaciones cargadas`);
    console.log(`🏷️ ${prp5s.length} códigos PRP5 cargados`);
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
    console.log(`🔍 Filtros aplicados: ${filteredData.length} registros`);
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
        { name: 'VENCIMIENTO CALIBRACIÓN', type: 'date', required: true },
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
    if (!newEquipo.ID || !newEquipo['NOMBRE DEL EQUIPO'] || !newEquipo['VENCIMIENTO CALIBRACIÓN']) {
        alert('⚠️ Por favor completa los campos obligatorios (ID, Nombre del Equipo y Vencimiento)');
        return;
    }
    
    // Asignar número consecutivo
    newEquipo.No = equiposData.length + 1;
    
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
    console.log(`➕ Nuevo equipo agregado: ${newEquipo.ID} - ${newEquipo['NOMBRE DEL EQUIPO']}`);
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
    console.log(`✏️ Modo edición para equipo: ${equiposData[index].ID}`);
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
    if (!updatedEquipo.ID || !updatedEquipo['NOMBRE DEL EQUIPO'] || !updatedEquipo['VENCIMIENTO CALIBRACIÓN']) {
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
    console.log(`💾 Equipo actualizado: ${updatedEquipo.ID}`);
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
        <p><strong>Vencimiento:</strong> ${formatDate(equipo['VENCIMIENTO CALIBRACIÓN']) || 'N/A'}</p>
        <p style="margin-top: 15px; color: #c0392b; font-weight: bold;">⚠️ Esta acción no se puede deshacer</p>
    `;
    
    document.getElementById('deleteInfo').style.display = 'block';
    document.getElementById('btnDelete').style.display = 'block';
    console.log(`🗑️ Preparado para eliminar equipo: ${equipo.ID}`);
}

// Confirmar eliminación
function confirmDelete() {
    if (currentDeleteIndex === -1) return;
    
    if (!confirm('¿Estás 100% seguro de eliminar este equipo? Esta acción es permanente.')) {
        return;
    }
    
    const equipoEliminado = equiposData[currentDeleteIndex];
    
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
    console.log(`🗑️ Equipo eliminado: ${equipoEliminado.ID}`);
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
        
        // Preparar datos para el Excel
        const wsData = [
            ['Listado de calibracion de equipos'], // Fila 1
            [], // Fila 2
            COLUMNS, // Fila 3 - Headers
            ...equiposData.map(equipo => COLUMNS.map(col => {
                if (col === 'No') return equipo.No;
                if (col === 'select') return ''; // Columna vacía
                return equipo[col] || '';
            })) // Datos
        ];
        
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        
        // Ajustar anchos de columna
        const colWidths = COLUMNS.map(col => {
            if (col === 'NOMBRE DEL EQUIPO' || col === 'Notas') return { wch: 30 };
            if (col === 'UBICACION' || col === 'RESPONSIBLE') return { wch: 20 };
            if (col === 'ID' || col === 'PRP5') return { wch: 15 };
            return { wch: 12 };
        });
        ws['!cols'] = colWidths;
        
        // Agregar la hoja al workbook
        XLSX.utils.book_append_sheet(wb, ws, 'Calibraciones');
        
        // Generar y descargar archivo
        const fecha = new Date().toISOString().split('T')[0].replace(/-/g, '');
        const hora = new Date().toTimeString().split(' ')[0].replace(/:/g, '').substring(0, 4);
        XLSX.writeFile(wb, `Calibraciones_${fecha}_${hora}.xlsx`);
        
        alert('✅ Archivo Excel descargado exitosamente');
        console.log(`📥 Excel descargado: ${equiposData.length} registros`);
        
    } catch (error) {
        console.error('Error al descargar Excel:', error);
        alert('❌ Error al generar el archivo Excel: ' + error.message);
    }
}

// Función para recargar datos desde el archivo
function reloadFromExcel() {
    if (confirm('¿Recargar datos desde el archivo Excel? Se perderán los cambios no guardados.\n\n¿Deseas continuar?')) {
        console.log('🔄 Recargando datos desde Excel...');
        loadExcelFromRepo();
    }
}

// Agregar botón de recarga al DOM
document.addEventListener('DOMContentLoaded', function() {
    // Agregar botón de recarga después de los otros botones
    const topBar = document.querySelector('.top-bar');
    const reloadBtn = document.createElement('button');
    reloadBtn.className = 'btn btn-info';
    reloadBtn.innerHTML = '🔄 Recargar Excel';
    reloadBtn.onclick = reloadFromExcel;
    reloadBtn.title = 'Recargar datos desde el archivo Excel original';
    topBar.appendChild(reloadBtn);
});

// Agregar estilos adicionales para mejor visualización
document.addEventListener('DOMContentLoaded', function() {
    const style = document.createElement('style');
    style.textContent = `
        .badge-prp5 {
            background-color: #e3f2fd;
            color: #0c2461;
            padding: 2px 8px;
            border-radius: 4px;
            font-size: 0.85em;
            font-weight: 600;
            white-space: nowrap;
        }
        
        .notas-cell {
            max-width: 200px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        
        .notas-cell:hover {
            white-space: normal;
            overflow: visible;
            position: relative;
            z-index: 100;
            background: white;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
        }
        
        table tbody tr:hover .notas-cell {
            white-space: normal;
            overflow: visible;
        }
    `;
    document.head.appendChild(style);
});

// Función para limpiar filtros
function clearFilters() {
    document.getElementById('filterUbicacion').value = '';
    document.getElementById('filterPRP5').value = '';
    document.getElementById('filterTipo').value = '';
    document.getElementById('filterEstado').value = '';
    document.getElementById('searchInput').value = '';
    
    filteredData = [...equiposData];
    renderTable();
    console.log('🧹 Filtros limpiados');
}

// Agregar botón de limpiar filtros
document.addEventListener('DOMContentLoaded', function() {
    const filtersContainer = document.getElementById('filtersContainer');
    const clearBtn = document.createElement('button');
    clearBtn.className = 'btn';
    clearBtn.innerHTML = '🧹 Limpiar Filtros';
    clearBtn.onclick = clearFilters;
    clearBtn.style.marginTop = '15px';
    clearBtn.style.marginLeft = '10px';
    clearBtn.style.background = '#7f8c8d';
    clearBtn.style.color = 'white';
    
    const applyBtn = document.querySelector('.filters-container .btn-primary');
    if (applyBtn) {
        applyBtn.parentNode.insertBefore(clearBtn, applyBtn.nextSibling);
    }
});
