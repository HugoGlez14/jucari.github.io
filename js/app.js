document.addEventListener("DOMContentLoaded", () => {
    const themeToggle = document.getElementById('theme-toggle');
    if(themeToggle) {
        if (document.documentElement.classList.contains('dark-mode')) {
            themeToggle.innerHTML = '<svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 3v1m0 16v1m9-9h-1M4 12H3m15.364 6.364l-.707-.707M6.343 6.343l-.707-.707m12.728 0l-.707.707M6.343 17.657l-.707.707M16 12a4 4 0 11-8 0 4 4 0 018 0z"></path></svg>';
        } else {
            themeToggle.innerHTML = '<svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M20.354 15.354A9 9 0 018.646 3.646 9.003 9.003 0 0012 21a9.003 9.003 0 008.354-5.646z"></path></svg>';
        }
        themeToggle.addEventListener('click', () => {
            document.documentElement.classList.toggle('dark-mode');
            if (document.documentElement.classList.contains('dark-mode')) {
                localStorage.setItem('theme', 'dark');
                themeToggle.innerHTML = '<svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 3v1m0 16v1m9-9h-1M4 12H3m15.364 6.364l-.707-.707M6.343 6.343l-.707-.707m12.728 0l-.707.707M6.343 17.657l-.707.707M16 12a4 4 0 11-8 0 4 4 0 018 0z"></path></svg>';
            } else {
                localStorage.setItem('theme', 'light');
                themeToggle.innerHTML = '<svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M20.354 15.354A9 9 0 018.646 3.646 9.003 9.003 0 0012 21a9.003 9.003 0 008.354-5.646z"></path></svg>';
            }
        });
    }
});

let productos = [];
let configDeptos = []; 
let customSelection = []; 
let ordenActual = { campo: 'nombre', asc: true }; 
let recentlyEdited = []; 
const LIMITE_POR_PDF = 300; 
let productosFiltradosGlobal = [];
let cancelPdfProcess = false; // Bandera para cancelar PDF

let currentPage = 1;
const itemsPerPage = 10;
let visualFilteredProductos = [];

const inputArchivo = document.getElementById('archivo-excel');
const filtroDept = document.getElementById('filtro-dept');
const filtroOrden = document.getElementById('filtro-orden');
const filtroLote = document.getElementById('filtro-lote');
const contenedorLote = document.getElementById('contenedor-lote');
const btnImprimirNativo = document.getElementById('btn-imprimir-nativo');
const btnDescargarPdf = document.getElementById('btn-descargar-pdf');
const btnExportarBase = document.getElementById('btn-exportar-base');
const btnExportarExcel = document.getElementById('btn-exportar-excel');
const btnExportarClip = document.getElementById('btn-exportar-clip');
const tablaRegistros = document.getElementById('tabla-registros');
const pdfZone = document.getElementById('pdf-zone');
const btnToggleFiltros = document.getElementById('btn-toggle-filtros');
const startPageInput = document.getElementById('start-page');
const searchInput = document.getElementById('search-input');
const tableFilterDept = document.getElementById('table-filter-dept');
const searchCount = document.getElementById('search-count');
const btnBulkEdit = document.getElementById('btn-bulk-edit');
const addNombre = document.getElementById('add-nombre');
const addPrecio = document.getElementById('add-precio');
const addDeptoSelect = document.getElementById('add-depto-select');
const addDeptoNew = document.getElementById('add-depto-new');

window.addEventListener('beforeunload', function (e) {
    if (productos.length > 0) { e.preventDefault(); e.returnValue = ''; }
});

function mostrarCustomAlert(mensaje) {
    const alertMsg = document.getElementById('alert-msg');
    const alertModal = document.getElementById('alert-modal');
    if(alertMsg && alertModal) {
        alertMsg.innerText = mensaje; alertModal.style.display = 'flex';
    } else { alert(mensaje); }
}
window.closeModal = function(id) { document.getElementById(id).style.display = 'none'; }

function calcularDigitoVerificadorEAN(base12) {
    let suma = 0;
    for (let i = 0; i < 12; i++) { suma += parseInt(base12[i]) * (i % 2 === 0 ? 1 : 3); }
    let residuo = suma % 10; return (residuo === 0) ? 0 : 10 - residuo;
}

function generarSiguienteEAN() {
    let base12 = "750000000000"; 
    if (productos.length > 0) {
        let ultimoEan = productos[productos.length - 1].ean;
        if (ultimoEan && ultimoEan.length === 13 && !isNaN(ultimoEan)) {
            let numero = parseInt(ultimoEan.substring(0, 12), 10);
            numero += 1; base12 = String(numero).padStart(12, '0');
        }
    }
    return base12 + calcularDigitoVerificadorEAN(base12);
}

function registrarDepto(nombreDep) {
    if (!configDeptos.some(d => d.nombre === nombreDep)) { configDeptos.push({ nombre: nombreDep, activo: true }); }
}

function mostrarLoader(titulo, subtitulo, mostrarBarra = false, mostrarBtnCancelar = false) {
    document.getElementById('loader-title').innerText = titulo;
    document.getElementById('loader-subtitle').innerText = subtitulo;
    const progressBg = document.getElementById('loader-progress');
    const btnCancel = document.getElementById('btn-cancel-pdf');
    
    if(progressBg) progressBg.style.display = mostrarBarra ? 'block' : 'none';
    if(document.getElementById('pdf-progress')) document.getElementById('pdf-progress').style.width = '0%';
    
    if(btnCancel) {
        btnCancel.style.display = mostrarBtnCancelar ? 'inline-block' : 'none';
        btnCancel.onclick = () => { cancelPdfProcess = true; };
    }

    document.getElementById('loader-overlay').style.display = 'flex';
}

function ocultarLoader() { 
    document.getElementById('loader-overlay').style.display = 'none'; 
    const btnCancel = document.getElementById('btn-cancel-pdf');
    if(btnCancel) btnCancel.style.display = 'none';
}

if(inputArchivo) {
    inputArchivo.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if(!file) return;

        const fileName = file.name.toLowerCase();
        if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls') && !fileName.endsWith('.csv')) {
            mostrarCustomAlert("Archivo no válido. Sube Excel o CSV."); e.target.value = ""; return;
        }

        mostrarLoader("Analizando...", "Leyendo documento...");
        setTimeout(() => {
            const reader = new FileReader();
            reader.onload = (evt) => {
                try {
                    const data = new Uint8Array(evt.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const jsonMatrix = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
                    procesarDatosDeExcel(jsonMatrix);
                } catch (error) {
                    ocultarLoader(); mostrarCustomAlert("Error al leer el archivo Excel.");
                }
            };
            reader.readAsArrayBuffer(file);
        }, 100); 
    });
}

function procesarDatosDeExcel(datos) {
    if (!datos || datos.length === 0) {
        ocultarLoader(); mostrarCustomAlert("El archivo está vacío."); return;
    }

    let headerRowIdx = 0; let colNombre = 0, colPrecio = 1, colDepto = 2, colEan = 3; let isKnownFormat = false;

    for (let i = 0; i < Math.min(5, datos.length); i++) {
        let fila = datos[i];
        if (!fila || fila.length === 0) continue;
        let headerStr = fila.map(h => String(h || '').toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""));
        let idxNombre = headerStr.findIndex(h => h.includes('nombre') || h.includes('descripcion') || h.includes('articulo'));
        let idxPrecio = headerStr.findIndex(h => h.includes('precio') || h.includes('costo') || h.includes('venta'));
        let idxDepto = headerStr.findIndex(h => h.includes('departamento') || h.includes('categoria') || h.includes('depto'));
        let idxEan = headerStr.findIndex(h => h.includes('codigo') || h.includes('ean') || h.includes('sku') || h.includes('barras'));

        if (idxNombre !== -1 && idxPrecio !== -1) {
            colNombre = idxNombre; colPrecio = idxPrecio;
            colDepto = idxDepto !== -1 ? idxDepto : -1; colEan = idxEan !== -1 ? idxEan : -1;
            headerRowIdx = i; isKnownFormat = true; break;
        }
    }

    let startRow = isKnownFormat ? headerRowIdx + 1 : 1; 
    if (!isKnownFormat && datos[0] && datos[0].length > 1) {
        let posiblePrecio = String(datos[0][1]).replace(/[^0-9.-]+/g,"");
        if (!isNaN(parseFloat(posiblePrecio)) && parseFloat(posiblePrecio) > 0 && String(datos[0][0]).trim() !== "Nombre") {
            startRow = 0; 
        }
    }

    let procesados = 0;
    let advertenciaRepetidos = false;

    for (let i = startRow; i < datos.length; i++) { 
        const fila = datos[i];
        if(!fila || fila.length === 0) continue;
        
        let strNombre = fila[colNombre] ? String(fila[colNombre]).trim() : "";
        if (!strNombre || strNombre.toLowerCase() === "nombre") continue;

        let eanFila = colEan !== -1 && fila[colEan] ? String(fila[colEan]).replace(/\D/g, '') : "";
        
        let eanToUse = eanFila;
        if(eanFila.length >= 8) {
            let existe = productos.some(p => p.ean === eanFila);
            if(existe) { eanToUse = generarSiguienteEAN(); advertenciaRepetidos = true; }
        } else {
            eanToUse = generarSiguienteEAN();
        }

        let precioRaw = String(fila[colPrecio] || '0').replace(/[^0-9.-]+/g,"");
        let precioLeido = parseFloat(precioRaw);
        let nombreDepto = colDepto !== -1 && fila[colDepto] ? String(fila[colDepto]).trim().toUpperCase() : 'GENERAL';
        
        const prod = {
            id: Date.now() + i, 
            nombre: strNombre.toUpperCase(),
            precio: isNaN(precioLeido) ? 0 : precioLeido,
            departamento: nombreDepto,
            ean: eanToUse 
        };
        productos.push(prod);
        registrarDepto(nombreDepto);
        procesados++;
    }
    
    if (procesados === 0) {
        mostrarCustomAlert("No se detectaron productos válidos.");
    } else {
        configDeptos.sort((a,b) => a.nombre.localeCompare(b.nombre));
        consolidarCambios(false); 
        if(advertenciaRepetidos) {
            setTimeout(() => mostrarCustomAlert("El Excel contenía Códigos de Barra repetidos. Se generaron códigos nuevos automáticamente para evitar conflictos."), 500);
        }
    }
    ocultarLoader();
}

window.cambiarOrdenGlobal = function() {
    const val = filtroOrden.value;
    if(val === 'nombre_asc') { ordenActual = { campo: 'nombre', asc: true }; }
    else if(val === 'precio_asc') { ordenActual = { campo: 'precio', asc: true }; }
    else if(val === 'precio_desc') { ordenActual = { campo: 'precio', asc: false }; }
    consolidarCambios(true);
};

function ordenarBaseDatos() {
    let mapOrden = new Map();
    configDeptos.forEach((d, i) => mapOrden.set(d.nombre, { orden: i, activo: d.activo }));

    productos.sort((a, b) => {
        let ordenA = mapOrden.has(a.departamento) ? mapOrden.get(a.departamento).orden : 999;
        let ordenB = mapOrden.has(b.departamento) ? mapOrden.get(b.departamento).orden : 999;
        let diffDepto = ordenA - ordenB;
        if (diffDepto !== 0) return diffDepto;

        let valA = a[ordenActual.campo]; let valB = b[ordenActual.campo];
        if (typeof valA === 'string') { return ordenActual.asc ? valA.localeCompare(valB) : valB.localeCompare(valA); } 
        else { return ordenActual.asc ? (valA - valB) : (valB - valA); }
    });
}

window.manejarNuevoDepto = function() {
    if(addDeptoSelect.value === "NEW") {
        addDeptoNew.style.display = "block"; addDeptoNew.focus();
    } else {
        addDeptoNew.style.display = "none"; addDeptoNew.value = "";
    }
};

function consolidarCambios(mostrarCarga = true) {
    if(mostrarCarga) mostrarLoader("Actualizando...", "Aplicando los cambios.");
    setTimeout(() => {
        productos.forEach(p => registrarDepto(p.departamento));
        ordenarBaseDatos(); 
        actualizarVistas();
        if(filtroDept.value) { filtroDept.dispatchEvent(new Event('change')); }
        if(mostrarCarga) ocultarLoader();
    }, 100);
}

if(document.getElementById('btn-agregar')) {
    document.getElementById('btn-agregar').addEventListener('click', () => {
        const nombre = addNombre.value.trim().toUpperCase();
        let precio = parseFloat(addPrecio.value);
        let depto = addDeptoSelect.value;
        if(depto === "NEW") { depto = addDeptoNew.value.trim().toUpperCase(); }
        let customEan = document.getElementById('add-ean') ? document.getElementById('add-ean').value.trim() : "";

        if (!nombre || isNaN(precio) || !depto) {
            mostrarCustomAlert("Completa Nombre, Precio y Departamento."); return;
        }

        if (customEan !== "" && productos.some(p => p.ean === customEan)) {
            mostrarCustomAlert("Ese código de barras ya pertenece a otro producto en el sistema.");
            return;
        }

        let finalEan = customEan !== "" ? customEan : generarSiguienteEAN();

        productos.push({ id: Date.now(), nombre, precio, departamento: depto, ean: finalEan });
        
        addNombre.value = ""; addPrecio.value = ""; addDeptoSelect.value = ""; addDeptoNew.value = ""; addDeptoNew.style.display = "none";
        if(document.getElementById('add-ean')) document.getElementById('add-ean').value = "";
        
        recentlyEdited = [productos[productos.length - 1].id];
        consolidarCambios();
        setTimeout(() => { recentlyEdited = []; filtrarTablaVisual(); }, 5000);
    });
}

window.eliminarProducto = function(id) {
    if(confirm("¿Seguro que quieres borrar este producto?")) {
        productos = productos.filter(p => p.id !== id);
        customSelection = customSelection.filter(selId => selId !== id);
        document.getElementById('count-selection').innerText = customSelection.length;
        consolidarCambios();
    }
};

window.editarProducto = function(id) {
    let idx = productos.findIndex(p => p.id === id);
    if(idx === -1) return;
    let p = productos[idx];
    document.getElementById('edit-id').value = p.id;
    document.getElementById('edit-nombre').value = p.nombre;
    document.getElementById('edit-precio').value = p.precio;
    document.getElementById('edit-depto').value = p.departamento;
    if(document.getElementById('edit-ean')) document.getElementById('edit-ean').value = p.ean;
    document.getElementById('edit-modal').style.display = 'flex';
};

window.guardarEdicion = function() {
    let id = parseInt(document.getElementById('edit-id').value);
    let n = document.getElementById('edit-nombre').value.trim();
    let pr = parseFloat(document.getElementById('edit-precio').value);
    let d = document.getElementById('edit-depto').value.trim();
    let newEan = document.getElementById('edit-ean') ? document.getElementById('edit-ean').value.trim() : "";

    if (!n || isNaN(pr) || !d) {
        mostrarCustomAlert("Los campos principales en la edición son obligatorios."); return;
    }

    if (newEan !== "" && productos.some(p => p.ean === newEan && p.id !== id)) {
        mostrarCustomAlert("Ese código de barras ya pertenece a otro producto distinto.");
        return;
    }

    let idx = productos.findIndex(p => p.id === id);
    if(idx !== -1) {
        productos[idx].nombre = n.toUpperCase(); 
        productos[idx].precio = pr; 
        productos[idx].departamento = d.toUpperCase();
        if(newEan !== "") productos[idx].ean = newEan;
        recentlyEdited = [id]; 
    }
    closeModal('edit-modal');
    consolidarCambios();
    setTimeout(() => { recentlyEdited = []; filtrarTablaVisual(); }, 5000);
};

window.abrirModalEdicionMasiva = function() {
    if (customSelection.length === 0) { mostrarCustomAlert("Selecciona al menos un producto."); return; }
    document.getElementById('bulk-count').innerText = customSelection.length;
    document.getElementById('bulk-prefijo').value = ""; document.getElementById('bulk-sufijo').value = ""; document.getElementById('bulk-precio').value = "";
    const bulkDeptoSelect = document.getElementById('bulk-depto-select');
    if (bulkDeptoSelect) bulkDeptoSelect.value = "";
    document.getElementById('bulk-depto-new').value = ""; document.getElementById('bulk-depto-new').style.display = "none";
    document.getElementById('bulk-edit-modal').style.display = 'flex';
};

window.manejarNuevoDeptoBulk = function() {
    if(document.getElementById('bulk-depto-select').value === "NEW") {
        document.getElementById('bulk-depto-new').style.display = "block"; document.getElementById('bulk-depto-new').focus();
    } else {
        document.getElementById('bulk-depto-new').style.display = "none"; document.getElementById('bulk-depto-new').value = "";
    }
};

window.aplicarEdicionMasiva = function() {
    let prefijo = document.getElementById('bulk-prefijo').value.trim();
    let sufijo = document.getElementById('bulk-sufijo').value.trim();
    let nuevoPrecio = parseFloat(document.getElementById('bulk-precio').value);
    let nuevoDepto = document.getElementById('bulk-depto-select').value;
    if(nuevoDepto === "NEW") nuevoDepto = document.getElementById('bulk-depto-new').value.trim().toUpperCase();

    if (!prefijo && !sufijo && isNaN(nuevoPrecio) && !nuevoDepto) { mostrarCustomAlert("No ingresaste ningún cambio."); return; }

    recentlyEdited = [...customSelection];

    customSelection.forEach(id => {
        let p = productos.find(prod => prod.id === id);
        if(p) {
            if(prefijo) p.nombre = prefijo + " " + p.nombre;
            if(sufijo) p.nombre = p.nombre + " " + sufijo;
            if(!isNaN(nuevoPrecio)) p.precio = nuevoPrecio;
            if(nuevoDepto) p.departamento = nuevoDepto;
        }
    });

    closeModal('bulk-edit-modal');
    customSelection = []; document.getElementById('count-selection').innerText = 0;
    consolidarCambios();
    setTimeout(() => { recentlyEdited = []; filtrarTablaVisual(); }, 5000);
};

window.toggleSeleccionar = function(id, isChecked, renderUI = true) {
    if (isChecked) {
        if (!customSelection.includes(id)) customSelection.push(id);
    } else {
        customSelection = customSelection.filter(selId => selId !== id);
    }
    document.getElementById('count-selection').innerText = customSelection.length;
    if(renderUI) {
        llenarDesplegable(); 
        if(customSelection.length === 0) { const cAll = document.getElementById('chk-all'); if(cAll) cAll.checked = false;}
    }
    if(renderUI && filtroDept.value === "CUSTOM") { filtroDept.dispatchEvent(new Event('change')); }
};

window.toggleAllVisual = function(isChecked) {
    const filas = document.querySelectorAll("#tabla-datos-body tr");
    filas.forEach(fila => {
        if (fila.style.display !== "none" && !fila.querySelector('td[colspan]')) {
            const checkbox = fila.querySelector('.chk-seleccionar');
            if (checkbox && checkbox.checked !== isChecked) {
                checkbox.checked = isChecked;
                toggleSeleccionar(parseInt(checkbox.dataset.id), isChecked, false); 
            }
        }
    });
    llenarDesplegable();
    if(filtroDept.value === "CUSTOM") { filtroDept.dispatchEvent(new Event('change')); }
};

window.abrirModalSeleccion = function() {
    dibujarListaSeleccion(); document.getElementById('seleccion-modal').style.display = 'flex';
};

window.limpiarSeleccion = function() {
    customSelection = []; document.getElementById('count-selection').innerText = 0;
    const chkAll = document.getElementById('chk-all'); if (chkAll) chkAll.checked = false;
    if(filtroDept.value === "CUSTOM") { filtroDept.value = ""; }
    consolidarCambios(); closeModal('seleccion-modal');
};

function dibujarListaSeleccion() {
    const lista = document.getElementById('lista-seleccion-custom');
    lista.innerHTML = '';
    if(customSelection.length === 0) {
        lista.innerHTML = '<p style="color:var(--secondary); text-align:center; padding: 20px;">Vacio.</p>'; return;
    }
    customSelection.forEach((id, index) => {
        const prod = productos.find(p => p.id === id);
        if(!prod) return;
        lista.innerHTML += `
            <div class="sel-item">
                <div><strong>${prod.nombre}</strong> <br><small style="color:var(--secondary);">${prod.departamento} | $${prod.precio.toFixed(2)}</small></div>
                <div style="display:flex; gap:5px;">
                    <button class="btn-action" onclick="moverSel(${index}, -1)"><svg width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 15l7-7 7 7"></path></svg></button>
                    <button class="btn-action" onclick="moverSel(${index}, 1)"><svg width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7"></path></svg></button>
                    <button class="btn-action btn-danger" onclick="eliminarDeSel(${id})"><svg width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg></button>
                </div>
            </div>
        `;
    });
}

window.moverSel = function(index, dir) {
    if (index + dir < 0 || index + dir >= customSelection.length) return;
    let temp = customSelection[index];
    customSelection[index] = customSelection[index + dir]; customSelection[index + dir] = temp;
    dibujarListaSeleccion();
    if (filtroDept.value === "CUSTOM") filtroDept.dispatchEvent(new Event('change'));
};

window.eliminarDeSel = function(id) {
    toggleSeleccionar(id, false);
    dibujarListaSeleccion(); filtrarTablaVisual(); 
    if (filtroDept.value === "CUSTOM") filtroDept.dispatchEvent(new Event('change'));
};

let searchTimeout;
window.triggerSearch = function() {
    clearTimeout(searchTimeout);
    searchTimeout = setTimeout(() => { filtrarTablaVisual(); }, 200); 
};

window.ordenarTabla = function(campo) {
    if (ordenActual.campo === campo) { ordenActual.asc = !ordenActual.asc; } 
    else { ordenActual.campo = campo; ordenActual.asc = true; }
    consolidarCambios();
};

window.filtrarTablaVisual = function() {
    const query = searchInput ? searchInput.value.toLowerCase().trim() : '';
    const deptoSeleccionado = tableFilterDept ? tableFilterDept.value.toLowerCase() : '';
    const terminosBusqueda = query.split(" ").filter(t => t !== "");

    visualFilteredProductos = productos.filter(p => {
        const textoFila = `${p.nombre} ${p.departamento} ${p.ean}`.toLowerCase();
        const coincideTexto = terminosBusqueda.length === 0 || terminosBusqueda.every(t => textoFila.includes(t));
        const coincideDepto = deptoSeleccionado === "" || p.departamento.toLowerCase() === deptoSeleccionado;
        return coincideTexto && coincideDepto;
    });

    currentPage = 1; 
    renderizarTablaPaginada();
};

function renderizarTablaPaginada() {
    if(!tablaRegistros) return;
    
    const totalPages = Math.ceil(visualFilteredProductos.length / itemsPerPage);
    if(currentPage > totalPages && totalPages > 0) currentPage = totalPages;
    if(currentPage < 1) currentPage = 1;

    const startIdx = (currentPage - 1) * itemsPerPage;
    const endIdx = startIdx + itemsPerPage;
    const paginaActualDatos = visualFilteredProductos.slice(startIdx, endIdx);

    let html = `<table><thead><tr>
        <th style="width: 40px; text-align:center;"><input type="checkbox" id="chk-all" onchange="toggleAllVisual(this.checked)" title="Seleccionar Visibles"></th>
        <th style="width: 40px; text-align:center;">#</th>
        <th class="sortable-th" onclick="ordenarTabla('nombre')">Nombre ↕️</th>
        <th class="sortable-th" onclick="ordenarTabla('precio')">Precio ↕️</th>
        <th class="sortable-th" onclick="ordenarTabla('departamento')">Depto ↕️</th>
        <th>EAN-13</th>
        <th style="text-align:center;">Acciones</th>
    </tr></thead><tbody id="tabla-datos-body">`;

    if(paginaActualDatos.length === 0) {
        html += `<tr><td colspan="7" style="text-align:center; padding: 30px; color: var(--text-muted);">No se encontraron resultados.</td></tr>`;
    } else {
        const iconEdit = `<svg width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path></svg>`;
        const iconTrash = `<svg width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg>`;

        paginaActualDatos.forEach((p, index) => {
            const rowNum = startIdx + index + 1;
            const textoPrecio = p.precio === 0 ? '---' : `$${p.precio.toFixed(2)}`;
            const isChecked = customSelection.includes(p.id) ? 'checked' : '';
            const isHighlighted = recentlyEdited.includes(p.id) ? 'class="highlight-row"' : '';

            html += `<tr ${isHighlighted}>
                <td style="text-align:center;"><input type="checkbox" class="chk-seleccionar" data-id="${p.id}" onchange="toggleSeleccionar(${p.id}, this.checked)" ${isChecked}></td>
                <td style="text-align:center; color: var(--secondary); font-size: 12px; font-weight:700;">${rowNum}</td>
                <td style="font-weight: 500;">${p.nombre}</td>
                <td style="color:var(--success); font-weight:700;">${textoPrecio}</td>
                <td>${p.departamento}</td>
                <td style="font-family: monospace; color: var(--secondary);">${p.ean}</td>
                <td>
                    <div class="actions-cell">
                        <button class="btn btn-action btn-edit" onclick="editarProducto(${p.id})" title="Editar">${iconEdit}</button>
                        <button class="btn btn-action btn-danger" onclick="eliminarProducto(${p.id})" title="Borrar">${iconTrash}</button>
                    </div>
                </td>
            </tr>`;
        });
    }
    
    html += `</tbody></table>`;
    tablaRegistros.innerHTML = html;
    if (searchCount) { searchCount.innerText = `Mostrando: ${visualFilteredProductos.length}`; }

    renderizarControlesPaginacion(totalPages);
}

function renderizarControlesPaginacion(totalPages) {
    let pagContainer = document.getElementById('pagination-container');
    if(!pagContainer) return;
    if (totalPages <= 1) { pagContainer.innerHTML = ''; return; }

    let html = '';
    html += `<button class="btn btn-secondary" style="height:35px; padding:0 10px;" ${currentPage === 1 ? 'disabled' : ''} onclick="cambiarPagina(${currentPage - 1})">Ant</button>`;
    
    let startP = Math.max(1, currentPage - 2);
    let endP = Math.min(totalPages, currentPage + 2);
    
    if(startP > 1) {
        html += `<button class="btn btn-secondary" style="height:35px; padding:0 15px;" onclick="cambiarPagina(1)">1</button>`;
        if(startP > 2) html += `<span style="align-self: center; color: var(--secondary);">...</span>`;
    }

    for (let i = startP; i <= endP; i++) {
        let btnClass = (i === currentPage) ? 'btn-primary' : 'btn-secondary';
        html += `<button class="btn ${btnClass}" style="height:35px; padding:0 15px;" onclick="cambiarPagina(${i})">${i}</button>`;
    }

    if(endP < totalPages) {
        if(endP < totalPages - 1) html += `<span style="align-self: center; color: var(--secondary);">...</span>`;
        html += `<button class="btn btn-secondary" style="height:35px; padding:0 15px;" onclick="cambiarPagina(${totalPages})">${totalPages}</button>`;
    }

    html += `<button class="btn btn-secondary" style="height:35px; padding:0 10px;" ${currentPage === totalPages ? 'disabled' : ''} onclick="cambiarPagina(${currentPage + 1})">Sig</button>`;
    pagContainer.innerHTML = html;
}

window.cambiarPagina = function(nuevaPagina) {
    currentPage = nuevaPagina;
    renderizarTablaPaginada();
};

function dibujarTabla() {
    filtrarTablaVisual();
}

window.abrirModalDeptos = function() {
    dibujarPanelFiltrosYOrden(); document.getElementById('dept-modal').style.display = 'flex';
};

function dibujarPanelFiltrosYOrden() {
    const lista = document.getElementById('lista-deptos-modal');
    lista.innerHTML = "";
    if(configDeptos.length > 0) {
        configDeptos.forEach((dep, index) => {
            const isChecked = dep.activo ? "checked" : "";
            lista.innerHTML += `
                <div class="sel-item" style="padding: 10px 15px;">
                    <label style="margin:0; flex:1; display:flex; align-items:center; cursor:pointer;">
                        <input type="checkbox" value="${dep.nombre}" ${isChecked} onchange="toggleActivoDepto(${index}, this.checked)" style="width:18px; height:18px;"> 
                        <strong style="margin-left: 10px; font-weight:500;">${dep.nombre}</strong>
                    </label>
                    <div style="display:flex; gap:5px;">
                        <button class="btn-action" onclick="moverDepto(${index}, -1)"><svg width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 15l7-7 7 7"></path></svg></button>
                        <button class="btn-action" onclick="moverDepto(${index}, 1)"><svg width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7"></path></svg></button>
                    </div>
                </div>
            `;
        });
    } else {
        lista.innerHTML = '<p style="color:var(--secondary); text-align:center;">No hay departamentos.</p>';
    }
}

window.toggleActivoDepto = function(index, estaActivo) {
    configDeptos[index].activo = estaActivo;
    if(filtroDept.value === "TODOS") { filtroDept.dispatchEvent(new Event('change')); }
};

window.moverDepto = function(index, direccion) {
    if (index + direccion < 0 || index + direccion >= configDeptos.length) return; 
    let temp = configDeptos[index];
    configDeptos[index] = configDeptos[index + direccion];
    configDeptos[index + direccion] = temp;
    dibujarPanelFiltrosYOrden();
    if(filtroDept.value === "TODOS") { filtroDept.dispatchEvent(new Event('change')); }
};

function actualizarVistas() {
    if(productos.length > 0) {
        if(btnExportarBase) btnExportarBase.disabled = false; 
        if(btnExportarExcel) btnExportarExcel.disabled = false; 
        if(btnExportarClip) btnExportarClip.disabled = false;
        searchInput.disabled = false; tableFilterDept.disabled = false; btnBulkEdit.disabled = false;
        btnToggleFiltros.disabled = false; filtroOrden.disabled = false;
    } else {
        if(btnExportarBase) btnExportarBase.disabled = true; 
        if(btnExportarExcel) btnExportarExcel.disabled = true; 
        if(btnExportarClip) btnExportarClip.disabled = true;
        searchInput.disabled = true; tableFilterDept.disabled = true; btnBulkEdit.disabled = true;
        btnToggleFiltros.disabled = true; filtroOrden.disabled = true;
    }
    dibujarTabla(); llenarDesplegable(); 
}

function llenarDesplegable() {
    const valorActual = filtroDept.value;
    const valorTablaActual = tableFilterDept.value;
    
    filtroDept.innerHTML = '<option value="">-- Selecciona Depto --</option>';
    filtroDept.innerHTML += '<option value="TODOS" style="font-weight:bold; color:var(--primary);">📁 TODOS LOS DEPARTAMENTOS</option>';
    
    if(customSelection.length > 0) {
        filtroDept.innerHTML += `<option value="CUSTOM" style="font-weight:bold; color:var(--success);">🛍️ CATÁLOGO PERSONALIZADO</option>`;
    }

    tableFilterDept.innerHTML = '<option value="">📁 Todos los departamentos</option>';
    
    const addDeptoSel = document.getElementById('add-depto-select');
    const bulkDeptoSel = document.getElementById('bulk-depto-select');
    const valAdd = addDeptoSel.value;
    const valBulk = bulkDeptoSel ? bulkDeptoSel.value : "";
    
    addDeptoSel.innerHTML = '<option value="">-- Seleccionar --</option><option value="NEW" style="font-weight:bold; color:var(--primary);">➕ Nuevo Depto...</option>';
    if(bulkDeptoSel) bulkDeptoSel.innerHTML = '<option value="">-- No cambiar --</option><option value="NEW" style="font-weight:bold; color:var(--primary);">➕ Nuevo Depto...</option>';

    configDeptos.forEach(dep => {
        filtroDept.innerHTML += `<option value="${dep.nombre}">Depto: ${dep.nombre}</option>`;
        tableFilterDept.innerHTML += `<option value="${dep.nombre}">${dep.nombre}</option>`;
        addDeptoSel.innerHTML += `<option value="${dep.nombre}">${dep.nombre}</option>`;
        if(bulkDeptoSel) bulkDeptoSel.innerHTML += `<option value="${dep.nombre}">${dep.nombre}</option>`;
    });
    
    filtroDept.disabled = false;
    
    if (valorActual && Array.from(filtroDept.options).some(o => o.value === valorActual)) {
        filtroDept.value = valorActual;
    } else {
        filtroDept.value = "";
        if(pdfZone) pdfZone.innerHTML = "";
        if(btnImprimirNativo) btnImprimirNativo.disabled = true; 
        if(btnDescargarPdf) btnDescargarPdf.disabled = true; 
        if(contenedorLote) contenedorLote.style.display = 'none';
    }

    if (valorTablaActual) tableFilterDept.value = valorTablaActual;
    if(valAdd !== "NEW") addDeptoSel.value = valAdd;
    if(bulkDeptoSel && valBulk !== "NEW") bulkDeptoSel.value = valBulk;
}

if(btnExportarBase) {
    btnExportarBase.addEventListener('click', () => {
        if(productos.length === 0) return;
        const datosParaExcel = productos.map(p => ({ "Nombre": p.nombre, "Precio": p.precio, "Departamento": p.departamento, "EAN13": p.ean }));
        const hoja = XLSX.utils.json_to_sheet(datosParaExcel);
        const libro = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(libro, hoja, "Base_Datos");
        XLSX.writeFile(libro, "Bazar_Base_Actualizada.xlsx");
    });
}

if(btnExportarExcel) {
    btnExportarExcel.addEventListener('click', () => {
        if(productos.length === 0) return;
        const datosParaExcel = productos.map(p => ({ "Código de Barras": p.ean, "Descripción": p.nombre, "Precio de Costo": "", "Precio de Venta": p.precio, "Precio de Mayoreo": p.precio, "Departamento": p.departamento, "Se vende por": "Unidad", "Usa Inventario": "Sí", "Inventario": 10, "Mínimo": 1 }));
        const hoja = XLSX.utils.json_to_sheet(datosParaExcel);
        const libro = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(libro, hoja, "Articulos_Eleventa");
        XLSX.writeFile(libro, "Importar_A_Eleventa.xlsx");
    });
}

if(btnExportarClip) {
    btnExportarClip.addEventListener('click', () => {
        if(productos.length === 0) return;
        const datosParaClip = productos.map(p => ({ "Nombre": p.nombre, "Categoría": p.departamento, "Precio": p.precio, "Código de barras": p.ean, "SKU": p.ean, "Inventario": 10 }));
        const hoja = XLSX.utils.json_to_sheet(datosParaClip);
        const libro = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(libro, hoja, "Clip_Import");
        XLSX.writeFile(libro, "Importar_A_Clip.xlsx");
    });
}

if(filtroDept) {
    filtroDept.addEventListener('change', () => {
        const seleccion = filtroDept.value;
        if(!seleccion) {
            if(btnImprimirNativo) btnImprimirNativo.disabled = true; 
            if(btnDescargarPdf) btnDescargarPdf.disabled = true; 
            if(contenedorLote) contenedorLote.style.display = 'none'; 
            if(pdfZone) pdfZone.innerHTML = "";
            return;
        }
        if(btnImprimirNativo) btnImprimirNativo.disabled = false; 
        if(btnDescargarPdf) btnDescargarPdf.disabled = false;

        if (seleccion === "TODOS") {
            let mapOrden = new Map();
            configDeptos.forEach((d, i) => mapOrden.set(d.nombre, { orden: i, activo: d.activo }));
            let productosValidos = productos.filter(p => { let cfg = mapOrden.get(p.departamento); return cfg && cfg.activo; });
            productosFiltradosGlobal = productosValidos.sort((a, b) => {
                let diffDepto = mapOrden.get(a.departamento).orden - mapOrden.get(b.departamento).orden;
                if (diffDepto !== 0) return diffDepto;
                let valA = a[ordenActual.campo]; let valB = b[ordenActual.campo];
                if (typeof valA === 'string') { return ordenActual.asc ? valA.localeCompare(valB) : valB.localeCompare(valA); } 
                else { return ordenActual.asc ? (valA - valB) : (valB - valA); }
            });
        } else if (seleccion === "CUSTOM") {
            productosFiltradosGlobal = customSelection.map(id => productos.find(p => p.id === id)).filter(p => p);
        } else {
            let productosValidos = productos.filter(p => p.departamento === seleccion);
            productosFiltradosGlobal = productosValidos.sort((a, b) => {
                let valA = a[ordenActual.campo]; let valB = b[ordenActual.campo];
                if (typeof valA === 'string') { return ordenActual.asc ? valA.localeCompare(valB) : valB.localeCompare(valA); } 
                else { return ordenActual.asc ? (valA - valB) : (valB - valA); }
            });
        }
        evaluarLotesYRenderizar();
    });
}

function evaluarLotesYRenderizar() {
    const totalLotes = Math.ceil(productosFiltradosGlobal.length / LIMITE_POR_PDF);
    if(!filtroLote) return;
    filtroLote.innerHTML = '';
    if (totalLotes > 1) {
        if(contenedorLote) contenedorLote.style.display = 'flex';
        for (let i = 1; i <= totalLotes; i++) {
            let inicio = ((i - 1) * LIMITE_POR_PDF) + 1;
            let fin = Math.min(i * LIMITE_POR_PDF, productosFiltradosGlobal.length);
            filtroLote.innerHTML += `<option value="${i}">Lote ${i} (Prod. ${inicio} al ${fin})</option>`;
        }
    } else {
        if(contenedorLote) contenedorLote.style.display = 'none';
    }
    renderizarGrid(1); 
}

if(filtroLote) {
    filtroLote.addEventListener('change', (e) => { renderizarGrid(parseInt(e.target.value)); });
}

function renderizarGrid(numeroLote) {
    if(!pdfZone) return;
    pdfZone.innerHTML = "";
    let inicio = (numeroLote - 1) * LIMITE_POR_PDF;
    let fin = inicio + LIMITE_POR_PDF;
    let loteAImprimir = productosFiltradosGlobal.slice(inicio, fin);
    let currentDept = "";

    const TICKET_HEIGHT = 175; 
    const TICKET_MARGIN_BOTTOM = 15;
    const ROW_HEIGHT = TICKET_HEIGHT + TICKET_MARGIN_BOTTOM; 
    const HEADER_HEIGHT = 50; 
    const MAX_HEIGHT_PER_PAGE = 1000; 
    
    let currentPageDiv = null;
    let currentHeight = 0;
    let itemsInRow = 0; 
    let globalPageNum = parseInt(startPageInput.value) || 1;

    function crearNuevaHoja() {
        if(currentPageDiv) {
            currentPageDiv.style.pageBreakAfter = "always";
            const footerHtml = document.createElement('div');
            footerHtml.className = 'page-footer';
            footerHtml.innerText = `Página ${globalPageNum}`;
            currentPageDiv.appendChild(footerHtml);
            pdfZone.appendChild(currentPageDiv);
            globalPageNum++;
        }
        currentPageDiv = document.createElement('div');
        currentPageDiv.className = 'pdf-page';
        currentHeight = 0;
        itemsInRow = 0;
    }

    crearNuevaHoja(); 

    loteAImprimir.forEach((prod, idx) => {
        if (prod.departamento !== currentDept) {
            if (idx !== 0 && filtroDept.value !== "CUSTOM") { crearNuevaHoja(); }
            currentDept = prod.departamento;
            const header = document.createElement('div');
            header.className = 'dept-header';
            header.innerText = `DEPARTAMENTO: ${currentDept}`;
            currentPageDiv.appendChild(header);
            currentHeight += HEADER_HEIGHT;
        }

        let spaceNeeded = (itemsInRow === 0) ? ROW_HEIGHT : 0;
        if (currentHeight + spaceNeeded > MAX_HEIGHT_PER_PAGE) { crearNuevaHoja(); }

        const card = document.createElement('div');
        card.className = 'product-card';
        const svgId = `barcode-${prod.id}-${idx}`;

        let htmlPrecio = prod.precio === 0 ? `<div class="precio-vacio"></div>` : `<div class="prod-price">$${prod.precio.toFixed(2)}</div>`;

        card.innerHTML = `
            <div class="prod-name">${prod.nombre}</div>
            ${htmlPrecio}
            <div class="prod-label">PRECIO DE MAYOREO</div>
            <div class="barcode-box"><svg id="${svgId}"></svg></div>
        `;
        currentPageDiv.appendChild(card);
        
        itemsInRow++;
        if (itemsInRow === 2) { currentHeight += ROW_HEIGHT; itemsInRow = 0; }

        try {
            setTimeout(() => {
                try{ JsBarcode(`#${svgId}`, prod.ean, { format: "EAN13", height: 40, displayValue: true, fontSize: 13, margin: 0 }); }
                catch(e){ JsBarcode(`#${svgId}`, prod.ean, { format: "CODE128", height: 40, displayValue: true, fontSize: 13, margin: 0 }); }
            }, 10);
        } catch(e) {}
    });

    if(currentPageDiv) {
        currentPageDiv.style.pageBreakAfter = "auto";
        const footerHtml = document.createElement('div');
        footerHtml.className = 'page-footer';
        footerHtml.innerText = `Página ${globalPageNum}`;
        currentPageDiv.appendChild(footerHtml);
        pdfZone.appendChild(currentPageDiv);
    }
}

if(btnImprimirNativo) {
    btnImprimirNativo.addEventListener('click', () => {
        renderizarGrid(parseInt(filtroLote.value) || 1);
        setTimeout(() => window.print(), 300);
    });
}

if(btnDescargarPdf) {
    btnDescargarPdf.addEventListener('click', async () => {
        renderizarGrid(parseInt(filtroLote.value) || 1);
        mostrarLoader("Generando PDF...", "Este proceso puede tardar un poco.", true, true);
        
        cancelPdfProcess = false; // Reiniciamos bandera de cancelar
        const progressBar = document.getElementById('pdf-progress');
        if(progressBar) progressBar.style.width = '0%';
        
        const pdfWrapper = document.getElementById('pdf-wrapper');
        pdfWrapper.style.overflow = "visible";
        window.scrollTo(0, 0); 

        let nombreSufijo = contenedorLote.style.display === 'flex' ? `_Parte${filtroLote.value}` : "";
        const nombreArchivo = filtroDept.value === "TODOS" || filtroDept.value === "CUSTOM"
            ? `Catalogo_Bazar${nombreSufijo}.pdf` : `Catalogo_${filtroDept.value}${nombreSufijo}.pdf`;

        setTimeout(async () => {
            const paginas = document.querySelectorAll('.pdf-page');
            const { jsPDF } = window.jspdf;
            const pdf = new jsPDF({ unit: 'px', format: [794, 1122], orientation: 'portrait' });

            try {
                for (let i = 0; i < paginas.length; i++) {
                    // Si el usuario presionó CANCELAR
                    if(cancelPdfProcess) {
                        pdfWrapper.style.overflow = "auto";
                        ocultarLoader();
                        mostrarCustomAlert("La generación de PDF fue detenida.");
                        consolidarCambios(true); // Resetea la vista
                        return;
                    }

                    const pageEl = paginas[i];
                    window.scrollTo(0, pageEl.offsetTop); 
                    pageEl.style.boxShadow = "none"; pageEl.style.border = "none"; pageEl.style.marginBottom = "0";

                    const canvas = await html2canvas(pageEl, { scale: 1.5, useCORS: true, logging: false, scrollY: 0, scrollX: 0 });
                    const imgData = canvas.toDataURL('image/jpeg', 0.95);

                    if (i > 0) pdf.addPage();
                    pdf.addImage(imgData, 'JPEG', 0, 0, 794, 1122);

                    pageEl.style.boxShadow = ""; pageEl.style.border = ""; pageEl.style.marginBottom = "";
                    if(progressBar) {
                        let porcentaje = Math.round(((i + 1) / paginas.length) * 100);
                        progressBar.style.width = porcentaje + '%';
                    }
                }

                pdf.save(nombreArchivo);
                pdfWrapper.style.overflow = "auto";
                ocultarLoader();
            } catch (err) {
                console.error(err);
                pdfWrapper.style.overflow = "auto";
                mostrarCustomAlert("Hubo un error al generar el PDF. Usa la Impresión Rápida.");
                ocultarLoader();
            }
        }, 1000); 
    });
}
