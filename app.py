<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Control y Seguimiento Operacional — Línea Viva</title>
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome CDN -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- Google Fonts Inter -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <!-- SheetJS CDN para lectura de Excel -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <!-- Chart.js CDN -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

    <script>
        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        brand: {
                            50: '#f0fdf4',
                            500: '#16a34a',
                            600: '#15803d',
                            700: '#166534',
                            900: '#14532d'
                        },
                        celsia: {
                            500: '#ff6600',
                            600: '#e65c00'
                        }
                    },
                    fontFamily: {
                        sans: ['Inter', 'sans-serif'],
                    }
                }
            }
        }
    </script>
    <style>
        body { font-family: 'Inter', sans-serif; background-color: #f8fafc; }
        .tab-btn.active {
            border-bottom: 3px solid #ff6600;
            color: #ff6600;
            font-weight: 700;
        }
        .custom-scroll::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scroll::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
    </style>
</head>
<body class="text-slate-800 antialiased min-h-screen flex flex-col">

    <!-- HEADER / BARRA SUPERIOR CORPORATIVA -->
    <header class="bg-slate-900 text-white shadow-lg sticky top-0 z-50 border-b border-slate-800">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
            <div class="flex items-center space-x-3">
                <div class="bg-amber-500 p-2 rounded-lg text-slate-900 font-bold">
                    <i class="fa-solid fa-bolt text-xl"></i>
                </div>
                <div>
                    <h1 class="text-lg font-bold tracking-tight leading-tight">Control Operacional — Línea Viva</h1>
                    <p class="text-xs text-slate-400">Celsia & Proyectos de Ingeniería S.A.</p>
                </div>
            </div>
            
            <div class="flex items-center space-x-4">
                <label for="excelFileInput" class="cursor-pointer bg-amber-500 hover:bg-amber-600 text-slate-950 font-semibold text-xs px-3.5 py-2 rounded-lg transition-all flex items-center shadow-md">
                    <i class="fa-solid fa-file-excel mr-2 text-sm"></i> Cargar / Actualizar Excel
                </label>
                <input type="file" id="excelFileInput" accept=".xlsx, .xls" class="hidden" />
                <span id="fileNameDisplay" class="text-xs text-slate-300 italic hidden sm:inline">Esperando archivo...</span>
            </div>
        </div>
    </header>

    <!-- LAYOUT PRINCIPAL -->
    <div class="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 grid grid-cols-1 lg:grid-cols-12 gap-6">

        <!-- PANEL LATERAL DE FILTROS -->
        <aside class="lg:col-span-3 bg-white rounded-xl shadow-sm border border-slate-200 p-5 space-y-5 h-fit sticky top-20">
            <div class="flex items-center justify-between border-b border-slate-100 pb-3">
                <h2 class="text-sm font-bold text-slate-900 uppercase tracking-wider flex items-center">
                    <i class="fa-solid fa-filter text-amber-500 mr-2"></i> Filtros de Control
                </h2>
                <button onclick="resetFilters()" class="text-xs text-slate-500 hover:text-amber-600 underline">Limpiar</button>
            </div>

            <!-- Selector Exclusión de Pesadas -->
            <div>
                <label class="block text-xs font-semibold text-slate-700 mb-2">Segmentación de Cuadrillas</label>
                <div class="space-y-1.5">
                    <label class="flex items-center text-xs text-slate-700 cursor-pointer p-2 rounded hover:bg-slate-50 border border-slate-200">
                        <input type="radio" name="cuadrillaType" value="LINV" checked onchange="applyFilters()" class="text-amber-500 focus:ring-amber-500">
                        <span class="ml-2 font-medium">Solo Línea Viva (LINV)</span>
                    </label>
                    <label class="flex items-center text-xs text-slate-700 cursor-pointer p-2 rounded hover:bg-slate-50 border border-slate-200">
                        <input type="radio" name="cuadrillaType" value="ALL" onchange="applyFilters()" class="text-amber-500 focus:ring-amber-500">
                        <span class="ml-2 font-medium">Todas (Incluye CUAD y CR)</span>
                    </label>
                </div>
            </div>

            <!-- Filtro Mes -->
            <div>
                <label class="block text-xs font-semibold text-slate-700 mb-1">Mes de Operación</label>
                <select id="filterMes" onchange="applyFilters()" class="w-full bg-slate-50 border border-slate-300 rounded-lg text-xs p-2.5 focus:ring-2 focus:ring-amber-500 focus:outline-none">
                    <option value="ALL">Todos los Meses</option>
                </select>
            </div>

            <!-- Filtro Cuadrilla -->
            <div>
                <label class="block text-xs font-semibold text-slate-700 mb-1">Cuadrilla (Grupo)</label>
                <select id="filterGrupo" onchange="applyFilters()" class="w-full bg-slate-50 border border-slate-300 rounded-lg text-xs p-2.5 focus:ring-2 focus:ring-amber-500 focus:outline-none">
                    <option value="ALL">Todas las Cuadrillas</option>
                </select>
            </div>

            <!-- Filtro Jefe de Cuadrilla -->
            <div>
                <label class="block text-xs font-semibold text-slate-700 mb-1">Jefe de Cuadrilla</label>
                <select id="filterJefe" onchange="applyFilters()" class="w-full bg-slate-50 border border-slate-300 rounded-lg text-xs p-2.5 focus:ring-2 focus:ring-amber-500 focus:outline-none">
                    <option value="ALL">Todos los Jefes</option>
                </select>
            </div>

            <!-- NUEVO: Filtro por Estado -->
            <div>
                <label class="block text-xs font-semibold text-slate-700 mb-1">Estado de Actividad</label>
                <select id="filterEstado" onchange="applyFilters()" class="w-full bg-slate-50 border border-slate-300 rounded-lg text-xs p-2.5 focus:ring-2 focus:ring-amber-500 focus:outline-none">
                    <option value="ALL">Todos los Estados</option>
                </select>
            </div>

            <!-- NUEVO: Filtro por Diligenciada en Sisproing -->
            <div>
                <label class="block text-xs font-semibold text-slate-700 mb-1">Diligenciada en Sisproing</label>
                <select id="filterSisproing" onchange="applyFilters()" class="w-full bg-slate-50 border border-slate-300 rounded-lg text-xs p-2.5 focus:ring-2 focus:ring-amber-500 focus:outline-none">
                    <option value="ALL">Todos los Estados Sisproing</option>
                </select>
            </div>

            <div class="pt-2 text-center text-xs text-slate-400">
                <span id="recordCountDisplay">0 registros visualizados</span>
            </div>
        </aside>

        <!-- ÁREA DE CONTENIDO Y DASHBOARD -->
        <main class="lg:col-span-9 space-y-6">

            <!-- TARJETAS DE KPIS PRINCIPALES -->
            <div class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
                <div class="bg-white p-3.5 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-slate-700">
                    <span class="text-[11px] font-bold text-slate-500 uppercase block">Total Actividades</span>
                    <span id="kpiTotal" class="text-xl font-extrabold text-slate-900 mt-1 block">0</span>
                </div>
                <div class="bg-white p-3.5 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-emerald-500">
                    <span class="text-[11px] font-bold text-emerald-600 uppercase block">Finalizadas</span>
                    <span id="kpiFinalizados" class="text-xl font-extrabold text-slate-900 mt-1 block">0</span>
                </div>
                <div class="bg-white p-3.5 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-amber-500">
                    <span class="text-[11px] font-bold text-amber-600 uppercase block">En Ejecución</span>
                    <span id="kpiEjecucion" class="text-xl font-extrabold text-slate-900 mt-1 block">0</span>
                </div>
                <div class="bg-white p-3.5 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-purple-500">
                    <span class="text-[11px] font-bold text-purple-600 uppercase block">Fin. SIN Diligenciar</span>
                    <span id="kpiFinSinDil" class="text-xl font-extrabold text-slate-900 mt-1 block">0</span>
                </div>
                <div class="bg-white p-3.5 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-indigo-500">
                    <span class="text-[11px] font-bold text-indigo-600 uppercase block">Re-Programadas</span>
                    <span id="kpiReprogramados" class="text-xl font-extrabold text-slate-900 mt-1 block">0</span>
                </div>
                <div class="bg-white p-3.5 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-rose-500">
                    <span class="text-[11px] font-bold text-rose-600 uppercase block">Canceladas</span>
                    <span id="kpiCancelados" class="text-xl font-extrabold text-slate-900 mt-1 block">0</span>
                </div>
            </div>

            <!-- PESTAÑAS DE NAVEGACIÓN -->
            <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                <div class="flex border-b border-slate-200 overflow-x-auto custom-scroll bg-slate-50">
                    <button onclick="switchTab('auditoria')" id="tab-auditoria" class="tab-btn active px-5 py-3.5 text-xs font-semibold text-slate-600 hover:text-slate-900 flex items-center whitespace-nowrap">
                        <i class="fa-solid fa-triangle-exclamation mr-2 text-amber-500"></i> Auditoría Crítica
                    </button>
                    <button onclick="switchTab('finalizadas')" id="tab-finalizadas" class="tab-btn px-5 py-3.5 text-xs font-semibold text-slate-600 hover:text-slate-900 flex items-center whitespace-nowrap">
                        <i class="fa-solid fa-circle-check mr-2 text-emerald-500"></i> Actividades Finalizadas
                    </button>
                    <button onclick="switchTab('finesSemana')" id="tab-finesSemana" class="tab-btn px-5 py-3.5 text-xs font-semibold text-slate-600 hover:text-slate-900 flex items-center whitespace-nowrap">
                        <i class="fa-solid fa-calendar-week mr-2 text-indigo-500"></i> Fines de Semana (Sáb/Dom)
                    </button>
                    <button onclick="switchTab('graficos')" id="tab-graficos" class="tab-btn px-5 py-3.5 text-xs font-semibold text-slate-600 hover:text-slate-900 flex items-center whitespace-nowrap">
                        <i class="fa-solid fa-chart-pie mr-2 text-sky-500"></i> Análisis Estadístico
                    </button>
                    <button onclick="switchTab('tabla')" id="tab-tabla" class="tab-btn px-5 py-3.5 text-xs font-semibold text-slate-600 hover:text-slate-900 flex items-center whitespace-nowrap">
                        <i class="fa-solid fa-table-list mr-2 text-slate-600"></i> Base de Datos Completa
                    </button>
                </div>

                <!-- CONTENIDO DE PESTAÑA 1: AUDITORÍA CRÍTICA -->
                <div id="content-auditoria" class="p-5 space-y-6">
                    <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                        
                        <!-- 1. En Ejecución con días transcurridos -->
                        <div class="bg-amber-50/50 border border-amber-200 rounded-xl p-4">
                            <div class="flex items-center justify-between mb-3">
                                <h3 class="text-xs font-bold text-amber-900 uppercase flex items-center">
                                    <i class="fa-solid fa-clock text-amber-600 mr-2"></i> Actividades EN EJECUCIÓN
                                </h3>
                                <span id="badgeEjecucion" class="bg-amber-200 text-amber-900 text-[10px] font-extrabold px-2 py-0.5 rounded-full">0</span>
                            </div>
                            <div class="overflow-x-auto max-h-60 custom-scroll">
                                <table class="w-full text-left text-xs text-slate-700">
                                    <thead class="bg-amber-100 text-amber-900 sticky top-0">
                                        <tr>
                                            <th class="p-2">Fecha</th>
                                            <th class="p-2 text-center">Días Activos</th>
                                            <th class="p-2">Cuadrilla</th>
                                            <th class="p-2">OM / Aviso</th>
                                            <th class="p-2">Sisproing</th>
                                        </tr>
                                    </thead>
                                    <tbody id="tableEjecucionBody" class="divide-y divide-amber-100">
                                        <!-- Filas dinámicas -->
                                    </tbody>
                                </table>
                            </div>
                        </div>

                        <!-- 2. Finalizado SIN Diligenciar en Sisproing -->
                        <div class="bg-purple-50/50 border border-purple-200 rounded-xl p-4">
                            <div class="flex items-center justify-between mb-3">
                                <h3 class="text-xs font-bold text-purple-900 uppercase flex items-center">
                                    <i class="fa-solid fa-file-circle-xmark text-purple-600 mr-2"></i> Finalizado SIN Diligenciar
                                </h3>
                                <span id="badgeFinSinDil" class="bg-purple-200 text-purple-900 text-[10px] font-extrabold px-2 py-0.5 rounded-full">0</span>
                            </div>
                            <div class="overflow-x-auto max-h-60 custom-scroll">
                                <table class="w-full text-left text-xs text-slate-700">
                                    <thead class="bg-purple-100 text-purple-900 sticky top-0">
                                        <tr>
                                            <th class="p-2">Fecha</th>
                                            <th class="p-2">Cuadrilla</th>
                                            <th class="p-2">Jefe Cuadrilla</th>
                                            <th class="p-2">OM / Aviso</th>
                                            <th class="p-2">Estado Sisproing</th>
                                        </tr>
                                    </thead>
                                    <tbody id="tableFinSinDilBody" class="divide-y divide-purple-100">
                                        <!-- Filas dinámicas -->
                                    </tbody>
                                </table>
                            </div>
                        </div>

                    </div>

                    <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">

                        <!-- 3. Actividades Re-Programadas -->
                        <div class="bg-indigo-50/50 border border-indigo-200 rounded-xl p-4">
                            <div class="flex items-center justify-between mb-3">
                                <h3 class="text-xs font-bold text-indigo-900 uppercase flex items-center">
                                    <i class="fa-solid fa-calendar-xmark text-indigo-600 mr-2"></i> Actividades Re-Programadas
                                </h3>
                                <span id="badgeReprogramados" class="bg-indigo-200 text-indigo-900 text-[10px] font-extrabold px-2 py-0.5 rounded-full">0</span>
                            </div>
                            <div class="overflow-x-auto max-h-60 custom-scroll">
                                <table class="w-full text-left text-xs text-slate-700">
                                    <thead class="bg-indigo-100 text-indigo-900 sticky top-0">
                                        <tr>
                                            <th class="p-2">Fecha</th>
                                            <th class="p-2">Cuadrilla</th>
                                            <th class="p-2">OM / Aviso</th>
                                            <th class="p-2">Observaciones</th>
                                        </tr>
                                    </thead>
                                    <tbody id="tableReprogramadosBody" class="divide-y divide-indigo-100">
                                        <!-- Filas dinámicas -->
                                    </tbody>
                                </table>
                            </div>
                        </div>

                        <!-- 4. Actividades Canceladas (Ajustadas con Sisproing) -->
                        <div class="bg-rose-50/50 border border-rose-200 rounded-xl p-4">
                            <div class="flex items-center justify-between mb-3">
                                <h3 class="text-xs font-bold text-rose-900 uppercase flex items-center">
                                    <i class="fa-solid fa-ban text-rose-600 mr-2"></i> Actividades Canceladas (Total)
                                </h3>
                                <span id="badgeCancelados" class="bg-rose-200 text-rose-900 text-[10px] font-extrabold px-2 py-0.5 rounded-full">0</span>
                            </div>
                            <div class="overflow-x-auto max-h-60 custom-scroll">
                                <table class="w-full text-left text-xs text-slate-700">
                                    <thead class="bg-rose-100 text-rose-900 sticky top-0">
                                        <tr>
                                            <th class="p-2">Fecha</th>
                                            <th class="p-2">Cuadrilla</th>
                                            <th class="p-2">OM / Aviso</th>
                                            <th class="p-2">Origen Cancelación</th>
                                        </tr>
                                    </thead>
                                    <tbody id="tableCanceladosBody" class="divide-y divide-rose-100">
                                        <!-- Filas dinámicas -->
                                    </tbody>
                                </table>
                            </div>
                        </div>

                    </div>
                </div>

                <!-- CONTENIDO DE PESTAÑA 2: ACTIVIDADES FINALIZADAS (NUEVO) -->
                <div id="content-finalizadas" class="p-5 hidden space-y-4">
                    <div class="flex items-center justify-between border-b border-slate-100 pb-3">
                        <div>
                            <h3 class="text-sm font-bold text-slate-900 uppercase">Panel de Control de Obras Finalizadas</h3>
                            <p class="text-xs text-slate-500">Muestra todas las actividades ejecutadas y su estado de registro en Sisproing.</p>
                        </div>
                        <button onclick="exportTableToCSV('tableFinalizadasFull', 'finalizadas_linea_viva.csv')" class="bg-emerald-600 hover:bg-emerald-700 text-white font-semibold text-xs px-3 py-1.5 rounded-lg">
                            <i class="fa-solid fa-download mr-1"></i> Exportar Finalizadas
                        </button>
                    </div>

                    <div class="overflow-x-auto max-h-96 custom-scroll">
                        <table id="tableFinalizadasFull" class="w-full text-left text-xs text-slate-700">
                            <thead class="bg-slate-100 text-slate-800 sticky top-0 border-b border-slate-200">
                                <tr>
                                    <th class="p-2.5">Mes</th>
                                    <th class="p-2.5">Fecha</th>
                                    <th class="p-2.5">Cuadrilla</th>
                                    <th class="p-2.5">Jefe de Cuadrilla</th>
                                    <th class="p-2.5">OM</th>
                                    <th class="p-2.5">Aviso (VP)</th>
                                    <th class="p-2.5">Sisproing</th>
                                    <th class="p-2.5">Observaciones</th>
                                </tr>
                            </thead>
                            <tbody id="tableFinalizadasBody" class="divide-y divide-slate-100">
                                <!-- Filas dinámicas -->
                            </tbody>
                        </table>
                    </div>
                </div>

                <!-- CONTENIDO DE PESTAÑA 3: FINES DE SEMANA -->
                <div id="content-finesSemana" class="p-5 hidden space-y-4">
                    <div class="border-b border-slate-100 pb-3">
                        <h3 class="text-sm font-bold text-slate-900 uppercase">Disponibilidad Operativa en Fines de Semana</h3>
                        <p class="text-xs text-slate-500">Auditoría de turnos de Sábados y Domingos (Validación de cuadrilla de turno disponible vs descanso).</p>
                    </div>

                    <div class="overflow-x-auto max-h-96 custom-scroll">
                        <table class="w-full text-left text-xs text-slate-700">
                            <thead class="bg-indigo-50 text-indigo-900 sticky top-0 border-b border-indigo-200">
                                <tr>
                                    <th class="p-2.5">Fecha (Sáb/Dom)</th>
                                    <th class="p-2.5">Mes</th>
                                    <th class="p-2.5">Cuadrilla</th>
                                    <th class="p-2.5">Jefe Cuadrilla</th>
                                    <th class="p-2.5">Estado Operativo</th>
                                    <th class="p-2.5">Sisproing</th>
                                    <th class="p-2.5">OM / Aviso</th>
                                </tr>
                            </thead>
                            <tbody id="tableFinesSemanaBody" class="divide-y divide-slate-100">
                                <!-- Filas dinámicas -->
                            </tbody>
                        </table>
                    </div>
                </div>

                <!-- CONTENIDO DE PESTAÑA 4: ANÁLISIS ESTADÍSTICO -->
                <div id="content-graficos" class="p-5 hidden space-y-6">
                    <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                        <div class="bg-slate-50 p-4 rounded-xl border border-slate-200">
                            <h4 class="text-xs font-bold text-slate-800 uppercase mb-3">Distribución por Mes y Estado</h4>
                            <div class="h-64">
                                <canvas id="chartMesEstado"></canvas>
                            </div>
                        </div>
                        <div class="bg-slate-50 p-4 rounded-xl border border-slate-200">
                            <h4 class="text-xs font-bold text-slate-800 uppercase mb-3">Estado de Diligenciamiento Sisproing</h4>
                            <div class="h-64 flex justify-center">
                                <canvas id="chartSisproing"></canvas>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- CONTENIDO DE PESTAÑA 5: BASE DE DATOS COMPLETA -->
                <div id="content-tabla" class="p-5 hidden space-y-4">
                    <div class="flex items-center justify-between border-b border-slate-100 pb-3">
                        <h3 class="text-sm font-bold text-slate-900 uppercase">Base de Datos Completa de Seguimiento</h3>
                        <button onclick="exportTableToCSV('tableFullData', 'seguimiento_completo.csv')" class="bg-slate-800 hover:bg-slate-900 text-white font-semibold text-xs px-3 py-1.5 rounded-lg">
                            <i class="fa-solid fa-file-csv mr-1"></i> Exportar CSV
                        </button>
                    </div>

                    <div class="overflow-x-auto max-h-96 custom-scroll">
                        <table id="tableFullData" class="w-full text-left text-xs text-slate-700">
                            <thead class="bg-slate-100 text-slate-800 sticky top-0 border-b border-slate-200">
                                <tr>
                                    <th class="p-2">Mes</th>
                                    <th class="p-2">Fecha</th>
                                    <th class="p-2">Cuadrilla</th>
                                    <th class="p-2">Jefe Cuadrilla</th>
                                    <th class="p-2">OM</th>
                                    <th class="p-2">Aviso (VP)</th>
                                    <th class="p-2">Estado</th>
                                    <th class="p-2">Sisproing</th>
                                    <th class="p-2">Observaciones</th>
                                </tr>
                            </thead>
                            <tbody id="tableFullDataBody" class="divide-y divide-slate-100">
                                <!-- Filas dinámicas -->
                            </tbody>
                        </table>
                    </div>
                </div>

            </div>

        </main>
    </div>

    <!-- LOGICA JAVASCRIPT DEL APLICATIVO -->
    <script>
        let rawData = [];
        let filteredData = [];
        let chart1Instance = null;
        let chart2Instance = null;

        // Carga de Excel local vía input
        document.getElementById('excelFileInput').addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;

            document.getElementById('fileNameDisplay').innerText = file.name;
            document.getElementById('fileNameDisplay').classList.remove('hidden');

            const reader = new FileReader();
            reader.onload = function(evt) {
                try {
                    const data = new Uint8Array(evt.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const json = XLSX.utils.sheet_to_json(worksheet, { raw: false });
                    
                    processRawData(json);
                } catch (err) {
                    alert("Error al leer el archivo Excel: " + err.message);
                }
            };
            reader.readAsArrayBuffer(file);
        });

        // Intentar autoreload de archivo por defecto si existe en la misma carpeta
        window.addEventListener('DOMContentLoaded', () => {
            fetch('COPIA PROGRAMADOR SEGUIMIENTO.xlsx')
                .then(res => {
                    if (res.ok) return res.arrayBuffer();
                    throw new Error("No encontrado");
                })
                .then(data => {
                    const workbook = XLSX.read(new Uint8Array(data), { type: 'array' });
                    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    const json = XLSX.utils.sheet_to_json(worksheet, { raw: false });
                    document.getElementById('fileNameDisplay').innerText = 'COPIA PROGRAMADOR SEGUIMIENTO.xlsx (Auto)';
                    document.getElementById('fileNameDisplay').classList.remove('hidden');
                    processRawData(json);
                })
                .catch(() => {
                    console.log("Cargue manual disponible mediante botón superior.");
                });
        });

        function processRawData(json) {
            const today = new Date();

            rawData = json.map((row, idx) => {
                const grupo = (row['GRUPO'] || 'SIN GRUPO').toString().trim();
                const estado = (row['Estado'] || 'Pendiente').toString().trim();
                const sisproing = (row['DILIGENCIADA EN SISPROING'] || 'SIN DILIGENCIAR').toString().trim();
                const mes = (row['MES'] || 'SIN MES').toString().trim().toUpperCase();
                const jefe = (row['JEFE DE CUADRILLA'] || 'NO ASIGNADO').toString().trim();
                const om = (row['OM'] || '').toString().trim();
                const aviso = (row['AVISO (VP)'] || '').toString().trim();
                const obs = (row['OBSERVACIONES'] || '').toString().trim();

                let fechaObj = null;
                let fechaStr = (row['FECHA'] || '').toString().trim();
                let diasEjecucion = 0;
                let esFinDeSemana = false;

                if (fechaStr) {
                    fechaObj = new Date(fechaStr);
                    if (!isNaN(fechaObj)) {
                        const diffTime = Math.abs(today - fechaObj);
                        diasEjecucion = Math.floor(diffTime / (1000 * 60 * 60 * 24));
                        const dayOfWeek = fechaObj.getDay(); // 0 = Domingo, 6 = Sábado
                        esFinDeSemana = (dayOfWeek === 0 || dayOfWeek === 6);
                    }
                }

                // Definición unificada de Cancelado
                const esCancelado = (estado.toLowerCase() === 'cancelado') || sisproing.toUpperCase().startsWith('CANCELAR');

                return {
                    id: idx,
                    grupo,
                    estado,
                    sisproing,
                    mes,
                    jefe,
                    om,
                    aviso,
                    obs,
                    fechaStr,
                    fechaObj,
                    diasEjecucion,
                    esFinDeSemana,
                    esCancelado
                };
            });

            populateFilterDropdowns();
            applyFilters();
        }

        function populateFilterDropdowns() {
            const selectMes = document.getElementById('filterMes');
            const selectGrupo = document.getElementById('filterGrupo');
            const selectJefe = document.getElementById('filterJefe');
            const selectEstado = document.getElementById('filterEstado');
            const selectSisproing = document.getElementById('filterSisproing');

            const meses = [...new Set(rawData.map(r => r.mes))].sort();
            const grupos = [...new Set(rawData.map(r => r.grupo))].sort();
            const jefes = [...new Set(rawData.map(r => r.jefe))].sort();
            const estados = [...new Set(rawData.map(r => r.estado))].sort();
            const sisproings = [...new Set(rawData.map(r => r.sisproing))].sort();

            fillSelectOptions(selectMes, meses);
            fillSelectOptions(selectGrupo, grupos);
            fillSelectOptions(selectJefe, jefes);
            fillSelectOptions(selectEstado, estados);
            fillSelectOptions(selectSisproing, sisproings);
        }

        function fillSelectOptions(selectElem, items) {
            const currentVal = selectElem.value;
            selectElem.innerHTML = `<option value="ALL">Todos</option>`;
            items.forEach(item => {
                if (item) {
                    const opt = document.createElement('option');
                    opt.value = item;
                    opt.textContent = item;
                    selectElem.appendChild(opt);
                }
            });
            selectElem.value = currentVal || 'ALL';
        }

        function applyFilters() {
            const typeCuadrilla = document.querySelector('input[name="cuadrillaType"]:checked').value;
            const mesVal = document.getElementById('filterMes').value;
            const grupoVal = document.getElementById('filterGrupo').value;
            const jefeVal = document.getElementById('filterJefe').value;
            const estadoVal = document.getElementById('filterEstado').value;
            const sisproingVal = document.getElementById('filterSisproing').value;

            filteredData = rawData.filter(r => {
                if (typeCuadrilla === 'LINV' && !r.grupo.startsWith('LINV')) return false;
                if (mesVal !== 'ALL' && r.mes !== mesVal) return false;
                if (grupoVal !== 'ALL' && r.grupo !== grupoVal) return false;
                if (jefeVal !== 'ALL' && r.jefe !== jefeVal) return false;
                if (estadoVal !== 'ALL' && r.estado !== estadoVal) return false;
                if (sisproingVal !== 'ALL' && r.sisproing !== sisproingVal) return false;
                return true;
            });

            document.getElementById('recordCountDisplay').innerText = `${filteredData.length} registros encontrados`;
            updateKPIs();
            renderAuditoriaTables();
            renderFinalizadasTable();
            renderFinesSemanaTable();
            renderFullTable();
            renderCharts();
        }

        function resetFilters() {
            document.querySelector('input[name="cuadrillaType"][value="LINV"]').checked = true;
            document.getElementById('filterMes').value = 'ALL';
            document.getElementById('filterGrupo').value = 'ALL';
            document.getElementById('filterJefe').value = 'ALL';
            document.getElementById('filterEstado').value = 'ALL';
            document.getElementById('filterSisproing').value = 'ALL';
            applyFilters();
        }

        function updateKPIs() {
            const total = filteredData.length;
            const finalizados = filteredData.filter(r => ['finalizado', 'ejecutado'].includes(r.estado.toLowerCase())).length;
            const ejecucion = filteredData.filter(r => r.estado.toLowerCase() === 'en ejecución').length;
            const finSinDil = filteredData.filter(r => ['finalizado', 'ejecutado'].includes(r.estado.toLowerCase()) && r.sisproing.toUpperCase() !== 'DILIGENCIADA').length;
            const reprogramados = filteredData.filter(r => ['re-programado', 'reprogramado'].includes(r.estado.toLowerCase())).length;
            const cancelados = filteredData.filter(r => r.esCancelado).length;

            document.getElementById('kpiTotal').innerText = total;
            document.getElementById('kpiFinalizados').innerText = finalizados;
            document.getElementById('kpiEjecucion').innerText = ejecucion;
            document.getElementById('kpiFinSinDil').innerText = finSinDil;
            document.getElementById('kpiReprogramados').innerText = reprogramados;
            document.getElementById('kpiCancelados').innerText = cancelados;

            document.getElementById('badgeEjecucion').innerText = ejecucion;
            document.getElementById('badgeFinSinDil').innerText = finSinDil;
            document.getElementById('badgeReprogramados').innerText = reprogramados;
            document.getElementById('badgeCancelados').innerText = cancelados;
        }

        function renderAuditoriaTables() {
            // Table 1: En Ejecucion
            const listEjecucion = filteredData.filter(r => r.estado.toLowerCase() === 'en ejecución');
            const tbodyEjec = document.getElementById('tableEjecucionBody');
            tbodyEjec.innerHTML = listEjecucion.map(r => `
                <tr class="hover:bg-amber-100/50">
                    <td class="p-2 whitespace-nowrap font-medium">${r.fechaStr}</td>
                    <td class="p-2 text-center font-bold text-amber-700 bg-amber-100 rounded">${r.diasEjecucion} días</td>
                    <td class="p-2 font-semibold">${r.grupo}</td>
                    <td class="p-2">${r.om || r.aviso}</td>
                    <td class="p-2">${r.sisproing}</td>
                </tr>
            `).join('') || '<tr><td colspan="5" class="p-3 text-center text-slate-400">Sin registros</td></tr>';

            // Table 2: Finalizado SIN Diligenciar
            const listFinSinDil = filteredData.filter(r => ['finalizado', 'ejecutado'].includes(r.estado.toLowerCase()) && r.sisproing.toUpperCase() !== 'DILIGENCIADA');
            const tbodyFinSinDil = document.getElementById('tableFinSinDilBody');
            tbodyFinSinDil.innerHTML = listFinSinDil.map(r => `
                <tr class="hover:bg-purple-100/50">
                    <td class="p-2 whitespace-nowrap">${r.fechaStr}</td>
                    <td class="p-2 font-semibold">${r.grupo}</td>
                    <td class="p-2">${r.jefe}</td>
                    <td class="p-2">${r.om || r.aviso}</td>
                    <td class="p-2"><span class="bg-purple-200 text-purple-900 font-bold px-1.5 py-0.5 rounded text-[10px]">${r.sisproing}</span></td>
                </tr>
            `).join('') || '<tr><td colspan="5" class="p-3 text-center text-slate-400">Sin registros</td></tr>';

            // Table 3: Re-programados
            const listReprog = filteredData.filter(r => ['re-programado', 'reprogramado'].includes(r.estado.toLowerCase()));
            const tbodyReprog = document.getElementById('tableReprogramadosBody');
            tbodyReprog.innerHTML = listReprog.map(r => `
                <tr class="hover:bg-indigo-100/50">
                    <td class="p-2 whitespace-nowrap">${r.fechaStr}</td>
                    <td class="p-2 font-semibold">${r.grupo}</td>
                    <td class="p-2">${r.om || r.aviso}</td>
                    <td class="p-2 text-slate-500 truncate max-w-xs">${r.obs || 'N/A'}</td>
                </tr>
            `).join('') || '<tr><td colspan="4" class="p-3 text-center text-slate-400">Sin registros</td></tr>';

            // Table 4: Cancelados
            const listCancel = filteredData.filter(r => r.esCancelado);
            const tbodyCanc = document.getElementById('tableCanceladosBody');
            tbodyCanc.innerHTML = listCancel.map(r => `
                <tr class="hover:bg-rose-100/50">
                    <td class="p-2 whitespace-nowrap">${r.fechaStr}</td>
                    <td class="p-2 font-semibold">${r.grupo}</td>
                    <td class="p-2">${r.om || r.aviso}</td>
                    <td class="p-2"><span class="bg-rose-200 text-rose-900 font-bold px-1.5 py-0.5 rounded text-[10px]">${r.sisproing.startsWith('CANCELAR') ? 'Sisproing (' + r.sisproing + ')' : 'Estado (Cancelado)'}</span></td>
                </tr>
            `).join('') || '<tr><td colspan="4" class="p-3 text-center text-slate-400">Sin registros</td></tr>';
        }

        function renderFinalizadasTable() {
            const listFinalizadas = filteredData.filter(r => ['finalizado', 'ejecutado'].includes(r.estado.toLowerCase()));
            const tbody = document.getElementById('tableFinalizadasBody');
            tbody.innerHTML = listFinalizadas.map(r => `
                <tr class="hover:bg-slate-50">
                    <td class="p-2.5 font-semibold text-slate-600">${r.mes}</td>
                    <td class="p-2.5 whitespace-nowrap">${r.fechaStr}</td>
                    <td class="p-2.5 font-bold text-amber-600">${r.grupo}</td>
                    <td class="p-2.5">${r.jefe}</td>
                    <td class="p-2.5 font-mono">${r.om}</td>
                    <td class="p-2.5 font-mono">${r.aviso}</td>
                    <td class="p-2.5">
                        <span class="px-2 py-0.5 rounded-full text-[10px] font-bold ${r.sisproing.toUpperCase() === 'DILIGENCIADA' ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'}">
                            ${r.sisproing}
                        </span>
                    </td>
                    <td class="p-2.5 text-slate-500">${r.obs || '-'}</td>
                </tr>
            `).join('') || '<tr><td colspan="8" class="p-4 text-center text-slate-400">No hay actividades finalizadas con los filtros aplicados.</td></tr>';
        }

        function renderFinesSemanaTable() {
            const listFds = filteredData.filter(r => r.esFinDeSemana);
            const tbody = document.getElementById('tableFinesSemanaBody');
            tbody.innerHTML = listFds.map(r => `
                <tr class="hover:bg-slate-50">
                    <td class="p-2.5 font-semibold text-indigo-700 whitespace-nowrap">${r.fechaStr}</td>
                    <td class="p-2.5">${r.mes}</td>
                    <td class="p-2.5 font-bold">${r.grupo}</td>
                    <td class="p-2.5">${r.jefe}</td>
                    <td class="p-2.5">
                        <span class="px-2 py-0.5 rounded text-[10px] font-bold ${r.estado.toLowerCase() === 'no labora' ? 'bg-slate-200 text-slate-700' : 'bg-amber-100 text-amber-900'}">
                            ${r.estado}
                        </span>
                    </td>
                    <td class="p-2.5">${r.sisproing}</td>
                    <td class="p-2.5 font-mono">${r.om || r.aviso}</td>
                </tr>
            `).join('') || '<tr><td colspan="7" class="p-4 text-center text-slate-400">No hay datos de fines de semana con los filtros aplicados.</td></tr>';
        }

        function renderFullTable() {
            const tbody = document.getElementById('tableFullDataBody');
            tbody.innerHTML = filteredData.map(r => `
                <tr class="hover:bg-slate-50">
                    <td class="p-2">${r.mes}</td>
                    <td class="p-2 whitespace-nowrap">${r.fechaStr}</td>
                    <td class="p-2 font-semibold">${r.grupo}</td>
                    <td class="p-2">${r.jefe}</td>
                    <td class="p-2 font-mono">${r.om}</td>
                    <td class="p-2 font-mono">${r.aviso}</td>
                    <td class="p-2 font-medium">${r.estado}</td>
                    <td class="p-2">${r.sisproing}</td>
                    <td class="p-2 text-slate-500">${r.obs}</td>
                </tr>
            `).join('') || '<tr><td colspan="9" class="p-4 text-center text-slate-400">Sin registros para mostrar.</td></tr>';
        }

        function renderCharts() {
            // Chart 1: Mes vs Estado
            const ctx1 = document.getElementById('chartMesEstado').getContext('2d');
            const mesesList = [...new Set(filteredData.map(r => r.mes))];
            
            const countFinalizados = mesesList.map(m => filteredData.filter(r => r.mes === m && ['finalizado','ejecutado'].includes(r.estado.toLowerCase())).length);
            const countEjecucion = mesesList.map(m => filteredData.filter(r => r.mes === m && r.estado.toLowerCase() === 'en ejecución').length);
            const countOtros = mesesList.map(m => filteredData.filter(r => r.mes === m && !['finalizado','ejecutado','en ejecución'].includes(r.estado.toLowerCase())).length);

            if (chart1Instance) chart1Instance.destroy();
            chart1Instance = new Chart(ctx1, {
                type: 'bar',
                data: {
                    labels: mesesList,
                    datasets: [
                        { label: 'Finalizado', data: countFinalizados, backgroundColor: '#10b981' },
                        { label: 'En Ejecución', data: countEjecucion, backgroundColor: '#f59e0b' },
                        { label: 'Otros / Pendientes', data: countOtros, backgroundColor: '#64748b' }
                    ]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { position: 'bottom' } }
                }
            });

            // Chart 2: Sisproing Pie
            const ctx2 = document.getElementById('chartSisproing').getContext('2d');
            const sisTypes = [...new Set(filteredData.map(r => r.sisproing))];
            const sisCounts = sisTypes.map(s => filteredData.filter(r => r.sisproing === s).length);

            if (chart2Instance) chart2Instance.destroy();
            chart2Instance = new Chart(ctx2, {
                type: 'doughnut',
                data: {
                    labels: sisTypes,
                    datasets: [{
                        data: sisCounts,
                        backgroundColor: ['#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#64748b']
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { position: 'bottom' } }
                }
            });
        }

        function switchTab(tabName) {
            document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
            document.getElementById(`tab-${tabName}`).classList.add('active');

            ['auditoria', 'finalizadas', 'finesSemana', 'graficos', 'tabla'].forEach(t => {
                document.getElementById(`content-${t}`).classList.add('hidden');
            });
            document.getElementById(`content-${tabName}`).classList.remove('hidden');
        }

        function exportTableToCSV(tableId, filename) {
            const table = document.getElementById(tableId);
            let csv = [];
            for (let i = 0; i < table.rows.length; i++) {
                let row = [], cols = table.rows[i].querySelectorAll("td, th");
                for (let j = 0; j < cols.length; j++) 
                    row.push('"' + cols[j].innerText.replace(/"/g, '""') + '"');
                csv.push(row.join(","));
            }
            const csvFile = new Blob([csv.join("\n")], { type: "text/csv;charset=utf-8;" });
            const downloadLink = document.createElement("a");
            downloadLink.download = filename;
            downloadLink.href = window.URL.createObjectURL(csvFile);
            downloadLink.style.display = "none";
            document.body.appendChild(downloadLink);
            downloadLink.click();
        }
    </script>
</body>
</html>
