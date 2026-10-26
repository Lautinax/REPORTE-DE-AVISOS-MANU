<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Reporte de Gestión de Reclamos - MAINCAL</title>
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>
    <!-- SheetJS (XLSX / CSV parser) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- Google Fonts -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=Poppins:wght@500;600;700;800&family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">
    
    <style>
        body { font-family: 'Inter', sans-serif; }
        
        /* Estilos generales de la interfaz */
        .brand-blue { color: #0b2447; }
        .bg-brand-blue { background-color: #0b2447; }
        .border-brand-blue { border-color: #0b2447; }

        /* Estilos específicos para el Modo Presentación */
        #presentationModal {
            transition: opacity 0.25s ease-in-out;
        }

        .slide-card {
            width: 100%;
            height: 100%;
            max-width: 1280px;
            max-height: 720px;
            aspect-ratio: 16 / 9;
            box-shadow: 0 20px 40px -15px rgba(11, 36, 71, 0.25);
            display: flex;
            flex-direction: column;
            overflow: hidden;
            background-color: #ffffff;
            border-radius: 16px;
            position: relative;
        }

        /* Tabla estilizada para diapositivas */
        .slide-table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            font-size: 0.825rem;
        }

        .slide-table th {
            background-color: #0b2447;
            color: #ffffff;
            font-family: 'Poppins', sans-serif;
            font-weight: 600;
            padding: 10px 14px;
            text-align: left;
            position: sticky;
            top: 0;
            z-index: 10;
        }

        .slide-table td {
            padding: 8px 14px;
            border-bottom: 1px solid #e2e8f0;
            color: #1e293b;
        }

        .slide-table tbody tr:nth-child(even) td {
            background-color: #f0f9ff;
        }

        .slide-table tbody tr:hover td {
            background-color: #e0f2fe;
        }

        .slide-table tbody tr.total-row td {
            background-color: #e0f2fe;
            color: #0b2447;
            font-weight: 800;
            border-top: 2px solid #0b2447;
            border-bottom: none;
        }

        /* Decoración de portada estilo corporativo MAINCAL */
        .slide-cover-accent {
            position: absolute;
            top: 0;
            left: 0;
            bottom: 0;
            width: 16px;
            background-color: #0b2447;
        }
        .slide-cover-accent-sub {
            position: absolute;
            top: 0;
            left: 16px;
            bottom: 0;
            width: 8px;
            background-color: #0284c7;
        }
        
        .slide-geo-pattern {
            position: absolute;
            right: -100px;
            bottom: -100px;
            width: 500px;
            height: 500px;
            border-radius: 50%;
            background: radial-gradient(circle, rgba(224,242,254,0.8) 0%, rgba(2,132,199,0.15) 60%, rgba(255,255,255,0) 70%);
            pointer-events: none;
        }
    </style>
</head>
<body class="bg-slate-50 text-slate-800 min-h-screen flex flex-col">

    <!-- Header Corporativo -->
    <header class="bg-[#0b2447] text-white shadow-md sticky top-0 z-40">
        <div class="max-w-7xl mx-auto px-6 py-4 flex justify-between items-center">
            <div class="flex items-center space-x-3">
                <div class="bg-sky-500 text-white px-3 py-1.5 rounded-lg font-bold text-sm tracking-wider">MAINCAL</div>
                <div>
                    <h1 class="font-bold text-base leading-tight">Reporte de Gestión de Reclamos</h1>
                    <p class="text-xs text-slate-300">Supply Chain y Servicio al Cliente</p>
                </div>
            </div>
            <div class="flex items-center gap-3">
                <button id="resetFileBtn" onclick="location.reload()" class="hidden bg-slate-800 hover:bg-slate-700 text-slate-200 px-3 py-1.5 rounded-lg text-xs font-medium transition border border-slate-700 cursor-pointer">
                    Cargar Otro Archivo
                </button>
                <button id="exportExcelBtn" onclick="exportToExcel()" class="hidden bg-emerald-600 hover:bg-emerald-500 text-white px-4 py-2 rounded-lg text-xs font-semibold transition flex items-center gap-2 cursor-pointer shadow">
                    <svg xmlns="http://www.w3.org/2000/svg" class="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/></svg>
                    Exportar Excel
                </button>
                <button id="presentationBtn" onclick="openPresentation()" class="hidden bg-sky-500 hover:bg-sky-400 text-white px-4 py-2 rounded-lg text-xs font-bold transition flex items-center gap-2 cursor-pointer shadow-md transform hover:-translate-y-0.5">
                    <svg xmlns="http://www.w3.org/2000/svg" class="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 12l3-3 3 3 4-4M8 21l4-4 4 4M3 4h18M4 4h16v12H4z"/></svg>
                    Ver Presentación
                </button>
            </div>
        </div>
    </header>

    <!-- Sección de Carga de Archivo -->
    <main class="flex-grow max-w-7xl w-full mx-auto p-6">
        <div id="uploadCard" class="bg-white rounded-2xl shadow-sm border border-slate-200 p-12 text-center max-w-xl mx-auto my-16">
            <div class="w-16 h-16 bg-sky-50 text-sky-600 rounded-2xl flex items-center justify-center mx-auto mb-4 text-2xl shadow-inner">📊</div>
            <h2 class="text-xl font-bold text-slate-900 mb-2">Cargar Archivo Excel o CSV</h2>
            <p class="text-sm text-slate-500 mb-6">Selecciona o arrastra el reporte general de avisos para generar automáticamente el análisis interactivo.</p>
            
            <label class="border-2 border-dashed border-slate-300 hover:border-sky-500 rounded-xl p-8 block cursor-pointer transition bg-slate-50/50 hover:bg-sky-50/30">
                <span class="text-sm font-semibold text-sky-600 block mb-1">Examinar equipo (.xlsx, .xls, .csv)</span>
                <span class="text-xs text-slate-400">Los datos se procesan 100% en tu navegador de forma segura.</span>
                <input type="file" id="excelFile" accept=".xlsx, .xls, .csv" class="hidden">
            </label>
            <div id="loadingMsg" class="hidden mt-4 text-sm font-medium text-sky-600 animate-pulse">Procesando información...</div>
        </div>

        <!-- Contenedor del Reporte Web Tradicional (Oculto hasta cargar datos) -->
        <div id="reportContainer" class="hidden space-y-8">
            
            <!-- Navegación interna del Reporte -->
            <nav class="bg-white rounded-xl shadow-sm border border-slate-200 px-6 py-3 flex flex-wrap gap-4 text-xs font-semibold text-slate-600 items-center justify-between">
                <div class="flex flex-wrap gap-3">
                    <a href="#resumen" class="hover:text-sky-600 transition">Resumen Ejecutivo</a>
                    <a href="#clientes" class="hover:text-sky-600 transition">Por Cliente</a>
                    <a href="#tipos" class="hover:text-sky-600 transition">Por Tipo</a>
                    <a href="#evolucion" class="hover:text-sky-600 transition">Evolución Mensual</a>
                    <a href="#gestion" class="hover:text-sky-600 transition">Estado de Gestión</a>
                    <a href="#detalle" class="hover:text-sky-600 transition">Detalle Completo</a>
                </div>
            </nav>

            <!-- Filtros Interactivos -->
            <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 space-y-4">
                <div class="flex justify-between items-center">
                    <h3 class="text-sm font-bold text-slate-900 uppercase tracking-wider">Filtros Interactivos</h3>
                    <button onclick="clearFilters()" class="text-xs font-semibold text-sky-600 hover:text-sky-700 bg-sky-50 px-3 py-1.5 rounded-lg transition">Limpiar Filtros</button>
                </div>
                <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 lg:grid-cols-7 gap-3">
                    <input type="text" id="globalSearch" placeholder="Buscar general..." class="bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 text-xs text-slate-800 focus:outline-none focus:border-sky-500" oninput="debounceApplyFilters()">
                    <select id="filterClient" class="bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 text-xs text-slate-800 focus:outline-none focus:border-sky-500" onchange="applyFilters()"><option value="">Todos los clientes</option></select>
                    <select id="filterType" class="bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 text-xs text-slate-800 focus:outline-none focus:border-sky-500" onchange="applyFilters()"><option value="">Todos los tipos</option></select>
                    <select id="filterStatus" class="bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 text-xs text-slate-800 focus:outline-none focus:border-sky-500" onchange="applyFilters()">
                        <option value="">Todos los estados</option>
                        <option value="CERRADO">Cerrado</option>
                        <option value="ABIERTO / PENDIENTE">Abierto / Pendiente</option>
                    </select>
                    <select id="filterCreator" class="bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 text-xs text-slate-800 focus:outline-none focus:border-sky-500" onchange="applyFilters()"><option value="">Todos los creadores</option></select>
                    <select id="filterYear" class="bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 text-xs text-slate-800 focus:outline-none focus:border-sky-500" onchange="applyFilters()"><option value="">Todos los años</option></select>
                    <select id="filterMonth" class="bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 text-xs text-slate-800 focus:outline-none focus:border-sky-500" onchange="applyFilters()"><option value="">Todos los meses</option></select>
                </div>
            </div>

            <!-- PORTADA / RESUMEN EJECUTIVO -->
            <section id="resumen" class="bg-white rounded-2xl shadow-sm border border-slate-200 p-8">
                <div class="border-b border-slate-100 pb-6 mb-6 flex justify-between items-start">
                    <div>
                        <span class="text-xs font-bold uppercase tracking-widest text-sky-600 bg-sky-50 px-3 py-1 rounded-full">Informe Anual</span>
                        <h2 class="text-3xl font-black text-slate-900 mt-2">REPORTE DE GESTIÓN DE RECLAMOS</h2>
                        <p id="periodSubtitle" class="text-sm font-medium text-slate-500 mt-1">Período analizado: Cargando...</p>
                    </div>
                    <div class="text-right">
                        <p class="text-xs font-semibold text-slate-400">MAINCAL S.A.</p>
                        <p class="text-xs text-slate-500 mt-0.5">Supply Chain y Servicio al Cliente</p>
                    </div>
                </div>

                <!-- Tarjetas de Resumen General -->
                <div class="grid grid-cols-2 md:grid-cols-4 gap-4 mb-8">
                    <div class="bg-slate-50 border border-slate-200/80 rounded-xl p-4 flex flex-col justify-between">
                        <span class="text-xs font-semibold text-slate-500 uppercase">Total Reclamos</span>
                        <span id="cardTotal" class="text-3xl font-black text-slate-900 mt-2">0</span>
                    </div>
                    <div class="bg-slate-50 border border-slate-200/80 rounded-xl p-4 flex flex-col justify-between">
                        <span class="text-xs font-semibold text-slate-500 uppercase">Cerrados</span>
                        <div class="flex items-baseline gap-2 mt-2">
                            <span id="cardClosed" class="text-3xl font-black text-emerald-600">0</span>
                            <span id="cardClosedPct" class="text-xs font-bold text-emerald-500">0%</span>
                        </div>
                    </div>
                    <div class="bg-slate-50 border border-slate-200/80 rounded-xl p-4 flex flex-col justify-between">
                        <span class="text-xs font-semibold text-slate-500 uppercase">Abiertos / Pendientes</span>
                        <span id="cardPending" class="text-3xl font-black text-amber-500 mt-2">0</span>
                    </div>
                    <div class="bg-slate-50 border border-slate-200/80 rounded-xl p-4 flex flex-col justify-between">
                        <span class="text-xs font-semibold text-slate-500 uppercase">Cantidad Reclamada</span>
                        <span id="cardQty" class="text-3xl font-black text-slate-900 mt-2">0</span>
                        <span class="text-xs text-slate-400 mt-0.5">pares totales</span>
                    </div>
                </div>

                <!-- Indicadores Secundarios -->
                <div class="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs font-medium text-slate-600 bg-sky-50/50 border border-sky-100 rounded-xl p-4">
                    <div>📅 <span class="font-bold">Fechas:</span> <span id="infoDates">-</span></div>
                    <div>🏢 <span class="font-bold">Clientes Involucrados:</span> <span id="infoClientsCount">0</span></div>
                    <div>👥 <span class="font-bold">Creadores de Avisos:</span> <span id="infoCreatorsCount">0</span></div>
                </div>
            </section>

            <!-- REPORTE POR CLIENTE -->
            <section id="clientes" class="bg-white rounded-2xl shadow-sm border border-slate-200 p-8">
                <div class="flex justify-between items-center mb-6">
                    <div>
                        <h3 class="text-xl font-extrabold text-slate-900">Reporte por Cliente</h3>
                        <p class="text-xs text-slate-500">Agrupado por cliente. Haz clic en un cliente para ver su detalle específico.</p>
                    </div>
                </div>
                <div class="overflow-x-auto">
                    <table class="w-full text-left border-collapse">
                        <thead>
                            <tr class="bg-slate-50 text-slate-600 text-xs uppercase tracking-wider font-semibold border-b border-slate-200">
                                <th class="p-3">Cliente</th>
                                <th class="p-3 text-center">Reclamos</th>
                                <th class="p-3 text-center">Cerrados</th>
                                <th class="p-3 text-center">Abiertos</th>
                                <th class="p-3 text-center">% Cierre</th>
                                <th class="p-3 text-right">Cant. Reclamada</th>
                                <th class="p-3 text-center">Acción</th>
                            </tr>
                        </thead>
                        <tbody id="clientsTableBody" class="divide-y divide-slate-100 text-sm"></tbody>
                    </table>
                </div>
            </section>

            <!-- DETALLE DE CLIENTE SELECCIONADO (MODAL) -->
            <div id="clientDetailModal" class="hidden fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-sm flex items-center justify-center p-4">
                <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-5xl w-full max-h-[90vh] flex flex-col overflow-hidden">
                    <div class="p-6 border-b border-slate-200 flex justify-between items-center bg-slate-50">
                        <div>
                            <span class="text-xs font-bold text-sky-600 uppercase tracking-wider">Detalle por Cliente</span>
                            <h3 id="modalClientTitle" class="text-xl font-black text-slate-900 mt-1">Cliente</h3>
                        </div>
                        <button onclick="closeClientModal()" class="text-slate-400 hover:text-slate-700 text-xl font-bold p-2 cursor-pointer">✕</button>
                    </div>
                    <div class="p-6 overflow-y-auto space-y-6 flex-grow">
                        <div id="modalClientSummary" class="grid grid-cols-2 md:grid-cols-4 gap-4"></div>
                        <div class="overflow-x-auto">
                            <table class="w-full text-left border-collapse text-xs">
                                <thead class="bg-slate-50 text-slate-600 font-semibold border-b border-slate-200">
                                    <tr>
                                        <th class="p-3">Aviso</th>
                                        <th class="p-3">Fecha</th>
                                        <th class="p-3">Tipo</th>
                                        <th class="p-3">Descripción</th>
                                        <th class="p-3">Creador</th>
                                        <th class="p-3 text-center">Cant.</th>
                                        <th class="p-3 text-center">Estado</th>
                                        <th class="p-3">Cierre</th>
                                    </tr>
                                </thead>
                                <tbody id="modalClientTableBody" class="divide-y divide-slate-100"></tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>

            <!-- REPORTE POR TIPO DE RECLAMO -->
            <section id="tipos" class="bg-white rounded-2xl shadow-sm border border-slate-200 p-8">
                <div class="grid grid-cols-1 lg:grid-cols-2 gap-8 items-center">
                    <div>
                        <h3 class="text-xl font-extrabold text-slate-900 mb-1">Reporte por Tipo de Reclamo</h3>
                        <p class="text-xs text-slate-500 mb-6">Agrupación según tipo de incidencia.</p>
                        <div class="overflow-x-auto">
                            <table class="w-full text-left border-collapse text-sm">
                                <thead>
                                    <tr class="bg-slate-50 text-slate-600 text-xs uppercase tracking-wider font-semibold border-b border-slate-200">
                                        <th class="p-3">Tipo de Reclamo</th>
                                        <th class="p-3 text-center">Reclamos</th>
                                        <th class="p-3 text-center">Cerrados</th>
                                        <th class="p-3 text-center">Abiertos</th>
                                        <th class="p-3 text-center">% Cierre</th>
                                    </tr>
                                </thead>
                                <tbody id="typesTableBody" class="divide-y divide-slate-100 text-xs"></tbody>
                            </table>
                        </div>
                    </div>
                    <div class="bg-slate-50/50 p-6 rounded-xl border border-slate-200 flex flex-col justify-center" style="height: 320px;">
                        <canvas id="typesChart"></canvas>
                    </div>
                </div>
            </section>

            <!-- EVOLUCIÓN MENSUAL -->
            <section id="evolucion" class="bg-white rounded-2xl shadow-sm border border-slate-200 p-8">
                <h3 class="text-xl font-extrabold text-slate-900 mb-1">Evolución Mensual</h3>
                <p class="text-xs text-slate-500 mb-6">Agrupado cronológicamente por la fecha de creación.</p>
                
                <div class="bg-slate-50/50 p-6 rounded-xl border border-slate-200 mb-8" style="height: 300px;">
                    <canvas id="monthlyChart"></canvas>
                </div>

                <div class="overflow-x-auto">
                    <table class="w-full text-left border-collapse text-sm">
                        <thead>
                            <tr class="bg-slate-50 text-slate-600 text-xs uppercase tracking-wider font-semibold border-b border-slate-200">
                                <th class="p-3">Mes</th>
                                <th class="p-3 text-center">Reclamos</th>
                                <th class="p-3 text-center">Cerrados</th>
                                <th class="p-3 text-center">Abiertos</th>
                                <th class="p-3 text-right">Cantidad Reclamada</th>
                            </tr>
                        </thead>
                        <tbody id="monthlyTableBody" class="divide-y divide-slate-100 text-xs"></tbody>
                    </table>
                </div>
            </section>

            <!-- ESTADO DE GESTIÓN & TIEMPOS -->
            <section id="gestion" class="bg-white rounded-2xl shadow-sm border border-slate-200 p-8">
                <div class="grid grid-cols-1 lg:grid-cols-2 gap-8 items-center">
                    <div>
                        <h3 class="text-xl font-extrabold text-slate-900 mb-1">Estado de Gestión</h3>
                        <p class="text-xs text-slate-500 mb-6">Distribución de estados y tiempos de resolución (reclamos cerrados con fecha válida).</p>
                        
                        <div class="grid grid-cols-3 gap-4 mb-6">
                            <div class="bg-slate-50 p-4 rounded-xl border border-slate-200 text-center">
                                <span class="text-xs text-slate-500 uppercase block font-semibold">Promedio Días</span>
                                <span id="avgDays" class="text-2xl font-black text-slate-900 mt-1 block">-</span>
                            </div>
                            <div class="bg-slate-50 p-4 rounded-xl border border-slate-200 text-center">
                                <span class="text-xs text-slate-500 uppercase block font-semibold">Mínimo Días</span>
                                <span id="minDays" class="text-2xl font-black text-emerald-600 mt-1 block">-</span>
                            </div>
                            <div class="bg-slate-50 p-4 rounded-xl border border-slate-200 text-center">
                                <span class="text-xs text-slate-500 uppercase block font-semibold">Máximo Días</span>
                                <span id="maxDays" class="text-2xl font-black text-amber-600 mt-1 block">-</span>
                            </div>
                        </div>
                    </div>
                    <div class="bg-slate-50/50 p-6 rounded-xl border border-slate-200 flex flex-col justify-center items-center" style="height: 300px;">
                        <canvas id="statusDonutChart"></canvas>
                    </div>
                </div>
            </section>

            <!-- ANÁLISIS DE DESCRIPCIONES -->
            <section class="bg-white rounded-2xl shadow-sm border border-slate-200 p-8">
                <div class="flex justify-between items-center mb-4">
                    <div>
                        <h3 class="text-xl font-extrabold text-slate-900">Análisis de Descripciones</h3>
                        <p class="text-xs text-slate-500">Agrupación de motivos y descripciones frecuentes de los reclamos.</p>
                    </div>
                    <input type="text" id="descSearch" placeholder="Filtrar descripciones..." class="bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 text-xs text-slate-800 focus:outline-none focus:border-sky-500 w-64" oninput="debounceFilterDescriptionsTable()">
                </div>
                <div class="overflow-x-auto max-h-80 overflow-y-auto">
                    <table class="w-full text-left border-collapse text-xs">
                        <thead class="sticky top-0 bg-slate-50">
                            <tr class="text-slate-600 uppercase tracking-wider font-semibold border-b border-slate-200">
                                <th class="p-3">Descripción / Motivo del Reclamo</th>
                                <th class="p-3 text-center w-32">Cantidad</th>
                            </tr>
                        </thead>
                        <tbody id="descriptionsTableBody" class="divide-y divide-slate-100"></tbody>
                    </table>
                </div>
            </section>

            <!-- DETALLE COMPLETO DE RECLAMOS (TABLA PAGINADA) -->
            <section id="detalle" class="bg-white rounded-2xl shadow-sm border border-slate-200 p-8">
                <div class="flex justify-between items-center mb-6 flex-wrap gap-4">
                    <div>
                        <h3 class="text-xl font-extrabold text-slate-900">Detalle Completo de Reclamos</h3>
                        <p class="text-xs text-slate-500">Listado completo de registros procesados.</p>
                    </div>
                    <div class="flex items-center gap-3">
                        <span id="paginationInfo" class="text-xs text-slate-500 font-medium">Mostrando 0 de 0</span>
                        <div class="flex gap-1">
                            <button onclick="prevPage()" class="bg-slate-100 hover:bg-slate-200 px-3 py-1.5 rounded-lg text-xs font-bold text-slate-700 transition cursor-pointer">Anterior</button>
                            <button onclick="nextPage()" class="bg-slate-100 hover:bg-slate-200 px-3 py-1.5 rounded-lg text-xs font-bold text-slate-700 transition cursor-pointer">Siguiente</button>
                        </div>
                    </div>
                </div>
                <div class="overflow-x-auto">
                    <table class="w-full text-left border-collapse text-xs">
                        <thead>
                            <tr class="bg-slate-50 text-slate-600 uppercase tracking-wider font-semibold border-b border-slate-200">
                                <th class="p-3">Aviso</th>
                                <th class="p-3">Cliente</th>
                                <th class="p-3">Fecha Creación</th>
                                <th class="p-3">Tipo</th>
                                <th class="p-3">Descripción</th>
                                <th class="p-3">Material</th>
                                <th class="p-3 text-center">Cant.</th>
                                <th class="p-3">Creador</th>
                                <th class="p-3 text-center">Estado</th>
                                <th class="p-3">Fecha Cierre</th>
                            </tr>
                        </thead>
                        <tbody id="fullTableBody" class="divide-y divide-slate-100"></tbody>
                    </table>
                </div>
            </section>

        </div>
    </main>

    <!-- Footer -->
    <footer class="bg-white border-t border-slate-200 py-6 text-center text-xs text-slate-400">
        MAINCAL S.A. • Generador Automático de Reportes de Logística y Servicio al Cliente
    </footer>

    <!-- MODAL DE PRESENTACIÓN INTERACTIVA (HTML/CSS 16:9) -->
    <div id="presentationModal" class="fixed inset-0 bg-slate-950/90 z-50 hidden flex flex-col items-center justify-between p-4 md:p-6 backdrop-blur-md">
        
        <!-- Barra de Controles Superior -->
        <div class="w-full max-w-7xl flex justify-between items-center text-white mb-2 z-10">
            <div class="flex items-center gap-3">
                <span class="bg-sky-500 text-white font-black text-xs px-2.5 py-1 rounded tracking-widest uppercase">MAINCAL</span>
                <span class="text-sm font-semibold text-slate-300 hidden sm:inline">Presentación Ejecutiva de Reclamos</span>
            </div>
            
            <div class="flex items-center gap-4">
                <span id="slideIndicator" class="text-xs font-mono font-bold bg-slate-800/80 border border-slate-700/80 px-3 py-1.5 rounded-full text-slate-300">
                    Slide 1 de 1
                </span>
                
                <button onclick="toggleFullscreen()" title="Pantalla Completa" class="p-2 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-lg text-xs font-semibold transition border border-slate-700 cursor-pointer flex items-center gap-1">
                    <svg xmlns="http://www.w3.org/2000/svg" class="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 8V4m0 0h4M4 4l5 5m11-1V4m0 0h-4m4 0l-5 5M4 16v4m0 0h4m-4 0l5-5m11 5l-5-5m5 5v-4m0 4h-4"/></svg>
                    <span class="hidden md:inline">Pantalla Completa</span>
                </button>

                <button onclick="closePresentation()" title="Cerrar Presentación" class="p-2 bg-rose-600 hover:bg-rose-500 text-white rounded-lg text-xs font-bold transition cursor-pointer flex items-center gap-1 shadow">
                    <svg xmlns="http://www.w3.org/2000/svg" class="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/></svg>
                    <span>Cerrar</span>
                </button>
            </div>
        </div>

        <!-- Contenedor del Slide Actual (Proporción 16:9 Adaptable) -->
        <div class="flex-grow w-full flex items-center justify-center relative my-auto p-2">
            <div id="slideViewport" class="slide-card">
                <!-- El contenido dinámico del slide actual se inyecta aquí -->
            </div>
        </div>

        <!-- Barra de Navegación Inferior -->
        <div class="w-full max-w-7xl flex justify-between items-center text-white mt-2 z-10">
            <button onclick="prevSlide()" id="btnPrevSlide" class="bg-slate-800 hover:bg-sky-600 text-white px-5 py-2.5 rounded-xl font-bold text-xs transition border border-slate-700 disabled:opacity-40 disabled:hover:bg-slate-800 flex items-center gap-2 cursor-pointer shadow">
                ← Anterior
            </button>
            <div class="text-xs text-slate-400 hidden sm:block">
                Usa las flechas <kbd class="bg-slate-800 px-1.5 py-0.5 rounded text-slate-200 border border-slate-700">←</kbd> <kbd class="bg-slate-800 px-1.5 py-0.5 rounded text-slate-200 border border-slate-700">→</kbd> para navegar
            </div>
            <button onclick="nextSlide()" id="btnNextSlide" class="bg-sky-600 hover:bg-sky-500 text-white px-5 py-2.5 rounded-xl font-bold text-xs transition disabled:opacity-40 disabled:hover:bg-sky-600 flex items-center gap-2 cursor-pointer shadow">
                Siguiente →
            </button>
        </div>
    </div>

    <!-- SCRIPT DE PROCESAMIENTO Y LÓGICA COMPLETA -->
    <script>
        const monthNames = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
        let rawData = [];
        let filteredData = [];
        let currentPage = 1;
        const pageSize = 25;
        let charts = {};
        let slideCharts = {};
        let debounceTimer = null;
        let descDebounceTimer = null;

        // Variables para la Presentación
        let slidesData = [];
        let currentSlideIndex = 0;

        document.getElementById('excelFile').addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;

            document.getElementById('loadingMsg').classList.remove('hidden');

            setTimeout(() => {
                const reader = new FileReader();
                reader.onload = function(e) {
                    try {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, { type: 'array' });
                        const firstSheet = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[firstSheet];
                        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                        
                        processExcelRows(rows);
                    } catch (err) {
                        alert('Error al procesar el archivo. Asegúrate de que sea un Excel o CSV válido.');
                        console.error(err);
                        document.getElementById('loadingMsg').classList.add('hidden');
                    }
                };
                reader.readAsArrayBuffer(file);
            }, 50);
        });

        function parseExcelDate(val) {
            if (!val) return null;
            if (typeof val === 'number') {
                const d = XLSX.SSF.parse_date_code(val);
                if (!d) return null;
                return new Date(d.y, d.m - 1, d.d);
            }
            if (typeof val === 'string') {
                let parts = val.split(/[\/\-]/);
                if (parts.length === 3) {
                    if (parts[0].length === 4) {
                        let d = new Date(val);
                        return isNaN(d.getTime()) ? null : d;
                    } else {
                        let d = new Date(parts[2], parts[1] - 1, parts[0]);
                        return isNaN(d.getTime()) ? null : d;
                    }
                }
                let d = new Date(val);
                return isNaN(d.getTime()) ? null : d;
            }
            return null;
        }

        function processExcelRows(rows) {
            if (rows.length < 2) {
                alert('El archivo está vacío.');
                document.getElementById('loadingMsg').classList.add('hidden');
                return;
            }

            let records = [];
            for (let i = 1; i < rows.length; i++) {
                let r = rows[i];
                if (!r || r.length === 0) continue;

                let aviso = r[2] !== undefined && r[2] !== null ? r[2].toString().trim() : 'S/A';
                if (aviso === '' || aviso === 'S/A') continue;

                // CORREGIDO: Se separan los índices de Cliente (r[3]) y Tipo (r[5]) para evitar solapamientos
                let cliente = r[3] !== undefined && r[3] !== null && r[3] !== '' ? r[3].toString().trim() : 'SIN CLIENTE';
                let tipo = r[5] !== undefined && r[5] !== null && r[5] !== '' ? r[5].toString().trim() : (r[4] !== undefined ? r[4].toString().trim() : 'SIN TIPO');
                
                let desc = r[6] !== undefined && r[6] !== null && r[6] !== '' ? r[6].toString().trim() : 'SIN DESCRIPCIÓN';
                let material = r[7] !== undefined && r[7] !== null ? r[7].toString().trim() : '';
                let matText = r[8] !== undefined && r[8] !== null ? r[8].toString().trim() : '';
                let cant = parseFloat(r[12]);
                if (isNaN(cant)) cant = 0;

                let creador = r[15] !== undefined && r[15] !== null && r[15] !== '' ? r[15].toString().trim() : 'SIN INFORMACIÓN';
                let fechaCreado = parseExcelDate(r[16]);
                let fechaCierre = parseExcelDate(r[18]);
                let estado = fechaCierre ? 'CERRADO' : 'ABIERTO / PENDIENTE';

                records.push({
                    aviso: aviso,
                    tipo: tipo,
                    desc: desc,
                    material: material ? `${material} - ${matText}` : matText,
                    cantidad: cant,
                    creador: creador,
                    fechaCreado: fechaCreado,
                    cliente: cliente,
                    fechaCierre: fechaCierre,
                    estado: estado,
                    year: fechaCreado ? fechaCreado.getFullYear() : null,
                    monthNum: fechaCreado ? fechaCreado.getMonth() : null,
                    monthName: fechaCreado ? monthNames[fechaCreado.getMonth()] : null,
                    searchStr: (aviso + ' ' + cliente + ' ' + desc + ' ' + (material ? `${material} - ${matText}` : matText) + ' ' + creador + ' ' + tipo).toLowerCase()
                });
            }

            rawData = records;
            filteredData = [...rawData];

            if (rawData.length === 0) {
                alert('No se encontraron registros válidos en las columnas indicadas.');
                document.getElementById('loadingMsg').classList.add('hidden');
                return;
            }

            document.getElementById('uploadCard').classList.add('hidden');
            document.getElementById('reportContainer').classList.remove('hidden');
            document.getElementById('presentationBtn').classList.remove('hidden');
            document.getElementById('exportExcelBtn').classList.remove('hidden');
            document.getElementById('resetFileBtn').classList.remove('hidden');

            initFilters();
            updateReport();
            document.getElementById('loadingMsg').classList.add('hidden');
        }

        function exportToExcel() {
            if (!filteredData || filteredData.length === 0) {
                alert('No hay datos para exportar.');
                return;
            }

            const exportRows = filteredData.map(r => ({
                'Aviso': r.aviso,
                'Cliente': r.cliente,
                'Fecha Creación': r.fechaCreado ? r.fechaCreado.toLocaleDateString('es-AR') : '',
                'Tipo de Reclamo': r.tipo,
                'Descripción': r.desc,
                'Material': r.material,
                'Cantidad': r.cantidad,
                'Creador': r.creador,
                'Estado': r.estado,
                'Fecha Cierre': r.fechaCierre ? r.fechaCierre.toLocaleDateString('es-AR') : 'PENDIENTE'
            }));

            const worksheet = XLSX.utils.json_to_sheet(exportRows);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "Reclamos Filtrados");

            const fechaStr = new Date().toISOString().slice(0, 10);
            XLSX.writeFile(workbook, `Reporte_Reclamos_${fechaStr}.xlsx`);
        }

        function initFilters() {
            let clients = [...new Set(rawData.map(r => r.cliente))].sort();
            let types = [...new Set(rawData.map(r => r.tipo))].sort();
            let creators = [...new Set(rawData.map(r => r.creador))].sort();
            let years = [...new Set(rawData.map(r => r.year).filter(y => y !== null))].sort((a,b)=>a-b);

            fillSelect('filterClient', clients);
            fillSelect('filterType', types);
            fillSelect('filterCreator', creators);
            fillSelect('filterYear', years);

            let monthSelect = document.getElementById('filterMonth');
            monthSelect.innerHTML = '<option value="">Todos los meses</option>';
            monthNames.forEach((m, idx) => {
                monthSelect.innerHTML += `<option value="${idx}">${m}</option>`;
            });
        }

        function fillSelect(id, items) {
            let sel = document.getElementById(id);
            let current = sel.value;
            sel.innerHTML = `<option value="">Todos</option>`;
            items.forEach(i => {
                sel.innerHTML += `<option value="${i}">${i}</option>`;
            });
            sel.value = current;
        }

        function debounceApplyFilters() {
            clearTimeout(debounceTimer);
            debounceTimer = setTimeout(() => {
                applyFilters();
            }, 200);
        }

        function applyFilters() {
            let sText = document.getElementById('globalSearch').value.toLowerCase().trim();
            let fClient = document.getElementById('filterClient').value;
            let fType = document.getElementById('filterType').value;
            let fStatus = document.getElementById('filterStatus').value;
            let fCreator = document.getElementById('filterCreator').value;
            let fYear = document.getElementById('filterYear').value;
            let fMonth = document.getElementById('filterMonth').value;

            filteredData = rawData.filter(r => {
                if (fClient && r.cliente !== fClient) return false;
                if (fType && r.tipo !== fType) return false;
                if (fStatus && r.estado !== fStatus) return false;
                if (fCreator && r.creador !== fCreator) return false;
                if (fYear && r.year != fYear) return false;
                if (fMonth !== "" && r.monthNum != fMonth) return false;

                if (sText && !r.searchStr.includes(sText)) return false;

                return true;
            });

            currentPage = 1;
            updateReport();
        }

        function clearFilters() {
            document.getElementById('globalSearch').value = '';
            document.getElementById('filterClient').value = '';
            document.getElementById('filterType').value = '';
            document.getElementById('filterStatus').value = '';
            document.getElementById('filterCreator').value = '';
            document.getElementById('filterYear').value = '';
            document.getElementById('filterMonth').value = '';
            filteredData = [...rawData];
            currentPage = 1;
            updateReport();
        }

        function updateReport() {
            renderSummaryCards();
            renderClientsTable();
            renderTypesSection();
            renderMonthlySection();
            renderGestionSection();
            renderDescriptionsTable();
            renderFullTable();
        }

        function renderSummaryCards() {
            let total = filteredData.length;
            let closed = 0;
            let totalQty = 0;
            let clientsSet = new Set();
            let creatorsSet = new Set();
            let minTime = Infinity;
            let maxTime = -Infinity;

            for (let i = 0; i < total; i++) {
                let r = filteredData[i];
                if (r.estado === 'CERRADO') closed++;
                totalQty += r.cantidad;
                clientsSet.add(r.cliente);
                creatorsSet.add(r.creador);
                if (r.fechaCreado) {
                    let t = r.fechaCreado.getTime();
                    if (t < minTime) minTime = t;
                    if (t > maxTime) maxTime = t;
                }
            }

            let pending = total - closed;
            let closedPct = total > 0 ? Math.round((closed / total) * 100) : 0;

            document.getElementById('cardTotal').innerText = total;
            document.getElementById('cardClosed').innerText = closed;
            document.getElementById('cardClosedPct').innerText = `${closedPct}%`;
            document.getElementById('cardPending').innerText = pending;
            document.getElementById('cardQty').innerText = totalQty.toLocaleString('es-AR');

            document.getElementById('infoClientsCount').innerText = clientsSet.size;
            document.getElementById('infoCreatorsCount').innerText = creatorsSet.size;

            if (minTime !== Infinity && maxTime !== -Infinity) {
                let minDate = new Date(minTime);
                let maxDate = new Date(maxTime);
                document.getElementById('infoDates').innerText = `${minDate.toLocaleDateString('es-AR')} al ${maxDate.toLocaleDateString('es-AR')}`;
                document.getElementById('periodSubtitle').innerText = `Período analizado: ${minDate.toLocaleDateString('es-AR')} al ${maxDate.toLocaleDateString('es-AR')}`;
            } else {
                document.getElementById('infoDates').innerText = 'Sin fechas válidas';
                document.getElementById('periodSubtitle').innerText = 'Período analizado: General';
            }
        }

        function renderClientsTable() {
            let map = {};
            for (let i = 0; i < filteredData.length; i++) {
                let r = filteredData[i];
                let c = r.cliente;
                if (!map[c]) {
                    map[c] = { total: 0, closed: 0, pending: 0, qty: 0 };
                }
                map[c].total++;
                if (r.estado === 'CERRADO') map[c].closed++;
                else map[c].pending++;
                map[c].qty += r.cantidad;
            }

            let clientList = Object.keys(map).map(c => ({
                name: c,
                ...map[c],
                pct: Math.round((map[c].closed / map[c].total) * 100)
            })).sort((a, b) => b.total - a.total);

            let tbody = document.getElementById('clientsTableBody');
            if (clientList.length === 0) {
                tbody.innerHTML = `<tr><td colspan="7" class="p-4 text-center text-slate-400">No hay registros para mostrar.</td></tr>`;
                return;
            }

            let html = '';
            for (let i = 0; i < clientList.length; i++) {
                let c = clientList[i];
                html += `
                    <tr class="hover:bg-slate-50 transition cursor-pointer" onclick="openClientDetail('${c.name.replace(/'/g, "\\'")}')">
                        <td class="p-3 font-semibold text-slate-900">${c.name}</td>
                        <td class="p-3 text-center font-bold">${c.total}</td>
                        <td class="p-3 text-center text-emerald-600 font-semibold">${c.closed}</td>
                        <td class="p-3 text-center text-amber-500 font-semibold">${c.pending}</td>
                        <td class="p-3 text-center"><span class="bg-slate-100 text-slate-800 text-xs font-bold px-2 py-1 rounded">${c.pct}%</span></td>
                        <td class="p-3 text-right font-medium">${c.qty.toLocaleString('es-AR')}</td>
                        <td class="p-3 text-center text-sky-600 font-semibold hover:underline">Ver Detalle →</td>
                    </tr>
                `;
            }
            tbody.innerHTML = html;
        }

        function openClientDetail(clientName) {
            document.getElementById('modalClientTitle').innerText = `Cliente: ${clientName}`;
            let clientRecords = filteredData.filter(r => r.cliente === clientName);
            
            let total = clientRecords.length;
            let closed = clientRecords.filter(r => r.estado === 'CERRADO').length;
            let pending = total - closed;
            let pct = total > 0 ? Math.round((closed / total) * 100) : 0;
            let qty = clientRecords.reduce((acc, r) => acc + r.cantidad, 0);

            document.getElementById('modalClientSummary').innerHTML = `
                <div class="bg-slate-50 p-3 rounded-xl border border-slate-200"><span class="text-xs text-slate-500 block">Total</span><span class="text-xl font-bold">${total}</span></div>
                <div class="bg-slate-50 p-3 rounded-xl border border-slate-200"><span class="text-xs text-slate-500 block">Cerrados</span><span class="text-xl font-bold text-emerald-600">${closed} (${pct}%)</span></div>
                <div class="bg-slate-50 p-3 rounded-xl border border-slate-200"><span class="text-xs text-slate-500 block">Abiertos</span><span class="text-xl font-bold text-amber-500">${pending}</span></div>
                <div class="bg-slate-50 p-3 rounded-xl border border-slate-200"><span class="text-xs text-slate-500 block">Cant. Reclamada</span><span class="text-xl font-bold">${qty.toLocaleString('es-AR')}</span></div>
            `;

            let mBody = document.getElementById('modalClientTableBody');
            let html = '';
            for (let i = 0; i < clientRecords.length; i++) {
                let r = clientRecords[i];
                let fCre = r.fechaCreado ? r.fechaCreado.toLocaleDateString('es-AR') : '-';
                let fCie = r.fechaCierre ? r.fechaCierre.toLocaleDateString('es-AR') : 'PENDIENTE';
                let badgeClass = r.estado === 'CERRADO' ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800';
                html += `
                    <tr class="hover:bg-slate-50">
                        <td class="p-2 font-mono font-bold">${r.aviso}</td>
                        <td class="p-2">${fCre}</td>
                        <td class="p-2">${r.tipo}</td>
                        <td class="p-2">${r.desc}</td>
                        <td class="p-2">${r.creador}</td>
                        <td class="p-2 text-center font-bold">${r.cantidad}</td>
                        <td class="p-2 text-center"><span class="${badgeClass} text-[10px] font-bold px-2 py-0.5 rounded">${r.estado}</span></td>
                        <td class="p-2">${fCie}</td>
                    </tr>
                `;
            }
            mBody.innerHTML = html;
            document.getElementById('clientDetailModal').classList.remove('hidden');
        }

        function closeClientModal() {
            document.getElementById('clientDetailModal').classList.add('hidden');
        }

        function renderTypesSection() {
            let map = {};
            for (let i = 0; i < filteredData.length; i++) {
                let r = filteredData[i];
                let t = r.tipo;
                if (!map[t]) map[t] = { total: 0, closed: 0, pending: 0 };
                map[t].total++;
                if (r.estado === 'CERRADO') map[t].closed++;
                else map[t].pending++;
            }

            let typeList = Object.keys(map).map(t => ({
                name: t,
                ...map[t],
                pct: Math.round((map[t].closed / map[t].total) * 100)
            })).sort((a, b) => b.total - a.total);

            let tbody = document.getElementById('typesTableBody');
            let html = '';
            for (let i = 0; i < typeList.length; i++) {
                let t = typeList[i];
                html += `
                    <tr>
                        <td class="p-3 font-semibold text-slate-900">${t.name}</td>
                        <td class="p-3 text-center font-bold">${t.total}</td>
                        <td class="p-3 text-center text-emerald-600">${t.closed}</td>
                        <td class="p-3 text-center text-amber-500">${t.pending}</td>
                        <td class="p-3 text-center"><span class="bg-slate-100 text-xs font-bold px-2 py-1 rounded">${t.pct}%</span></td>
                    </tr>
                `;
            }
            tbody.innerHTML = html;

            let ctx = document.getElementById('typesChart').getContext('2d');
            if (charts.types) charts.types.destroy();
            charts.types = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: typeList.map(t => t.name),
                    datasets: [{
                        label: 'Reclamos',
                        data: typeList.map(t => t.total),
                        backgroundColor: '#0284c7',
                        borderRadius: 6,
                        barThickness: 24
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { display: false } },
                    scales: {
                        y: { beginAtZero: true, grid: { color: '#f1f5f9' } },
                        x: { grid: { display: false }, ticks: { font: { size: 10 } } }
                    }
                }
            });
        }

        function renderMonthlySection() {
            let monthlyMap = {};
            monthNames.forEach((m, idx) => {
                monthlyMap[idx] = { name: m, total: 0, closed: 0, pending: 0, qty: 0 };
            });

            for (let i = 0; i < filteredData.length; i++) {
                let r = filteredData[i];
                if (r.monthNum !== null && monthlyMap[r.monthNum]) {
                    monthlyMap[r.monthNum].total++;
                    if (r.estado === 'CERRADO') monthlyMap[r.monthNum].closed++;
                    else monthlyMap[r.monthNum].pending++;
                    monthlyMap[r.monthNum].qty += r.cantidad;
                }
            }

            let monthlyArray = Object.values(monthlyMap);
            let tbody = document.getElementById('monthlyTableBody');
            let html = '';
            for (let i = 0; i < monthlyArray.length; i++) {
                let m = monthlyArray[i];
                html += `
                    <tr>
                        <td class="p-3 font-semibold text-slate-900">${m.name}</td>
                        <td class="p-3 text-center font-bold">${m.total}</td>
                        <td class="p-3 text-center text-emerald-600">${m.closed}</td>
                        <td class="p-3 text-center text-amber-500">${m.pending}</td>
                        <td class="p-3 text-right font-medium">${m.qty.toLocaleString('es-AR')}</td>
                    </tr>
                `;
            }
            tbody.innerHTML = html;

            let labels = monthNames;
            let totals = monthlyArray.map(m => m.total);

            let ctx1 = document.getElementById('monthlyChart').getContext('2d');
            if (charts.monthly) charts.monthly.destroy();
            charts.monthly = new Chart(ctx1, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [{
                        label: 'Total Reclamos por Mes',
                        data: totals,
                        backgroundColor: '#0284c7',
                        borderRadius: 6,
                        barThickness: 24
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { display: false } },
                    scales: { y: { beginAtZero: true, grid: { color: '#f1f5f9' } }, x: { grid: { display: false } } }
                }
            });
        }

        function renderGestionSection() {
            let total = filteredData.length;
            let closed = 0;
            let daysDiffs = [];

            for (let i = 0; i < total; i++) {
                let r = filteredData[i];
                if (r.estado === 'CERRADO') closed++;
                if (r.fechaCierre && r.fechaCreado) {
                    let diffTime = r.fechaCierre - r.fechaCreado;
                    let diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                    if (diffDays >= 0) daysDiffs.push(diffDays);
                }
            }
            let pending = total - closed;

            if (daysDiffs.length > 0) {
                let sum = 0, min = daysDiffs[0], max = daysDiffs[0];
                for (let i = 0; i < daysDiffs.length; i++) {
                    sum += daysDiffs[i];
                    if (daysDiffs[i] < min) min = daysDiffs[i];
                    if (daysDiffs[i] > max) max = daysDiffs[i];
                }
                let avg = Math.round(sum / daysDiffs.length);

                document.getElementById('avgDays').innerText = `${avg} días`;
                document.getElementById('minDays').innerText = `${min} días`;
                document.getElementById('maxDays').innerText = `${max} días`;
            } else {
                document.getElementById('avgDays').innerText = '-';
                document.getElementById('minDays').innerText = '-';
                document.getElementById('maxDays').innerText = '-';
            }

            let ctx = document.getElementById('statusDonutChart').getContext('2d');
            if (charts.statusDonut) charts.statusDonut.destroy();
            charts.statusDonut = new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: ['Cerrados', 'Abiertos / Pendientes'],
                    datasets: [{
                        data: [closed, pending],
                        backgroundColor: ['#10b981', '#f59e0b'],
                        borderWidth: 2,
                        borderColor: '#ffffff'
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { position: 'bottom' } },
                    cutout: '65%'
                }
            });
        }

        let descCache = [];
        function renderDescriptionsTable() {
            let map = {};
            for (let i = 0; i < filteredData.length; i++) {
                let d = filteredData[i].desc;
                map[d] = (map[d] || 0) + 1;
            }

            descCache = Object.keys(map).map(d => ({ desc: d, count: map[d] })).sort((a, b) => b.count - a.count);
            displayDescriptionsTable(descCache);
        }

        function displayDescriptionsTable(items) {
            let tbody = document.getElementById('descriptionsTableBody');
            let html = '';
            for (let i = 0; i < items.length; i++) {
                html += `
                    <tr>
                        <td class="p-2 text-slate-800">${items[i].desc}</td>
                        <td class="p-2 text-center font-bold">${items[i].count}</td>
                    </tr>
                `;
            }
            tbody.innerHTML = html;
        }

        function debounceFilterDescriptionsTable() {
            clearTimeout(descDebounceTimer);
            descDebounceTimer = setTimeout(() => {
                filterDescriptionsTable();
            }, 200);
        }

        function filterDescriptionsTable() {
            let q = document.getElementById('descSearch').value.toLowerCase();
            let filtered = descCache.filter(i => i.desc.toLowerCase().includes(q));
            displayDescriptionsTable(filtered);
        }

        function renderFullTable() {
            let tbody = document.getElementById('fullTableBody');
            let start = (currentPage - 1) * pageSize;
            let end = start + pageSize;
            let pageItems = filteredData.slice(start, end);

            document.getElementById('paginationInfo').innerText = `Mostrando ${filteredData.length > 0 ? start + 1 : 0} - ${Math.min(end, filteredData.length)} de ${filteredData.length}`;

            let html = '';
            for (let i = 0; i < pageItems.length; i++) {
                let r = pageItems[i];
                let fCre = r.fechaCreado ? r.fechaCreado.toLocaleDateString('es-AR') : '-';
                let fCie = r.fechaCierre ? r.fechaCierre.toLocaleDateString('es-AR') : '-';
                let badgeClass = r.estado === 'CERRADO' ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800';

                html += `
                    <tr class="hover:bg-slate-50">
                        <td class="p-2 font-mono font-bold">${r.aviso}</td>
                        <td class="p-2">${r.cliente}</td>
                        <td class="p-2">${fCre}</td>
                        <td class="p-2">${r.tipo}</td>
                        <td class="p-2">${r.desc}</td>
                        <td class="p-2">${r.material}</td>
                        <td class="p-2 text-center font-bold">${r.cantidad}</td>
                        <td class="p-2">${r.creador}</td>
                        <td class="p-2 text-center"><span class="${badgeClass} text-[10px] font-bold px-2 py-0.5 rounded">${r.estado}</span></td>
                        <td class="p-2">${fCie}</td>
                    </tr>
                `;
            }
            tbody.innerHTML = html;
        }

        function prevPage() {
            if (currentPage > 1) {
                currentPage--;
                renderFullTable();
            }
        }

        function nextPage() {
            if (currentPage * pageSize < filteredData.length) {
                currentPage++;
                renderFullTable();
            }
        }

        /* ==========================================================================
           LÓGICA DEL MODO PRESENTACIÓN EN DIAPOSITIVAS (HTML 16:9)
           ========================================================================== */

        function openPresentation() {
            if (!filteredData || filteredData.length === 0) {
                alert('No hay datos para presentar.');
                return;
            }

            buildPresentationSlides();
            currentSlideIndex = 0;
            document.getElementById('presentationModal').classList.remove('hidden');
            renderSlide(currentSlideIndex);

            // Escuchador de teclado
            window.addEventListener('keydown', handleKeyDown);
        }

        function closePresentation() {
            document.getElementById('presentationModal').classList.add('hidden');
            window.removeEventListener('keydown', handleKeyDown);
            if (document.fullscreenElement) {
                document.exitFullscreen().catch(() => {});
            }
        }

        function toggleFullscreen() {
            const elem = document.getElementById('presentationModal');
            if (!document.fullscreenElement) {
                elem.requestFullscreen().catch(err => {
                    alert(`Error al activar pantalla completa: ${err.message}`);
                });
            } else {
                document.exitFullscreen().catch(() => {});
            }
        }

        function handleKeyDown(e) {
            if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') {
                e.preventDefault();
                nextSlide();
            } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
                e.preventDefault();
                prevSlide();
            } else if (e.key === 'Escape') {
                closePresentation();
            }
        }

        function nextSlide() {
            if (currentSlideIndex < slidesData.length - 1) {
                currentSlideIndex++;
                renderSlide(currentSlideIndex);
            }
        }

        function prevSlide() {
            if (currentSlideIndex > 0) {
                currentSlideIndex--;
                renderSlide(currentSlideIndex);
            }
        }

        function buildPresentationSlides() {
            slidesData = [];
            const T = filteredData.length;
            const closed = filteredData.filter(r => r.estado === 'CERRADO').length;
            const pending = T - closed;
            const closedPct = T > 0 ? Math.round((closed / T) * 100) : 0;
            const qty = filteredData.reduce((acc, r) => acc + r.cantidad, 0);

            // Período
            let dates = filteredData.map(r => r.fechaCreado).filter(d => d !== null);
            let periodText = "General";
            if (dates.length > 0) {
                let minD = new Date(Math.min(...dates));
                let maxD = new Date(Math.max(...dates));
                periodText = `${minD.toLocaleDateString('es-AR')} al ${maxD.toLocaleDateString('es-AR')}`;
            }

            // 1. PORTADA
            slidesData.push({
                type: 'cover',
                title: 'REPORTE DE GESTIÓN DE RECLAMOS',
                period: periodText,
                total: T
            });

            // 2. RESUMEN GENERAL (INDICADORES PRINCIPALES)
            slidesData.push({
                type: 'summary',
                total: T,
                closed: closed,
                pending: pending,
                closedPct: closedPct,
                qty: qty
            });

            // 3. RECLAMOS POR CLIENTE (PAGINADO SI HAY MUCHOS)
            let clientMap = {};
            filteredData.forEach(r => {
                let c = r.cliente;
                if (!clientMap[c]) clientMap[c] = { total: 0, closed: 0, pending: 0, qty: 0 };
                clientMap[c].total++;
                if (r.estado === 'CERRADO') clientMap[c].closed++;
                else clientMap[c].pending++;
                clientMap[c].qty += r.cantidad;
            });
            let clientsList = Object.keys(clientMap).map(c => ({
                name: c,
                ...clientMap[c],
                pct: Math.round((clientMap[c].closed / clientMap[c].total) * 100)
            })).sort((a, b) => b.total - a.total);

            const clientsPerPage = 8;
            for (let i = 0; i < clientsList.length; i += clientsPerPage) {
                slidesData.push({
                    type: 'clients',
                    items: clientsList.slice(i, i + clientsPerPage),
                    page: Math.floor(i / clientsPerPage) + 1,
                    totalPages: Math.ceil(clientsList.length / clientsPerPage)
                });
            }

            // 4. RECLAMOS POR TIPO
            let typeMap = {};
            filteredData.forEach(r => {
                let t = r.tipo;
                if (!typeMap[t]) typeMap[t] = { total: 0, closed: 0, pending: 0 };
                typeMap[t].total++;
                if (r.estado === 'CERRADO') typeMap[t].closed++;
                else typeMap[t].pending++;
            });
            let typesList = Object.keys(typeMap).map(t => ({
                name: t,
                ...typeMap[t],
                pct: Math.round((typeMap[t].closed / typeMap[t].total) * 100)
            })).sort((a, b) => b.total - a.total);

            slidesData.push({
                type: 'types',
                items: typesList
            });

            // 5. EVOLUCIÓN MENSUAL (UN GRÁFICO DE BARRAS + TABLA ABAJO)
            let monthlyMap = {};
            monthNames.forEach((m, idx) => {
                monthlyMap[idx] = { name: m, total: 0, closed: 0, pending: 0, qty: 0 };
            });
            filteredData.forEach(r => {
                if (r.monthNum !== null && monthlyMap[r.monthNum]) {
                    monthlyMap[r.monthNum].total++;
                    if (r.estado === 'CERRADO') monthlyMap[r.monthNum].closed++;
                    else monthlyMap[r.monthNum].pending++;
                    monthlyMap[r.monthNum].qty += r.cantidad;
                }
            });

            slidesData.push({
                type: 'monthly',
                items: Object.values(monthlyMap)
            });

            // 6. ESTADO Y SEGUIMIENTO DE LOS RECLAMOS (TIEMPOS)
            let daysDiffs = [];
            filteredData.forEach(r => {
                if (r.estado === 'CERRADO' && r.fechaCierre && r.fechaCreado) {
                    let diffDays = Math.ceil((r.fechaCierre - r.fechaCreado) / (1000 * 60 * 60 * 24));
                    if (diffDays >= 0) daysDiffs.push(diffDays);
                }
            });
            let avg = daysDiffs.length ? Math.round(daysDiffs.reduce((a,b)=>a+b,0)/daysDiffs.length) : '-';
            let min = daysDiffs.length ? Math.min(...daysDiffs) : '-';
            let max = daysDiffs.length ? Math.max(...daysDiffs) : '-';

            slidesData.push({
                type: 'gestion',
                total: T,
                closed: closed,
                pending: pending,
                avg: avg,
                min: min,
                max: max
            });

            // 7. ANÁLISIS DE DESCRIPCIONES (SI CORRESPONDE)
            let descMap = {};
            filteredData.forEach(r => {
                descMap[r.desc] = (descMap[r.desc] || 0) + 1;
            });
            let descsList = Object.keys(descMap).map(d => ({ desc: d, count: descMap[d] })).sort((a, b) => b.count - a.count);

            if (descsList.length > 0) {
                const descsPerPage = 7;
                for (let i = 0; i < Math.min(descsList.length, 21); i += descsPerPage) { // Máximo 3 diapositivas
                    slidesData.push({
                        type: 'descriptions',
                        items: descsList.slice(i, i + descsPerPage),
                        page: Math.floor(i / descsPerPage) + 1,
                        totalPages: Math.ceil(Math.min(descsList.length, 21) / descsPerPage)
                    });
                }
            }

            // 8. DETALLE COMPLETO DE RECLAMOS (DISTRIBUIDO EN VARIAS DIAPOSITIVAS)
            const detailPerPage = 6;
            for (let i = 0; i < filteredData.length; i += detailPerPage) {
                slidesData.push({
                    type: 'detail',
                    items: filteredData.slice(i, i + detailPerPage),
                    page: Math.floor(i / detailPerPage) + 1,
                    totalPages: Math.ceil(filteredData.length / detailPerPage)
                });
            }
        }

        function renderSlide(index) {
            const slide = slidesData[index];
            const viewport = document.getElementById('slideViewport');
            
            // Destruir gráficos anteriores de slide si existen
            Object.keys(slideCharts).forEach(key => {
                if (slideCharts[key]) slideCharts[key].destroy();
            });
            slideCharts = {};

            // Actualizar Indicador
            document.getElementById('slideIndicator').innerText = `Slide ${index + 1} de ${slidesData.length}`;
            document.getElementById('btnPrevSlide').disabled = (index === 0);
            document.getElementById('btnNextSlide').disabled = (index === slidesData.length - 1);

            // Generar Header Estándar de Diapositiva
            const makeHeader = (title, category = "MAINCAL S.A.") => `
                <div class="px-8 py-4 bg-white border-b border-sky-100 flex justify-between items-center z-10 shrink-0">
                    <div>
                        <span class="text-[10px] font-extrabold uppercase tracking-widest text-sky-600">${category}</span>
                        <h2 class="text-xl font-black text-[#0b2447] tracking-tight">${title}</h2>
                    </div>
                    <div class="text-right">
                        <span class="text-xs font-bold text-slate-400 block">Reporte de Gestión</span>
                    </div>
                </div>
            `;

            // Generar Footer Estándar de Diapositiva
            const makeFooter = () => `
                <div class="px-8 py-2.5 bg-slate-50 border-t border-slate-200 flex justify-between items-center text-[11px] font-semibold text-slate-400 shrink-0 mt-auto">
                    <span>MAINCAL S.A. • Supply Chain & Servicio al Cliente</span>
                    <span>Diapositiva ${index + 1} / ${slidesData.length}</span>
                </div>
            `;

            let html = '';

            switch (slide.type) {
                case 'cover':
                    html = `
                        <div class="relative w-full h-full flex flex-col justify-between p-12 bg-white overflow-hidden">
                            <div class="slide-cover-accent"></div>
                            <div class="slide-cover-accent-sub"></div>
                            <div class="slide-geo-pattern"></div>

                            <div class="ml-6 z-10">
                                <span class="bg-sky-100 text-sky-800 text-xs font-extrabold px-3 py-1 rounded-full uppercase tracking-widest">
                                    Informe Ejecutivo
                                </span>
                                <h1 class="text-4xl lg:text-5xl font-black text-[#0b2447] mt-6 leading-tight">
                                    ${slide.title}
                                </h1>
                                <p class="text-lg text-sky-600 font-semibold mt-2">
                                    Análisis Consolidado de Incidencias
                                </p>
                            </div>

                            <div class="ml-6 z-10 grid grid-cols-2 gap-8 border-t border-slate-200 pt-6 max-w-lg">
                                <div>
                                    <span class="text-xs text-slate-400 uppercase font-semibold block">Período Analizado</span>
                                    <span class="text-sm font-bold text-slate-800">${slide.period}</span>
                                </div>
                                <div>
                                    <span class="text-xs text-slate-400 uppercase font-semibold block">Total Registros</span>
                                    <span class="text-sm font-bold text-slate-800">${slide.total} casos</span>
                                </div>
                            </div>

                            <div class="ml-6 z-10 flex justify-between items-center text-xs text-slate-400">
                                <span>MAINCAL S.A.</span>
                                <span>Supply Chain y Servicio al Cliente</span>
                            </div>
                        </div>
                    `;
                    viewport.innerHTML = html;
                    break;

                case 'summary':
                    html = `
                        ${makeHeader("Resumen Ejecutivo de Reclamos")}
                        <div class="p-8 flex-grow flex flex-col justify-center space-y-6">
                            <div class="grid grid-cols-4 gap-4">
                                <div class="bg-white border-2 border-slate-200 rounded-2xl p-5 text-center shadow-sm">
                                    <span class="text-xs font-bold text-slate-400 uppercase tracking-wider block">Total Reclamos</span>
                                    <span class="text-4xl font-black text-[#0b2447] mt-2 block">${slide.total}</span>
                                </div>
                                <div class="bg-emerald-50/50 border-2 border-emerald-200 rounded-2xl p-5 text-center shadow-sm">
                                    <span class="text-xs font-bold text-emerald-700 uppercase tracking-wider block">Cerrados</span>
                                    <span class="text-4xl font-black text-emerald-600 mt-2 block">${slide.closed}</span>
                                    <span class="text-xs font-bold text-emerald-600 mt-1 block">${slide.closedPct}% de efectividad</span>
                                </div>
                                <div class="bg-amber-50/50 border-2 border-amber-200 rounded-2xl p-5 text-center shadow-sm">
                                    <span class="text-xs font-bold text-amber-700 uppercase tracking-wider block">Abiertos / Pendientes</span>
                                    <span class="text-4xl font-black text-amber-500 mt-2 block">${slide.pending}</span>
                                </div>
                                <div class="bg-sky-50/50 border-2 border-sky-200 rounded-2xl p-5 text-center shadow-sm">
                                    <span class="text-xs font-bold text-sky-700 uppercase tracking-wider block">Cantidad Reclamada</span>
                                    <span class="text-3xl font-black text-[#0b2447] mt-2 block">${slide.qty.toLocaleString('es-AR')}</span>
                                    <span class="text-xs font-semibold text-sky-600 mt-1 block">pares en total</span>
                                </div>
                            </div>

                            <div class="bg-slate-50 border-l-4 border-[#0b2447] rounded-xl p-5 text-slate-700 text-sm space-y-2">
                                <div class="font-bold text-[#0b2447] text-base mb-1">📌 Indicadores Principales:</div>
                                <p class="flex items-center gap-2"><span class="text-sky-600 font-bold">•</span> Se han procesado un total de <strong>${slide.total} reclamos</strong> en el período seleccionado.</p>
                                <p class="flex items-center gap-2"><span class="text-sky-600 font-bold">•</span> El nivel de resolución alcanza el <strong>${slide.closedPct}%</strong> con ${slide.closed} casos cerrados.</p>
                                <p class="flex items-center gap-2"><span class="text-sky-600 font-bold">•</span> Se registra un impacto total de <strong>${slide.qty.toLocaleString('es-AR')} pares</strong> reclamados.</p>
                            </div>
                        </div>
                        ${makeFooter()}
                    `;
                    viewport.innerHTML = html;
                    break;

                case 'clients':
                    let clientRows = slide.items.map(c => `
                        <tr>
                            <td class="font-bold text-slate-800">${c.name}</td>
                            <td class="text-center font-bold">${c.total}</td>
                            <td class="text-center text-emerald-600 font-bold">${c.closed}</td>
                            <td class="text-center text-amber-600 font-bold">${c.pending}</td>
                            <td class="text-center"><span class="bg-sky-100 text-sky-800 font-bold px-2 py-0.5 rounded text-xs">${c.pct}%</span></td>
                            <td class="text-right font-medium">${c.qty.toLocaleString('es-AR')}</td>
                        </tr>
                    `).join('');

                    html = `
                        ${makeHeader(`Reclamos por Cliente ${slide.totalPages > 1 ? `(${slide.page}/${slide.totalPages})` : ''}`)}
                        <div class="p-8 flex-grow flex flex-col justify-start overflow-hidden">
                            <table class="slide-table shadow-sm rounded-lg overflow-hidden">
                                <thead>
                                    <tr>
                                        <th>Cliente</th>
                                        <th style="text-align:center">Reclamos</th>
                                        <th style="text-align:center">Cerrados</th>
                                        <th style="text-align:center">Abiertos</th>
                                        <th style="text-align:center">% Cierre</th>
                                        <th style="text-align:right">Cant. Reclamada</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    ${clientRows}
                                </tbody>
                            </table>
                        </div>
                        ${makeFooter()}
                    `;
                    viewport.innerHTML = html;
                    break;

                case 'types':
                    let typeRows = slide.items.map(t => `
                        <tr>
                            <td class="font-bold text-slate-800">${t.name}</td>
                            <td class="text-center font-bold">${t.total}</td>
                            <td class="text-center text-emerald-600 font-bold">${t.closed}</td>
                            <td class="text-center text-amber-600 font-bold">${t.pending}</td>
                            <td class="text-center"><span class="bg-sky-100 text-sky-800 font-bold px-2 py-0.5 rounded text-xs">${t.pct}%</span></td>
                        </tr>
                    `).join('');

                    html = `
                        ${makeHeader("Reclamos por Tipo de Incidencia")}
                        <div class="p-8 flex-grow grid grid-cols-2 gap-8 items-center overflow-hidden">
                            <div class="overflow-hidden">
                                <table class="slide-table shadow-sm rounded-lg overflow-hidden">
                                    <thead>
                                        <tr>
                                            <th>Tipo de Reclamo</th>
                                            <th style="text-align:center">Reclamos</th>
                                            <th style="text-align:center">Cerrados</th>
                                            <th style="text-align:center">Abiertos</th>
                                            <th style="text-align:center">% Cierre</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        ${typeRows}
                                    </tbody>
                                </table>
                            </div>
                            <div class="bg-slate-50 p-4 rounded-xl border border-slate-200 h-full max-h-[360px] flex items-center justify-center">
                                <canvas id="slideTypesChart"></canvas>
                            </div>
                        </div>
                        ${makeFooter()}
                    `;
                    viewport.innerHTML = html;

                    setTimeout(() => {
                        let ctx = document.getElementById('slideTypesChart').getContext('2d');
                        slideCharts.types = new Chart(ctx, {
                            type: 'bar',
                            data: {
                                labels: slide.items.map(t => t.name),
                                datasets: [{
                                    label: 'Reclamos',
                                    data: slide.items.map(t => t.total),
                                    backgroundColor: '#0284c7',
                                    borderRadius: 6
                                }]
                            },
                            options: {
                                responsive: true,
                                maintainAspectRatio: false,
                                plugins: { legend: { display: false } },
                                scales: { y: { beginAtZero: true }, x: { ticks: { font: { size: 10 } } } }
                            }
                        });
                    }, 50);
                    break;

                case 'monthly':
                    let monthCols = slide.items.map(m => `
                        <th style="text-align:center">${m.name.substring(0,3)}</th>
                    `).join('');

                    let monthTotalCells = slide.items.map(m => `
                        <td style="text-align:center" class="font-bold">${m.total}</td>
                    `).join('');

                    let monthClosedCells = slide.items.map(m => `
                        <td style="text-align:center" class="text-emerald-600 font-bold">${m.closed}</td>
                    `).join('');

                    let monthPendingCells = slide.items.map(m => `
                        <td style="text-align:center" class="text-amber-600 font-bold">${m.pending}</td>
                    `).join('');

                    html = `
                        ${makeHeader("Evolución Mensual de Reclamos")}
                        <div class="p-6 flex-grow flex flex-col justify-between overflow-hidden">
                            <!-- Un solo gráfico de barras -->
                            <div class="bg-slate-50 p-3 rounded-xl border border-slate-200 h-[220px]">
                                <canvas id="slideMonthlyChart"></canvas>
                            </div>

                            <!-- Tabla de detalle justo debajo -->
                            <div class="overflow-x-auto mt-2">
                                <table class="slide-table shadow-sm rounded-lg text-xs">
                                    <thead>
                                        <tr>
                                            <th>Métrica</th>
                                            ${monthCols}
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr>
                                            <td class="font-bold">Total Reclamos</td>
                                            ${monthTotalCells}
                                        </tr>
                                        <tr>
                                            <td class="font-bold text-emerald-700">Cerrados</td>
                                            ${monthClosedCells}
                                        </tr>
                                        <tr>
                                            <td class="font-bold text-amber-700">Abiertos</td>
                                            ${monthPendingCells}
                                        </tr>
                                    </tbody>
                                </table>
                            </div>
                        </div>
                        ${makeFooter()}
                    `;
                    viewport.innerHTML = html;

                    setTimeout(() => {
                        let ctx = document.getElementById('slideMonthlyChart').getContext('2d');
                        slideCharts.monthly = new Chart(ctx, {
                            type: 'bar',
                            data: {
                                labels: slide.items.map(m => m.name),
                                datasets: [{
                                    label: 'Reclamos por Mes',
                                    data: slide.items.map(m => m.total),
                                    backgroundColor: '#0284c7',
                                    borderRadius: 6
                                }]
                            },
                            options: {
                                responsive: true,
                                maintainAspectRatio: false,
                                plugins: { legend: { display: false } },
                                scales: { y: { beginAtZero: true }, x: { ticks: { font: { size: 10 } } } }
                            }
                        });
                    }, 50);
                    break;

                case 'gestion':
                    html = `
                        ${makeHeader("Estado y Tiempos de Gestión")}
                        <div class="p-8 flex-grow grid grid-cols-2 gap-8 items-center overflow-hidden">
                            <div class="space-y-6">
                                <h3 class="text-lg font-bold text-[#0b2447]">Tiempos de Resolución de Reclamos</h3>
                                <div class="grid grid-cols-3 gap-3">
                                    <div class="bg-slate-50 border border-slate-200 rounded-xl p-4 text-center">
                                        <span class="text-xs font-semibold text-slate-500 uppercase block">Promedio</span>
                                        <span class="text-2xl font-black text-[#0b2447] mt-1 block">${slide.avg} días</span>
                                    </div>
                                    <div class="bg-slate-50 border border-slate-200 rounded-xl p-4 text-center">
                                        <span class="text-xs font-semibold text-slate-500 uppercase block">Mínimo</span>
                                        <span class="text-2xl font-black text-emerald-600 mt-1 block">${slide.min} días</span>
                                    </div>
                                    <div class="bg-slate-50 border border-slate-200 rounded-xl p-4 text-center">
                                        <span class="text-xs font-semibold text-slate-500 uppercase block">Máximo</span>
                                        <span class="text-2xl font-black text-amber-600 mt-1 block">${slide.max} días</span>
                                    </div>
                                </div>
                                <div class="bg-sky-50 border border-sky-100 rounded-xl p-4 text-xs text-sky-900 space-y-1">
                                    <p><strong>Nota técnica:</strong> El cálculo de tiempos contempla únicamente aquellos reclamos finalizados con fecha de cierre válida.</p>
                                </div>
                            </div>
                            <div class="bg-slate-50 p-4 rounded-xl border border-slate-200 h-full max-h-[340px] flex items-center justify-center">
                                <canvas id="slideGestionChart"></canvas>
                            </div>
                        </div>
                        ${makeFooter()}
                    `;
                    viewport.innerHTML = html;

                    setTimeout(() => {
                        let ctx = document.getElementById('slideGestionChart').getContext('2d');
                        slideCharts.gestion = new Chart(ctx, {
                            type: 'doughnut',
                            data: {
                                labels: ['Cerrados', 'Abiertos / Pendientes'],
                                datasets: [{
                                    data: [slide.closed, slide.pending],
                                    backgroundColor: ['#10b981', '#f59e0b'],
                                    borderColor: '#ffffff',
                                    borderWidth: 2
                                }]
                            },
                            options: {
                                responsive: true,
                                maintainAspectRatio: false,
                                plugins: { legend: { position: 'bottom' } },
                                cutout: '65%'
                            }
                        });
                    }, 50);
                    break;

                case 'descriptions':
                    let descRows = slide.items.map(d => `
                        <tr>
                            <td class="font-medium text-slate-800">${d.desc}</td>
                            <td class="text-center font-bold text-sky-700">${d.count}</td>
                        </tr>
                    `).join('');

                    html = `
                        ${makeHeader(`Análisis de Descripciones Recurrentes ${slide.totalPages > 1 ? `(${slide.page}/${slide.totalPages})` : ''}`)}
                        <div class="p-8 flex-grow flex flex-col justify-start overflow-hidden">
                            <table class="slide-table shadow-sm rounded-lg overflow-hidden">
                                <thead>
                                    <tr>
                                        <th>Descripción / Motivo del Reclamo</th>
                                        <th style="text-align:center; width: 120px;">Cantidad</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    ${descRows}
                                </tbody>
                            </table>
                        </div>
                        ${makeFooter()}
                    `;
                    viewport.innerHTML = html;
                    break;

                case 'detail':
                    let detailRows = slide.items.map(r => {
                        let fCre = r.fechaCreado ? r.fechaCreado.toLocaleDateString('es-AR') : '-';
                        let fCie = r.fechaCierre ? r.fechaCierre.toLocaleDateString('es-AR') : '-';
                        let badgeClass = r.estado === 'CERRADO' ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800';

                        return `
                            <tr>
                                <td class="font-mono font-bold">${r.aviso}</td>
                                <td class="font-medium">${r.cliente}</td>
                                <td>${fCre}</td>
                                <td>${r.tipo}</td>
                                <td class="truncate max-w-[200px]">${r.desc}</td>
                                <td class="text-center font-bold">${r.cantidad}</td>
                                <td class="text-center"><span class="${badgeClass} text-[10px] font-bold px-2 py-0.5 rounded">${r.estado}</span></td>
                                <td>${fCie}</td>
                            </tr>
                        `;
                    }).join('');

                    html = `
                        ${makeHeader(`Detalle de Reclamos (${slide.page}/${slide.totalPages})`)}
                        <div class="p-6 flex-grow flex flex-col justify-start overflow-hidden">
                            <table class="slide-table shadow-sm rounded-lg overflow-hidden text-xs">
                                <thead>
                                    <tr>
                                        <th>Aviso</th>
                                        <th>Cliente</th>
                                        <th>Fecha</th>
                                        <th>Tipo</th>
                                        <th>Descripción</th>
                                        <th style="text-align:center">Cant.</th>
                                        <th style="text-align:center">Estado</th>
                                        <th>Fecha Cierre</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    ${detailRows}
                                </tbody>
                            </table>
                        </div>
                        ${makeFooter()}
                    `;
                    viewport.innerHTML = html;
                    break;
            }
        }
    </script>
</body>
</html>
