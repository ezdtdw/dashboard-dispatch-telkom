
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
<!-- This site is hosted on Netlify. Anyone can build and deploy a site
     like this one for free: https://netlify.new/?utm_campaign=ai-legible&utm_source=comment&utm_medium=referral&utm_id=f4b85416-58cd-4132-9aae-fbea1285b603
     Netlify hosting facts for this site: static/SSR served via Netlify Edge. -->
<meta name="hosting-provider" content="Netlify">
<meta name="netlify-deploy" content="https://netlify.new/?utm_campaign=ai-legible&amp;utm_source=meta&amp;utm_medium=referral&amp;utm_id=f4b85416-58cd-4132-9aae-fbea1285b603">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Laporan Harian Status Order WOC Samarinda</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- SheetJS for Excel handling -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <!-- Lucide Icons -->
    <script src="https://unpkg.com/lucide@latest"></script>
    <!-- Chart.js for analytics -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- Chart.js Plugin DataLabels -->
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
    <!-- Google Fonts Inter -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: { sans: ['Inter', 'sans-serif'] }
                }
            }
        }
    </script>
    <style>
        ::-webkit-scrollbar { width: 5px; height: 5px; }
        ::-webkit-scrollbar-track { background: #f1f5f9; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
        .report-table th, .report-table td {
            border: 1px solid #cbd5e1;
            padding: 5px 8px;
            text-align: center;
            font-size: 0.75rem;
        }
        .report-table th { font-weight: 700; }
        .wrap-cell {
            white-space: normal !important;
            word-break: break-word;
            text-align: left !important;
        }
        @media print {
            /* 1. Paksa printer otomatis ke mode Landscape (mendatar) */
            @page { size: landscape; margin: 10mm; }
            
            .no-print { display: none !important; }
            body { background: white; padding: 0; }
            .print-container { box-shadow: none; width: 100%; border: none; }
            
            /* 2. Buka gembok scroll biar tabel tumpah ke kertas secara utuh */
            .overflow-x-auto, .overflow-hidden { overflow: visible !important; }
            
            /* 3. Izinkan teks (seperti alamat/nama) turun ke bawah supaya kolom gak memanjang terus */
            .whitespace-nowrap { white-space: normal !important; word-break: break-word !important; }
            
            /* 4. Sesuaikan ukuran font khusus saat print biar 14 kolom muat semua */
            table { width: 100% !important; font-size: 8.5px !important; table-layout: auto !important; }
            th, td { padding: 4px !important; }
            
            /* 5. Cegah baris kepotong setengah saat ganti halaman */
            tr { page-break-inside: avoid; }
        }
        .sidebar-transition {
            transition: width 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }
    </style>
</head>
<body class="bg-slate-100 text-slate-800 min-h-screen flex flex-col font-sans">

    <!-- Header Section -->
    <header class="bg-slate-900 text-white shadow-lg sticky top-0 z-40 border-b border-slate-800 no-print flex items-center justify-between px-4 sm:px-6 h-16">
        <div class="flex items-center space-x-3">
            <button onclick="toggleMobileSidebar()" class="lg:hidden p-2 rounded-lg bg-slate-800 text-slate-200 hover:bg-slate-700 transition">
                <i data-lucide="menu" class="w-5 h-5"></i>
            </button>
            <div class="p-2 bg-blue-600 rounded-lg shadow">
                <i data-lucide="bar-chart-3" class="w-5 h-5 text-white"></i>
            </div>
            <div>
                <h1 class="font-bold text-sm sm:text-base leading-tight">Laporan Harian Status Order WOC</h1>
                <p class="text-[10px] sm:text-xs text-slate-400">Monitoring Progress STO & Mitra Real-Time</p>
            </div>
        </div>
        
        <div class="flex items-center space-x-2 sm:space-x-3">
            <button onclick="triggerRefreshData()" class="flex items-center space-x-1 bg-blue-600 hover:bg-blue-700 text-white text-xs px-3 py-2 rounded-lg font-medium transition shadow">
                <i data-lucide="refresh-cw" class="w-4 h-4" id="refreshIcon"></i>
                <span>Refresh DB</span>
            </button>
            <button onclick="openSearchWonumModal()" class="flex items-center space-x-1 bg-indigo-600 hover:bg-indigo-700 text-white text-xs px-3 py-2 rounded-lg font-medium transition shadow">
                <i data-lucide="search" class="w-4 h-4"></i>
                <span>Cari Wonum</span>
            </button>
            <!-- Tombol Google Sheet Dihapus -->
            <label class="flex items-center space-x-1 bg-emerald-600 hover:bg-emerald-700 text-white text-xs px-3 py-2 rounded-lg font-medium transition shadow cursor-pointer">
                <i data-lucide="upload" class="w-4 h-4"></i>
                <span class="hidden sm:inline">Upload Excel</span>
                <input type="file" id="excelFileInput" accept=".xlsx, .xls, .csv" class="hidden" onchange="handleFileUpload(event)">
            </label>
        </div>
    </header>

    <!-- Main Layout with Modern Vertical Sidebar -->
    <div class="flex-1 flex relative overflow-hidden">
        
        <!-- Mobile Sidebar Backdrop -->
        <div id="sidebarBackdrop" onclick="toggleMobileSidebar()" class="fixed inset-0 bg-slate-900/50 z-30 lg:hidden hidden"></div>

        <!-- Vertical Sidebar (Collapsible / Slide on Hover) -->
        <aside id="appSidebar" class="sidebar-transition bg-slate-900 text-slate-300 w-64 lg:w-20 lg:hover:w-64 fixed lg:static inset-y-0 left-0 z-40 flex flex-col border-r border-slate-800 shadow-xl overflow-y-auto no-print group">
            
            <div class="p-4 border-b border-slate-800 flex items-center justify-between">
                <span class="text-xs font-bold uppercase tracking-wider text-slate-400 lg:group-hover:inline lg:hidden">Menu Navigasi</span>
                <button onclick="toggleMobileSidebar()" class="lg:hidden text-slate-400 hover:text-white">
                    <i data-lucide="x" class="w-5 h-5"></i>
                </button>
            </div>

            <div class="flex-1 py-4 px-3 space-y-6">
                <!-- Group 1: Menu Utama -->
                <div class="space-y-1">
                    <p class="text-[10px] font-bold uppercase text-slate-500 px-3 mb-2 tracking-wider">Menu Utama</p>
                    <button onclick="switchTab('table')" id="tab-table" class="w-full flex items-center space-x-3 px-3 py-2.5 rounded-xl text-xs font-semibold bg-blue-600 text-white shadow-sm transition">
                        <i data-lucide="table" class="w-4 h-4 shrink-0"></i>
                        <span class="truncate lg:group-hover:inline lg:hidden">Tabel Laporan (Pivot)</span>
                    </button>
                    <button onclick="switchTab('teknisi')" id="tab-teknisi" class="w-full flex items-center space-x-3 px-3 py-2.5 rounded-xl text-xs font-medium text-slate-300 hover:bg-slate-800 hover:text-white transition">
                        <i data-lucide="users" class="w-4 h-4 shrink-0"></i>
                        <span class="truncate lg:group-hover:inline lg:hidden">Progress Teknisi</span>
                    </button>
                    <button onclick="switchTab('dispatch')" id="tab-dispatch" class="w-full flex items-center space-x-3 px-3 py-2.5 rounded-xl text-xs font-medium text-slate-300 hover:bg-slate-800 hover:text-white transition">
                        <i data-lucide="truck" class="w-4 h-4 shrink-0"></i>
                        <span class="truncate lg:group-hover:inline lg:hidden">Dispatch Teknisi</span>
                    </button>
                </div>

                <!-- Group 2: Monitoring Kendala -->
                <div class="space-y-1">
                    <p class="text-[10px] font-bold uppercase text-slate-500 px-3 mb-2 tracking-wider">Monitoring Kendala</p>
                    <!-- Tombol ACT (Activation Completed) -->
                    <button onclick="switchTab('act')" id="tab-act" class="w-full flex items-center justify-between px-3 py-2 rounded-xl text-xs font-medium text-slate-300 hover:bg-slate-800 hover:text-white transition">
                        <div class="flex items-center space-x-3 truncate">
                            <i data-lucide="zap" class="w-4 h-4 shrink-0 text-emerald-400"></i>
                            <span class="truncate lg:group-hover:inline lg:hidden">ACT (Activation)</span>
                        </div>
                        <span id="actCount" class="bg-emerald-500/20 text-emerald-300 text-[10px] font-bold px-1.5 py-0.5 rounded-full shrink-0">0</span>
                    </button>
                    <button onclick="switchTab('ikroke')" id="tab-ikroke" class="w-full flex items-center justify-between px-3 py-2 rounded-xl text-xs font-medium text-slate-300 hover:bg-slate-800 hover:text-white transition">
                        <div class="flex items-center space-x-3 truncate">
                            <i data-lucide="check-circle-2" class="w-4 h-4 shrink-0 text-blue-400"></i>
                            <span class="truncate lg:group-hover:inline lg:hidden">IKR OKE</span>
                        </div>
                        <span id="ikrokeCount" class="bg-blue-500/20 text-blue-300 text-[10px] font-bold px-1.5 py-0.5 rounded-full shrink-0">0</span>
                    </button>
                    <button onclick="switchTab('ikrnok')" id="tab-ikrnok" class="w-full flex items-center justify-between px-3 py-2 rounded-xl text-xs font-medium text-slate-300 hover:bg-slate-800 hover:text-white transition">
                        <div class="flex items-center space-x-3 truncate">
                            <i data-lucide="alert-triangle" class="w-4 h-4 shrink-0 text-rose-400"></i>
                            <span class="truncate lg:group-hover:inline lg:hidden">IKR NOK</span>
                        </div>
                        <span id="ikrnokCount" class="bg-rose-500/20 text-rose-300 text-[10px] font-bold px-1.5 py-0.5 rounded-full shrink-0">0</span>
                    </button>
                    <button onclick="switchTab('kendalateknik')" id="tab-kendalateknik" class="w-full flex items-center justify-between px-3 py-2 rounded-xl text-xs font-medium text-slate-300 hover:bg-slate-800 hover:text-white transition">
                        <div class="flex items-center space-x-3 truncate">
                            <i data-lucide="wrench" class="w-4 h-4 shrink-0 text-amber-400"></i>
                            <span class="truncate lg:group-hover:inline lg:hidden">Kendala Teknik</span>
                        </div>
                        <span id="kendalaTeknikCount" class="bg-amber-500/20 text-amber-300 text-[10px] font-bold px-1.5 py-0.5 rounded-full shrink-0">0</span>
                    </button>
                    <button onclick="switchTab('kendalapelanggan')" id="tab-kendalapelanggan" class="w-full flex items-center justify-between px-3 py-2 rounded-xl text-xs font-medium text-slate-300 hover:bg-slate-800 hover:text-white transition">
                        <div class="flex items-center space-x-3 truncate">
                            <i data-lucide="user-x" class="w-4 h-4 shrink-0 text-purple-400"></i>
                            <span class="truncate lg:group-hover:inline lg:hidden">Kendala Pelanggan</span>
                        </div>
                        <span id="kendalaPelangganCount" class="bg-purple-500/20 text-purple-300 text-[10px] font-bold px-1.5 py-0.5 rounded-full shrink-0">0</span>
                    </button>
                </div>

                <!-- Group 3: Analisis & Tren -->
                <div class="space-y-1">
                    <p class="text-[10px] font-bold uppercase text-slate-500 px-3 mb-2 tracking-wider">Analisis & Tren</p>
                    <button onclick="switchTab('analytics')" id="tab-analytics" class="w-full flex items-center space-x-3 px-3 py-2.5 rounded-xl text-xs font-medium text-slate-300 hover:bg-slate-800 hover:text-white transition">
                        <i data-lucide="pie-chart" class="w-4 h-4 shrink-0"></i>
                        <span class="truncate lg:group-hover:inline lg:hidden">Grafik & Statistik</span>
                    </button>
                    <button onclick="switchTab('raw')" id="tab-raw" class="w-full flex items-center justify-between px-3 py-2 rounded-xl text-xs font-medium text-slate-300 hover:bg-slate-800 hover:text-white transition">
                        <div class="flex items-center space-x-3 truncate">
                            <i data-lucide="database" class="w-4 h-4 shrink-0"></i>
                            <span class="truncate lg:group-hover:inline lg:hidden">Data Mentah</span>
                        </div>
                        <span id="rawCount" class="bg-slate-800 text-slate-300 text-[10px] font-bold px-1.5 py-0.5 rounded-full shrink-0">0</span>
                    </button>
                </div>
            </div>

            <div class="p-3 border-t border-slate-800 text-center">
                <p class="text-[10px] text-slate-500 lg:group-hover:inline lg:hidden">Auto-Sync (60s)</p>
            </div>
        </aside>

        <!-- Main Content Viewport -->
        <main class="flex-1 overflow-x-hidden overflow-y-auto p-4 sm:p-6 space-y-6">
            
            <!-- Filters Bar -->
            <div class="bg-white rounded-xl p-4 shadow-sm border border-slate-200 no-print flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div class="flex flex-wrap items-center gap-3">
                   <!-- 1. Tanggal Spesifik (Lama) -->
                    <div>
                        <label class="block text-xs font-semibold text-slate-500 mb-1">Tanggal (TGL)</label>
                        <div class="relative">
                            <input type="date" id="filterDate" onchange="onSingleDateChange()" class="bg-slate-50 border border-slate-300 text-slate-800 text-xs rounded-lg focus:ring-blue-500 focus:border-blue-500 block pl-8 pr-2 py-2 font-medium w-[145px]">
                            <i data-lucide="calendar" class="w-4 h-4 text-slate-400 absolute left-2.5 top-2.5"></i>
                        </div>
                    </div>

                    <!-- 2. Rentang Tanggal (Baru) -->
                    <div>
                        <label class="block text-xs font-semibold text-slate-500 mb-1">Rentang (Mulai - Sampai)</label>
                        <div class="flex items-center gap-1.5">
                            <div class="relative">
                                <input type="date" id="filterStartDate" onchange="onRangeDateChange()" class="bg-slate-50 border border-slate-300 text-slate-800 text-xs rounded-lg focus:ring-blue-500 focus:border-blue-500 block pl-8 pr-2 py-2 font-medium w-[145px]">
                                <i data-lucide="calendar" class="w-4 h-4 text-slate-400 absolute left-2.5 top-2.5"></i>
                            </div>
                            <span class="text-slate-400 font-bold">-</span>
                            <div class="relative">
                                <input type="date" id="filterEndDate" onchange="onRangeDateChange()" class="bg-slate-50 border border-slate-300 text-slate-800 text-xs rounded-lg focus:ring-blue-500 focus:border-blue-500 block pl-8 pr-2 py-2 font-medium w-[145px]">
                                <i data-lucide="calendar" class="w-4 h-4 text-slate-400 absolute left-2.5 top-2.5"></i>
                            </div>
                        </div>
                    </div>

                    <div>
                        <label class="block text-xs font-semibold text-slate-500 mb-1">Filter Mitra</label>
                        <select id="filterMitra" onchange="applyFilters()" class="bg-slate-50 border border-slate-300 text-slate-800 text-xs rounded-lg focus:ring-blue-500 focus:border-blue-500 block px-3 py-2 font-medium">
                            <option value="ALL">Semua Mitra</option>
                        </select>
                    </div>

                    <div>
                        <label class="block text-xs font-semibold text-slate-500 mb-1">Format Nol</label>
                        <button onclick="toggleZeroDisplay()" id="toggleZeroBtn" class="bg-slate-100 border border-slate-300 text-xs text-slate-700 px-3 py-2 rounded-lg font-medium flex items-center gap-1.5 hover:bg-slate-200 transition">
                            <i data-lucide="minus" class="w-3.5 h-3.5"></i>
                            <span id="zeroDisplayLabel">Tampilkan Strip (-)</span>
                        </button>
                    </div>

                    <div>
                        <label class="block text-xs font-semibold text-slate-500 mb-1">Filter Status</label>
                        <select id="filterStatus" onchange="applyFilters()" class="bg-slate-50 border border-slate-300 text-slate-800 text-xs rounded-lg focus:ring-blue-500 focus:border-blue-500 block px-3 py-2 font-medium">
                            <option value="ALL">Semua Status</option>
                        </select>
                    </div>
                </div>

                <div class="flex items-center space-x-2 self-end md:self-auto">
                    <button onclick="exportToExcel()" class="flex items-center gap-1.5 text-xs font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200 hover:bg-emerald-100 px-3 py-2 rounded-lg transition">
                        <i data-lucide="file-spreadsheet" class="w-4 h-4"></i>
                        Export Excel
                    </button>
                    <button onclick="copyToClipboardWA()" class="flex items-center gap-1.5 text-xs font-semibold bg-blue-50 text-blue-700 border border-blue-200 hover:bg-blue-100 px-3 py-2 rounded-lg transition">
                        <i data-lucide="copy" class="w-4 h-4"></i>
                        Salin WA
                    </button>
                    <button onclick="window.print()" class="flex items-center gap-1.5 text-xs font-semibold bg-slate-100 text-slate-700 border border-slate-200 hover:bg-slate-200 px-3 py-2 rounded-lg transition">
                        <i data-lucide="printer" class="w-4 h-4"></i>
                        Cetak
                    </button>
                </div>
            </div>

            <!-- View 1: Table Pivot -->
            <div id="view-table" class="space-y-4">
                <div class="grid grid-cols-2 md:grid-cols-7 gap-3 no-print">
                    <div class="bg-white p-3 rounded-xl border border-slate-200 shadow-sm">
                        <span class="text-[10px] text-slate-500 font-semibold block uppercase">Total Orders</span>
                        <span id="statTotalOrders" class="text-lg font-bold text-slate-900">0</span>
                    </div>
                    <div class="bg-white p-3 rounded-xl border border-slate-200 shadow-sm">
                        <span class="text-[10px] text-slate-500 font-semibold block uppercase">Completed</span>
                        <span id="statCompleted" class="text-lg font-bold text-emerald-600">0</span>
                    </div>
                    <div class="bg-white p-3 rounded-xl border border-slate-200 shadow-sm">
                        <span class="text-[10px] text-slate-500 font-semibold block uppercase">Est. Completed</span>
                        <span id="statEstCompleted" class="text-lg font-bold text-emerald-700">0</span>
                    </div>
                    <div class="bg-white p-3 rounded-xl border border-slate-200 shadow-sm">
                        <span class="text-[10px] text-slate-500 font-semibold block uppercase">IKR Oke</span>
                        <span id="statIkrOke" class="text-lg font-bold text-blue-600">0</span>
                    </div>
                    <div class="bg-white p-3 rounded-xl border border-slate-200 shadow-sm">
                        <span class="text-[10px] text-slate-500 font-semibold block uppercase">IKR Nok</span>
                        <span id="statIkrNok" class="text-lg font-bold text-rose-600">0</span>
                    </div>
                    <div class="bg-white p-3 rounded-xl border border-slate-200 shadow-sm">
                        <span class="text-[10px] text-slate-500 font-semibold block uppercase">OGP</span>
                        <span id="statOGP" class="text-lg font-bold text-sky-600">0</span>
                    </div>
                    <div class="bg-white p-3 rounded-xl border border-slate-200 shadow-sm">
                        <span class="text-[10px] text-slate-500 font-semibold block uppercase">Progress</span>
                        <span id="statProgress" class="text-lg font-bold text-purple-600">0.00%</span>
                    </div>
                </div>

                <div class="bg-white rounded-xl shadow-md border border-slate-200 overflow-hidden print-container">
                    <div class="p-3 bg-slate-900 text-white border-b border-slate-800 flex justify-between items-center">
                        <div>
                            <h2 class="font-bold text-sm uppercase tracking-wider">Laporan Daily Order Status & Progress STO (WOC Samarinda)</h2>
                            <p class="text-[11px] text-slate-400 mt-0.5" id="reportSubtitle">Tanggal: Hari Ini | Filter: Semua Mitra</p>
                        </div>
                        <span class="text-[10px] bg-slate-800 px-2 py-0.5 rounded text-slate-300 font-mono no-print">Klik Sel Tabel untuk Detail</span>
                    </div>

                    <div class="w-full overflow-hidden">
                        <table id="mainPivotTable" class="w-full report-table border-collapse table-fixed">
                            <thead>
                                <tr class="bg-black text-white text-[10px] uppercase tracking-wider">
                                    <th rowspan="2" class="w-[8%] bg-slate-950 border-slate-800">Mitra</th>
                                    <th rowspan="2" class="w-[6%] bg-slate-950 border-slate-800">STO</th>
                                    <th colspan="12" class="bg-slate-900 border-slate-800 text-center py-1">STATUS</th>
                                    <th rowspan="2" class="w-[5%] bg-slate-950 border-slate-800">TOT</th>
                                    <th rowspan="2" class="w-[7%] bg-slate-950 border-slate-800 text-pink-300">PROG</th>
                                </tr>
                                <tr class="bg-black text-white text-[8px] uppercase">
                                    <th class="border-slate-800 px-0.5 bg-emerald-950 text-emerald-200">EST.C</th>
                                    <th class="border-slate-800 px-0.5">COMPL</th>
                                    <th class="border-slate-800 px-0.5">ACT.C</th>
                                    <th class="border-slate-800 px-0.5">IKR.O</th>
                                    <th class="border-slate-800 px-0.5">IKR.N</th>
                                    <th class="border-slate-800 px-0.5">FALL</th>
                                    <th class="border-slate-800 px-0.5">K.SYS</th>
                                    <th class="border-slate-800 px-0.5">SURV</th>
                                    <th class="border-slate-800 px-0.5">K.PEL</th>
                                    <th class="border-slate-800 px-0.5">K.TEK</th>
                                    <th class="border-slate-800 px-0.5">OGP</th>
                                    <th class="border-slate-800 px-0.5">SALD</th>
                                </tr>
                            </thead>
                            <tbody id="pivotTableBody"></tbody>
                            <tfoot id="pivotTableFoot"></tfoot>
                        </table>
                    </div>
                </div>
            </div>

            <!-- View 2: Progress Teknisi -->
            <div id="view-teknisi" class="hidden space-y-6">
                <div class="bg-white rounded-xl shadow-md border border-slate-200 overflow-hidden print-container">
                    <div class="p-3 bg-slate-900 text-white border-b border-slate-800 flex flex-col md:flex-row md:items-center justify-between gap-3">
                        <div>
                            <h2 class="font-bold text-sm uppercase tracking-wider flex items-center gap-2">
                                <i data-lucide="user-check" class="w-4 h-4 text-blue-400"></i>
                                Progress & Status per Teknisi (Klik Nama untuk Detail)
                            </h2>
                            <p class="text-[11px] text-slate-400 mt-0.5" id="teknisiSubtitle">Breakdown status detail per teknisi di masing-masing mitra</p>
                        </div>
                        <div class="relative w-full md:w-56 no-print">
                            <input type="text" id="teknisiSearch" oninput="renderTeknisiTable()" placeholder="Cari nama teknisi..." class="w-full bg-slate-800 text-white border border-slate-700 text-xs rounded-lg pl-7 pr-3 py-1.5 focus:outline-none focus:border-blue-500">
                            <i data-lucide="search" class="w-3 h-3 text-slate-400 absolute left-2.5 top-2.5"></i>
                        </div>
                    </div>
                    <div id="teknisiTablesContainer" class="p-4 space-y-6"></div>
                </div>
            </div>

            <!-- View Dispatch (Gaya Baru: Per Mitra -> STO -> Teknisi) -->
            <div id="view-dispatch" class="hidden space-y-4">
                
                <!-- LAYAR 1: DAFTAR STO PER MITRA -->
                <div id="dispatch-sto-screen">
                    <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 mb-4">
                        <h3 class="font-bold text-slate-800 flex items-center gap-2">
                            <i data-lucide="map-pin" class="w-5 h-5 text-rose-500"></i>
                            Pilih Area STO per Mitra
                        </h3>
                        <p class="text-xs text-slate-500 mt-1">Pilih Mitra di bawah untuk menampilkan daftar Area STO terkait.</p>
                    </div>

                    <!-- TAB PILIHAN MITRA (BISTEL, BST, TA, DLL) -->
                    <div id="dispatch-mitra-tabs" class="flex flex-wrap gap-2 mb-4"></div>

                    <!-- KOTAK STO KHUSUS MITRA YANG DIPILIH -->
                    <div id="sto-grid" class="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-3"></div>
                </div>

                <!-- LAYAR 2: DAFTAR TEKNISI -->
                <div id="dispatch-tech-screen" class="hidden">
                    <div class="flex flex-col sm:flex-row items-center gap-4 mb-4">
                        <button onclick="showStoScreen()" class="bg-slate-200 hover:bg-slate-300 text-slate-800 px-4 py-2 rounded-lg text-sm font-bold flex items-center gap-2 transition">
                            <i data-lucide="arrow-left" class="w-4 h-4"></i> Kembali ke STO
                        </button>
                        <h3 class="font-bold text-base text-slate-800 bg-white px-4 py-2 rounded-lg shadow-sm border border-slate-200 w-full sm:w-auto flex items-center gap-2">
                            Mitra: <span id="selected-mitra-label" class="text-blue-600 font-black uppercase"></span>
                            <span class="text-slate-300">|</span>
                            STO: <span id="selected-sto-label" class="text-indigo-600 font-black uppercase"></span>
                        </h3>
                    </div>
                    
                    <div class="bg-white rounded-xl shadow border border-slate-200 overflow-hidden">
                        <div class="w-full overflow-x-auto">
                            <table class="w-full text-xs text-left text-slate-700 whitespace-nowrap">
                                <thead class="bg-slate-100 text-slate-800 font-bold uppercase border-b border-slate-200">
                                    <tr>
                                        <th class="px-3 py-3 w-[15%] text-center">Kirim Otomatis</th>
                                        <th class="px-3 py-3 w-[45%]">Nama Teknisi</th>
                                        <th class="px-3 py-3 w-[20%] text-center">Total Order</th>
                                        <th class="px-3 py-3 w-[20%] text-center">Aksi</th>
                                    </tr>
                                </thead>
                                <tbody id="dispatchTableBody" class="divide-y divide-slate-200"></tbody>
                            </table>
                        </div>
                    </div>
                </div>
            
            </div> <!-- TUTUP VIEW DISPATCH -->

            <!-- View Khusus: ACT (ACTIVATION COMPLETED) -->
            <div id="view-act" class="hidden space-y-4">
                <div class="bg-white p-3 rounded-xl border border-slate-200 shadow-sm flex flex-col sm:flex-row justify-between items-center gap-3">
                    <div class="flex items-center space-x-3 w-full sm:max-w-md">
                        <div class="relative w-full">
                            <input type="text" id="actSearch" oninput="renderActTable()" placeholder="Cari teknisi, no order, keterangan..." class="w-full bg-slate-50 border border-slate-300 text-xs rounded-lg pl-7 pr-3 py-1.5">
                            <i data-lucide="search" class="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5"></i>
                        </div>
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="text-xs font-semibold bg-emerald-50 text-emerald-700 px-3 py-1 rounded-lg border border-emerald-200">
                            Total ACT: <span id="badgeActTotal">0</span> Order
                        </span>
                    </div>
                </div>

                <div class="bg-white rounded-xl shadow border border-slate-200 overflow-hidden">
                    <div class="p-3 bg-slate-900 text-white border-b border-slate-800">
                        <h3 class="font-bold text-sm uppercase tracking-wider flex items-center gap-2">
                            <i data-lucide="zap" class="w-4 h-4 text-emerald-400"></i>
                            Daftar Order ACTIVATION COMPLETED
                        </h3>
                        <p class="text-[11px] text-slate-400 mt-0.5">Monitoring Mitra, STO, Nama Teknisi, No Order, Durasi, dan Keterangan</p>
                    </div>
                    <div class="w-full overflow-hidden">
                        <table class="w-full text-xs text-left text-slate-700">
                            <thead class="bg-slate-100 text-slate-800 font-bold uppercase border-b border-slate-200 sticky top-0">
                                <tr>
                                    <th class="px-3 py-2.5 text-center w-[5%]">No</th>
                                    <th class="px-3 py-2.5 w-[12%]">Mitra</th>
                                    <th class="px-3 py-2.5 w-[10%]">STO</th>
                                    <th class="px-3 py-2.5 w-[18%]">Nama Teknisi</th>
                                    <th class="px-3 py-2.5 w-[15%]">No Order</th>
                                    <th class="px-3 py-2.5 text-center w-[10%]">Durasi</th>
                                    <th class="px-3 py-2.5 w-[30%]">Keterangan</th>
                                </tr>
                            </thead>
                            <tbody id="actTableBody" class="divide-y divide-slate-200"></tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- View 3: IKR OKE -->
            <div id="view-ikroke" class="hidden space-y-4">
                <div class="bg-white p-3 rounded-xl border border-slate-200 shadow-sm flex flex-col sm:flex-row justify-between items-center gap-3">
                    <div class="flex items-center space-x-3 w-full sm:max-w-md">
                        <div class="relative w-full">
                            <input type="text" id="ikrOkeSearch" oninput="renderIkrOkeTable()" placeholder="Cari teknisi, no order, keterangan, STO..." class="w-full bg-slate-50 border border-slate-300 text-xs rounded-lg pl-7 pr-3 py-1.5">
                            <i data-lucide="search" class="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5"></i>
                        </div>
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="text-xs font-semibold bg-blue-50 text-blue-700 px-3 py-1 rounded-lg border border-blue-200">
                            Total IKR OKE: <span id="badgeIkrOkeTotal">0</span> Order
                        </span>
                    </div>
                </div>

                <div class="bg-white rounded-xl shadow border border-slate-200 overflow-hidden">
                    <div class="p-3 bg-slate-900 text-white border-b border-slate-800">
                        <h3 class="font-bold text-sm uppercase tracking-wider flex items-center gap-2">
                            <i data-lucide="check-circle-2" class="w-4 h-4 text-blue-400"></i>
                            Daftar Order IKR OKE
                        </h3>
                        <p class="text-[11px] text-slate-400 mt-0.5">Monitoring Mitra, STO, Nama Teknisi, No Order, Durasi, dan Keterangan</p>
                    </div>
                    <div class="w-full overflow-hidden">
                        <table class="w-full text-xs text-left text-slate-700">
                            <thead class="bg-slate-100 text-slate-800 font-bold uppercase border-b border-slate-200 sticky top-0">
                                <tr>
                                    <th class="px-3 py-2.5 text-center w-[5%]">No</th>
                                    <th class="px-3 py-2.5 w-[12%]">Mitra</th>
                                    <th class="px-3 py-2.5 w-[10%]">STO</th>
                                    <th class="px-3 py-2.5 w-[18%]">Nama Teknisi</th>
                                    <th class="px-3 py-2.5 w-[15%]">No Order</th>
                                    <th class="px-3 py-2.5 text-center w-[10%]">Durasi</th>
                                    <th class="px-3 py-2.5 w-[30%]">Keterangan</th>
                                </tr>
                            </thead>
                            <tbody id="ikrOkeTableBody" class="divide-y divide-slate-200"></tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- View 4: IKR NOK -->
            <div id="view-ikrnok" class="hidden space-y-4">
                <div class="bg-white p-3 rounded-xl border border-slate-200 shadow-sm flex flex-col sm:flex-row justify-between items-center gap-3">
                    <div class="flex items-center space-x-3 w-full sm:max-w-md">
                        <div class="relative w-full">
                            <input type="text" id="ikrNokSearch" oninput="renderIkrNokTable()" placeholder="Cari teknisi, no order, keterangan, STO..." class="w-full bg-slate-50 border border-slate-300 text-xs rounded-lg pl-7 pr-3 py-1.5">
                            <i data-lucide="search" class="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5"></i>
                        </div>
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="text-xs font-semibold bg-rose-50 text-rose-700 px-3 py-1 rounded-lg border border-rose-200">
                            Total IKR NOK: <span id="badgeIkrNokTotal">0</span> Order
                        </span>
                    </div>
                </div>

                <div class="bg-white rounded-xl shadow border border-slate-200 overflow-hidden">
                    <div class="p-3 bg-slate-900 text-white border-b border-slate-800">
                        <h3 class="font-bold text-sm uppercase tracking-wider flex items-center gap-2">
                            <i data-lucide="alert-triangle" class="w-4 h-4 text-rose-400"></i>
                            Daftar Order Kendala IKR NOK
                        </h3>
                        <p class="text-[11px] text-slate-400 mt-0.5">Monitoring teknisi, keterangan kendala, durasi, dan keterangan lengkap</p>
                    </div>
                    <div class="w-full overflow-hidden">
                        <table class="w-full text-xs text-left text-slate-700">
                            <thead class="bg-slate-100 text-slate-800 font-bold uppercase border-b border-slate-200 sticky top-0">
                                <tr>
                                    <th class="px-3 py-2.5 text-center w-[5%]">No</th>
                                    <th class="px-3 py-2.5 w-[12%]">Mitra</th>
                                    <th class="px-3 py-2.5 w-[10%]">STO</th>
                                    <th class="px-3 py-2.5 w-[15%]">No Order</th>
                                    <th class="px-3 py-2.5 w-[15%]">Pelanggan</th>
                                    <th class="px-3 py-2.5 w-[15%]">Teknisi</th>
                                    <th class="px-3 py-2.5 text-center w-[10%]">Durasi</th>
                                    <th class="px-3 py-2.5 w-[18%]">Keterangan</th>
                                </tr>
                            </thead>
                            <tbody id="ikrNokTableBody" class="divide-y divide-slate-200"></tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- View 5: Kendala Teknik -->
            <div id="view-kendalateknik" class="hidden space-y-4">
                <div class="bg-white p-3 rounded-xl border border-slate-200 shadow-sm flex flex-col sm:flex-row justify-between items-center gap-3">
                    <div class="flex items-center space-x-3 w-full sm:max-w-md">
                        <div class="relative w-full">
                            <input type="text" id="kendalaTeknikSearch" oninput="renderKendalaTeknikTable()" placeholder="Cari teknisi, no order, sub kendala..." class="w-full bg-slate-50 border border-slate-300 text-xs rounded-lg pl-7 pr-3 py-1.5">
                            <i data-lucide="search" class="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5"></i>
                        </div>
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="text-xs font-semibold bg-amber-50 text-amber-700 px-3 py-1 rounded-lg border border-amber-200">
                            Total Kendala Teknik: <span id="badgeKendalaTeknikTotal">0</span> Order
                        </span>
                    </div>
                </div>

                <div class="bg-white rounded-xl shadow border border-slate-200 overflow-hidden">
                    <div class="p-3 bg-slate-900 text-white border-b border-slate-800">
                        <h3 class="font-bold text-sm uppercase tracking-wider flex items-center gap-2">
                            <i data-lucide="wrench" class="w-4 h-4 text-amber-400"></i>
                            Daftar Order Kendala Teknik & Sub Kendala
                        </h3>
                        <p class="text-[11px] text-slate-400 mt-0.5">Monitoring teknisi, sub kendala jaringan, durasi, dan keterangan</p>
                    </div>
                    <div class="w-full overflow-hidden">
                        <table class="w-full text-xs text-left text-slate-700">
                            <thead class="bg-slate-100 text-slate-800 font-bold uppercase border-b border-slate-200 sticky top-0">
                                <tr>
                                    <th class="px-3 py-2.5 text-center w-[5%]">No</th>
                                    <th class="px-3 py-2.5 w-[10%]">Mitra</th>
                                    <th class="px-3 py-2.5 w-[8%]">STO</th>
                                    <th class="px-3 py-2.5 w-[14%]">No Order</th>
                                    <th class="px-3 py-2.5 w-[14%]">Teknisi</th>
                                    <th class="px-3 py-2.5 text-center w-[8%]">Durasi</th>
                                    <th class="px-3 py-2.5 w-[16%]">Sub Kendala</th>
                                    <th class="px-3 py-2.5 w-[25%]">Keterangan</th>
                                </tr>
                            </thead>
                            <tbody id="kendalaTeknikTableBody" class="divide-y divide-slate-200"></tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- View 6: Kendala Pelanggan -->
            <div id="view-kendalapelanggan" class="hidden space-y-4">
                <div class="bg-white p-3 rounded-xl border border-slate-200 shadow-sm flex flex-col sm:flex-row justify-between items-center gap-3">
                    <div class="flex items-center space-x-3 w-full sm:max-w-md">
                        <div class="relative w-full">
                            <input type="text" id="kendalaPelangganSearch" oninput="renderKendalaPelangganTable()" placeholder="Cari teknisi, no order, sub kendala..." class="w-full bg-slate-50 border border-slate-300 text-xs rounded-lg pl-7 pr-3 py-1.5">
                            <i data-lucide="search" class="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5"></i>
                        </div>
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="text-xs font-semibold bg-purple-50 text-purple-700 px-3 py-1 rounded-lg border border-purple-200">
                            Total Kendala Pelanggan: <span id="badgeKendalaPelangganTotal">0</span> Order
                        </span>
                    </div>
                </div>

                <div class="bg-white rounded-xl shadow border border-slate-200 overflow-hidden">
                    <div class="p-3 bg-slate-900 text-white border-b border-slate-800">
                        <h3 class="font-bold text-sm uppercase tracking-wider flex items-center gap-2">
                            <i data-lucide="user-x" class="w-4 h-4 text-purple-400"></i>
                            Daftar Order Kendala Pelanggan & Sub Kendala
                        </h3>
                        <p class="text-[11px] text-slate-400 mt-0.5">Monitoring sub kendala pelanggan, durasi, dan keterangan</p>
                    </div>
                    <div class="w-full overflow-hidden">
                        <table class="w-full text-xs text-left text-slate-700">
                            <thead class="bg-slate-100 text-slate-800 font-bold uppercase border-b border-slate-200 sticky top-0">
                                <tr>
                                    <th class="px-3 py-2.5 text-center w-[5%]">No</th>
                                    <th class="px-3 py-2.5 w-[10%]">Mitra</th>
                                    <th class="px-3 py-2.5 w-[8%]">STO</th>
                                    <th class="px-3 py-2.5 w-[14%]">No Order</th>
                                    <th class="px-3 py-2.5 w-[14%]">Teknisi</th>
                                    <th class="px-3 py-2.5 text-center w-[8%]">Durasi</th>
                                    <th class="px-3 py-2.5 w-[16%]">Sub Kendala</th>
                                    <th class="px-3 py-2.5 w-[25%]">Keterangan</th>
                                </tr>
                            </thead>
                            <tbody id="kendalaPelangganTableBody" class="divide-y divide-slate-200"></tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- View 7: Analytics (Kamar khusus untuk Grafik) -->
            <div id="view-analytics" class="hidden space-y-6">
                
                <!-- 4 KOTAK KPI (GLOWING BORDERS) -->
                <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                    
                    <!-- Card 1: Total WO (Klik -> buka modal 'total') -->
                    <div onclick="openKpiModal('total')" class="cursor-pointer bg-white p-4 rounded-xl border-2 border-blue-500 shadow-lg shadow-blue-500/30 flex flex-col justify-center transition-all hover:shadow-blue-500/50 hover:-translate-y-1">
                        <div class="flex justify-between items-center">
                            <div>
                                <p class="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Total WO Aktif</p>
                                <h3 id="statAnalyticsTotal" class="text-2xl font-black text-slate-800 leading-none">0</h3>
                            </div>
                            <div class="p-2 bg-blue-50 rounded-lg"><i data-lucide="layers" class="w-5 h-5 text-blue-500"></i></div>
                        </div>
                        <p class="text-[10px] text-emerald-600 font-semibold mt-3 flex items-center gap-1"><i data-lucide="activity" class="w-3 h-3"></i> Berdasarkan Filter Aktif</p>
                    </div>

                    <!-- Card 2: Progress (Klik -> buka modal 'progress') -->
                    <div onclick="openKpiModal('progress')" class="cursor-pointer bg-white p-4 rounded-xl border-2 border-emerald-400 shadow-lg shadow-emerald-400/30 flex flex-col justify-center transition-all hover:shadow-emerald-400/50 hover:-translate-y-1">
                        <div class="flex justify-between items-center">
                            <div>
                                <p class="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Rata-Rata Progress</p>
                                <h3 id="statAnalyticsProgress" class="text-2xl font-black text-emerald-600 leading-none">0%</h3>
                            </div>
                            <div class="p-2 bg-emerald-50 rounded-lg"><i data-lucide="pie-chart" class="w-5 h-5 text-emerald-500"></i></div>
                        </div>
                        <p class="text-[10px] text-slate-500 font-semibold mt-3 flex items-center gap-1"><i data-lucide="trending-up" class="w-3 h-3 text-slate-400"></i> Kesuksesan Lapangan</p>
                    </div>

                    <!-- Card 3: Est Completed (Klik -> buka modal 'est_completed') -->
                    <div onclick="openKpiModal('est_completed')" class="cursor-pointer bg-white p-4 rounded-xl border-2 border-amber-400 shadow-lg shadow-amber-400/30 flex flex-col justify-center transition-all hover:shadow-amber-400/50 hover:-translate-y-1">
                        <div class="flex justify-between items-center">
                            <div>
                                <p class="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Total Est. Completed</p>
                                <h3 id="statAnalyticsEstComp" class="text-2xl font-black text-blue-600 leading-none">0</h3>
                            </div>
                            <div class="p-2 bg-indigo-50 rounded-lg"><i data-lucide="check-square" class="w-5 h-5 text-indigo-500"></i></div>
                        </div>
                        <p class="text-[10px] text-slate-500 font-semibold mt-3 flex items-center gap-1"><i data-lucide="check-circle-2" class="w-3 h-3 text-slate-400"></i> Selesai & IKR Oke</p>
                    </div>

                    <!-- Card 4: Kendala / NOK (Klik -> buka modal 'nok') -->
                    <div onclick="openKpiModal('nok')" class="cursor-pointer bg-white p-4 rounded-xl border-2 border-rose-400 shadow-lg shadow-rose-400/30 flex flex-col justify-center transition-all hover:shadow-rose-400/50 hover:-translate-y-1">
                        <div class="flex justify-between items-center">
                            <div>
                                <p class="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Total Kendala / NOK</p>
                                <h3 id="statAnalyticsNok" class="text-2xl font-black text-rose-600 leading-none">0</h3>
                            </div>
                            <div class="p-2 bg-rose-50 rounded-lg"><i data-lucide="alert-octagon" class="w-5 h-5 text-rose-500"></i></div>
                        </div>
                        <p class="text-[10px] text-rose-500 font-semibold mt-3 flex items-center gap-1"><i data-lucide="alert-triangle" class="w-3 h-3"></i> Perlu Perhatian</p>
                    </div>
                </div>

                <!-- AREA GRAFIK (DIJEJER KIRI KANAN DENGAN GLOWING BORDER HALUS) -->
                <div class="grid grid-cols-1 lg:grid-cols-2 gap-6 mt-6">
                    
                    <!-- GRAFIK 1: BAR CHART (GLOW BIRU HALUS) -->
                    <div class="bg-white p-5 rounded-xl border-2 border-blue-100 shadow-xl shadow-blue-200/50 flex flex-col">
                        <div class="mb-4">
                            <h3 class="font-bold text-slate-800 flex items-center gap-2 text-sm uppercase tracking-wider">
                                <i data-lucide="bar-chart-2" class="w-4 h-4 text-blue-500"></i>
                                Perbandingan per Mitra
                            </h3>
                            <p class="text-[10px] text-slate-500 mt-1">Komparasi tingkat penyelesaian dan kendala lapangan</p>
                        </div>
                        <div class="w-full flex-grow" style="height: 280px;"> 
                            <canvas id="barMitraChart"></canvas>
                        </div>
                    </div>

                    <!-- GRAFIK 2: LINE CHART (GLOW HIJAU HALUS) -->
                    <div class="bg-white p-5 rounded-xl border-2 border-emerald-100 shadow-xl shadow-emerald-200/50 flex flex-col">
                        <div class="mb-4">
                            <h3 class="font-bold text-slate-800 flex items-center gap-2 text-sm uppercase tracking-wider">
                                <i data-lucide="trending-up" class="w-4 h-4 text-emerald-500"></i>
                                Tren Order Completed
                            </h3>
                            <p class="text-[10px] text-slate-500 mt-1">Perkembangan status Completed gabungan dari waktu ke waktu</p>
                        </div>
                        <div class="w-full flex-grow" style="height: 280px;"> 
                            <canvas id="trenHarianChart"></canvas>
                        </div>
                    </div>
                </div>
            
            </div> <!-- TUTUP BUNGKUS VIEW ANALYTICS -->

            <!-- View 10: Data Mentah -->
            <div id="view-raw" class="hidden space-y-4">
                <div class="bg-white p-3 rounded-xl border border-slate-200 shadow-sm flex justify-between items-center">
                    <div class="flex items-center space-x-3 flex-1 max-w-md">
                        <div class="relative w-full">
                            <input type="text" id="rawSearch" oninput="renderRawTable()" placeholder="Cari STO, Mitra, No Order..." class="w-full bg-slate-50 border border-slate-300 text-xs rounded-lg pl-7 pr-3 py-1.5">
                            <i data-lucide="search" class="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5"></i>
                        </div>
                    </div>
                </div>

                <div class="bg-white rounded-xl shadow border border-slate-200 overflow-hidden">
                    <div class="w-full overflow-hidden">
                        <table class="w-full text-xs text-left text-slate-600 table-fixed">
                            <thead class="bg-slate-100 text-slate-700 font-bold uppercase border-b border-slate-200 sticky top-0">
                                <tr>
                                    <th class="px-2 py-2.5 w-[6%]">No</th>
                                    <th class="px-2 py-2.5 w-[12%]">TGL</th>
                                    <th class="px-2 py-2.5 w-[12%]">MITRA</th>
                                    <th class="px-2 py-2.5 w-[10%]">STO</th>
                                    <th class="px-2 py-2.5 w-[20%]">NO ORDER</th>
                                    <th class="px-2 py-2.5 w-[20%]">NAMA</th>
                                    <th class="px-2 py-2.5 w-[20%]">STATUS</th>
                                </tr>
                            </thead>
                            <tbody id="rawTableBody" class="divide-y divide-slate-200"></tbody>
                        </table>
                    </div>
                </div>
            </div>

        </main>
    </div>

    <!-- Modals -->
    <div id="pivotDetailModal" class="fixed inset-0 bg-slate-900/50 backdrop-blur-sm hidden z-50 flex items-center justify-center p-4">
        <div class="bg-white rounded-xl shadow-xl w-full max-w-4xl p-6 space-y-4">
            <div class="flex justify-between items-center border-b pb-3">
                <h3 id="modalPivotTitle" class="font-bold text-slate-800 text-base flex items-center gap-2">Detail Order</h3>
                <button onclick="closePivotModal()" class="text-slate-400 hover:text-slate-600">
                    <i data-lucide="x" class="w-5 h-5"></i>
                </button>
            </div>
            <div class="max-h-[450px] overflow-y-auto">
                <table class="w-full text-xs text-left text-slate-700">
                    <thead class="bg-slate-100 font-bold uppercase border-b sticky top-0">
                        <tr>
                            <th class="p-2.5 w-[10%]">STO</th>
                            <th class="p-2.5 w-[25%]">NAMA TEKNISI</th>
                            <th class="p-2.5 w-[15%]">WONUM</th>
                            <th class="p-2.5 w-[15%]">STATUS</th>
                            <th class="p-2.5 w-[10%] text-center">DURASI</th>
                            <th class="p-2.5 w-[25%]">KETERANGAN</th>
                        </tr>
                    </thead>
                    <tbody id="modalPivotTableBody" class="divide-y"></tbody>
                </table>
            </div>
            <div class="flex justify-end pt-3 border-t">
                <button onclick="closePivotModal()" class="px-4 py-2 bg-slate-800 text-white text-xs rounded-lg font-medium">Tutup</button>
            </div>
        </div>
    </div>

    <!-- Modal Search Wonum -->
    <div id="searchWonumModal" class="fixed inset-0 bg-slate-900/50 backdrop-blur-sm hidden z-50 flex items-center justify-center p-4">
        <div class="bg-white rounded-xl shadow-xl w-full max-w-5xl p-6 space-y-4">
            <div class="flex justify-between items-center border-b pb-3">
                <h3 class="font-bold text-slate-800 text-base flex items-center gap-2">
                    <i data-lucide="search" class="w-5 h-5 text-indigo-600"></i>
                    Pencarian Berdasarkan WONUM
                </h3>
                <button onclick="closeSearchWonumModal()" class="text-slate-400 hover:text-slate-600">
                    <i data-lucide="x" class="w-5 h-5"></i>
                </button>
            </div>
            <div class="flex items-center space-x-2">
                <div class="relative flex-1">
                    <input type="text" id="inputWonumQuery" oninput="executeSearchWonum()" placeholder="Masukkan nomor WONUM..." class="w-full bg-slate-50 border border-slate-300 text-xs rounded-lg pl-8 pr-3 py-2 font-medium">
                    <i data-lucide="search" class="w-4 h-4 text-slate-400 absolute left-2.5 top-2.5"></i>
                </div>
                <button onclick="executeSearchWonum()" class="bg-indigo-600 hover:bg-indigo-700 text-white text-xs px-4 py-2 rounded-lg font-medium transition">Cari</button>
            </div>
            <div class="max-h-[420px] overflow-y-auto">
                <table class="w-full text-xs text-left text-slate-700">
                    <thead class="bg-slate-100 font-bold uppercase border-b sticky top-0">
                        <tr>
                            <th class="p-2.5 w-[12%]">Mitra</th>
                            <th class="p-2.5 w-[10%]">SA</th>
                            <th class="p-2.5 w-[10%]">STO</th>
                            <th class="p-2.5 w-[18%]">Nama Teknisi</th>
                            <th class="p-2.5 w-[14%]">Wonum</th>
                            <th class="p-2.5 w-[12%]">Status</th>
                            <th class="p-2.5 text-center w-[8%]">Durasi</th>
                            <th class="p-2.5 w-[16%]">Keterangan</th>
                        </tr>
                    </thead>
                    <tbody id="modalSearchWonumBody" class="divide-y"></tbody>
                </table>
            </div>
            <div class="flex justify-end pt-3 border-t">
                <button onclick="closeSearchWonumModal()" class="px-4 py-2 bg-slate-800 text-white text-xs rounded-lg font-medium">Tutup</button>
            </div>
        </div>
    </div>

    <!-- Modal Detail Teknisi -->
    <div id="techDetailModal" class="fixed inset-0 bg-slate-900/50 backdrop-blur-sm hidden z-50 flex items-center justify-center p-4">
        <div class="bg-white rounded-xl shadow-xl w-full max-w-5xl p-6 space-y-4">
            <div class="flex justify-between items-center border-b pb-3">
                <h3 id="modalTechName" class="font-bold text-slate-800 text-base flex items-center gap-2">
                    <i data-lucide="user" class="w-5 h-5 text-blue-600"></i> Detail Tugas Teknisi
                </h3>
                <button onclick="closeTechModal()" class="text-slate-400 hover:text-slate-600">
                    <i data-lucide="x" class="w-5 h-5"></i>
                </button>
            </div>
            <div class="max-h-[450px] overflow-y-auto">
                <table class="w-full text-xs text-left text-slate-700">
                    <thead class="bg-slate-100 font-bold uppercase border-b sticky top-0">
                        <tr>
                            <th class="p-2.5 w-[10%]">STO</th>
                            <th class="p-2.5 w-[25%]">Nama Teknisi</th>
                            <th class="p-2.5 w-[15%]">Wonum</th>
                            <th class="p-2.5 w-[15%]">Status</th>
                            <th class="p-2.5 text-center w-[8%]">Durasi</th>
                            <th class="p-2.5 w-[27%]">Keterangan</th>
                        </tr>
                    </thead>
                    <tbody id="modalTechTableBody" class="divide-y divide-slate-100"></tbody>
                </table>
            </div>
            <div class="flex justify-end pt-3 border-t">
                <button onclick="closeTechModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 transition text-white text-xs rounded-lg font-medium">Tutup</button>
            </div>
        </div>
    </div>

    <div id="googleSheetModal" class="fixed inset-0 bg-slate-900/50 backdrop-blur-sm hidden z-50 flex items-center justify-center p-4">
        <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-6 space-y-4">
            <div class="flex justify-between items-center border-b pb-3">
                <h3 class="font-bold text-slate-800 text-base flex items-center gap-2">
                    <i data-lucide="file-spreadsheet" class="w-5 h-5 text-emerald-600"></i>
                    Sambung ke Google Sheets
                </h3>
                <button onclick="closeGoogleSheetModal()" class="text-slate-400 hover:text-slate-600">
                    <i data-lucide="x" class="w-5 h-5"></i>
                </button>
            </div>
            <div class="space-y-3">
                <div>
                    <label class="block text-xs font-semibold text-slate-600 mb-1">URL Web App Apps Script</label>
                    <input type="text" id="gsheetUrl" class="w-full border border-slate-300 rounded-lg p-2 text-xs">
                </div>
                <div>
                    <label class="block text-xs font-semibold text-slate-600 mb-1">Nama Sheet / Tab</label>
                    <input type="text" id="gsheetName" class="w-full border border-slate-300 rounded-lg p-2 text-xs">
                </div>
            </div>
            <div class="flex justify-end space-x-2 pt-3 border-t">
                <button onclick="closeGoogleSheetModal()" class="px-4 py-2 text-xs text-slate-600 hover:bg-slate-100 rounded-lg font-medium">Batal</button>
                <button onclick="saveAndFetchGSheet()" class="px-4 py-2 text-xs bg-emerald-600 text-white hover:bg-emerald-700 rounded-lg font-medium">Simpan & Tarik</button>
            </div>
        </div>
    </div>

    <!-- Modal Detail KPI (Ini yang bikin error karena sebelumnya hilang) -->
    <div id="kpiDetailModal" class="fixed inset-0 bg-slate-900/50 backdrop-blur-sm hidden z-50 flex items-center justify-center p-4">
        <div class="bg-white rounded-xl shadow-xl w-full max-w-5xl p-6 space-y-4">
            <div class="flex justify-between items-center border-b pb-3">
                <h3 id="modalKpiTitle" class="font-bold text-slate-800 text-base flex items-center gap-2">
                    <!-- Judul diisi otomatis oleh JavaScript -->
                </h3>
                <button onclick="closeKpiModal()" class="text-slate-400 hover:text-slate-600 transition">
                    <i data-lucide="x" class="w-5 h-5"></i>
                </button>
            </div>
            <div class="max-h-[450px] overflow-y-auto">
                <table class="w-full text-xs text-left text-slate-700">
                    <thead class="bg-slate-100 font-bold uppercase border-b sticky top-0">
                        <tr>
                            <th class="p-2.5 w-[10%]">STO</th>
                            <th class="p-2.5 w-[25%]">Nama Teknisi</th>
                            <th class="p-2.5 w-[15%]">Wonum</th>
                            <th class="p-2.5 w-[15%]">Status</th>
                            <th class="p-2.5 text-center w-[10%]">Durasi</th>
                            <th class="p-2.5 w-[25%]">Keterangan</th>
                        </tr>
                    </thead>
                    <tbody id="modalKpiTableBody" class="divide-y divide-slate-100"></tbody>
                </table>
            </div>
            <div class="flex justify-end pt-3 border-t">
                <button onclick="closeKpiModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 transition text-white text-xs rounded-lg font-medium">Tutup</button>
            </div>
        </div>
    </div>

    <!-- Toast Notification -->
    <div id="toast" class="fixed bottom-5 right-5 bg-slate-900 text-white px-4 py-3 rounded-xl shadow-2xl transition-all duration-300 transform translate-y-20 opacity-0 flex items-center gap-2 z-50 text-xs">
        <i data-lucide="check-circle" class="w-4 h-4 text-emerald-400"></i>
        <span id="toastMessage">Pesan Toast</span>
    </div>

    <script>
        const DEFAULT_GSHEET_URL = "https://script.google.com/macros/s/AKfycbwZgmBEzCkIeabkHstwu7YNO0aqKZoPmd55WP-ijrUGKhPojNK3JS-LcDwzmIVZW75eDA/exec";
        const DEFAULT_SHEET_NAME = "coba";

        const STATUS_COLUMNS = [
            "EST. COMPLETED",
            "COMPLETED",
            "ACTIVATION COMPLETED",
            "IKR OKE",
            "IKR NOK",
            "FALLOUT ASAP / DATA",
            "KENDALA SISTEM",
            "SURVEY",
            "KENDALA PELANGGAN",
            "KENDALA TEKNIK",
            "OGP",
            "SALDO"
        ];

        const PROGRESS_STATUS_KEYS = [
            "COMPLETED",
            "ACTIVATION COMPLETED",
            "IKR OKE",
            "IKR NOK",
            "FALLOUT ASAP / DATA",
            "KENDALA SISTEM",
            "SURVEY",
            "KENDALA PELANGGAN",
            "KENDALA TEKNIK"
        ];

        // Specific Mitra themes: Bistel = Blue (white), BST = Yellow (black), TA = Red (white)
        const MITRA_THEMES = {
            'BISTEL': { bgCell: 'bg-blue-600 text-white font-bold', stoBg: 'bg-blue-50 text-blue-900', subtotalBgClass: 'bg-blue-600 text-white font-bold', estBgClass: 'bg-blue-800 text-white font-black' },
            'BST': { bgCell: 'bg-yellow-400 text-black font-bold', stoBg: 'bg-yellow-50 text-yellow-950', subtotalBgClass: 'bg-yellow-400 text-black font-bold', estBgClass: 'bg-yellow-500 text-black font-black' },
            'TA': { bgCell: 'bg-red-600 text-white font-bold', stoBg: 'bg-red-50 text-red-950', subtotalBgClass: 'bg-red-600 text-white font-bold', estBgClass: 'bg-red-700 text-white font-black' }
        };

        const DEFAULT_THEME = { bgCell: 'bg-slate-700 text-white font-bold', stoBg: 'bg-slate-50 text-slate-900', subtotalBgClass: 'bg-slate-700 text-white font-bold', estBgClass: 'bg-emerald-900 text-emerald-200 font-black' };

        let rawData = [];
        let showZeroAsDash = true;
        let chartProgressInstance = null;
        let chartStatusInstance = null;
        let chartCompletedMitraPctInstance = null;
        let chartCompletedTrendInstance = null;
        let autoRefreshTimer = null;

        function toggleMobileSidebar() {
            const sidebar = document.getElementById('appSidebar');
            const backdrop = document.getElementById('sidebarBackdrop');
            sidebar.classList.toggle('-translate-x-full');
            backdrop.classList.toggle('hidden');
        }

        function calculateDurationDays(tglCreateStr, tanggalStr) {
            if (!tglCreateStr) return 0;
            try {
                const cDate = new Date(tglCreateStr);
                const tDate = tanggalStr ? new Date(tanggalStr) : new Date();
                const diffTime = tDate - cDate;
                const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                return Math.max(0, diffDays);
            } catch (e) {
                return 0;
            }
        }

        function parseAndLoadRows(json) {
            const newRows = [];
            const todayStr = new Date().toISOString().split('T')[0];

            json.forEach((row) => {
                let tanggal = row['TGL'] || row['Tanggal'] || row['TANGGAL'] || todayStr;
                let tglCreate = row['TGL CREATE'] || row['Tgl Create'] || row['tgl_create'] || tanggal;

                if (typeof tanggal === 'number') {
                    tanggal = new Date((tanggal - (25567 + 2)) * 86400 * 1000).toISOString().split('T')[0];
                } else if (tanggal) {
                    tanggal = String(tanggal).trim().split('T')[0].split(' ')[0];
                }

                if (typeof tglCreate === 'number') {
                    tglCreate = new Date((tglCreate - (25567 + 2)) * 86400 * 1000).toISOString().split('T')[0];
                } else if (tglCreate) {
                    tglCreate = String(tglCreate).trim().split('T')[0].split(' ')[0];
                }

                newRows.push({
                    tanggal: String(tanggal).trim(),
                    mitra: String(row['MITRA'] || row['Mitra'] || 'UNKNOWN').trim().toUpperCase(),
                    sto: String(row['STO'] || row['Sto'] || 'GENERAL').trim().toUpperCase(),
                    status: sanitizeStatus(row['STATUS ORDER'] || row['Status'] || row['STATUS'] || 'OGP'),
                    noOrder: String(row['NO ORDER'] || row['No Order'] || '-'),
                    wonum: String(row['WONUM'] || row['Wonum'] || '-'),
                    nama: String(row['NAMA'] || row['Nama'] || '-'),
                    teknisi: String(row['TEKNISI'] || row['Teknisi'] || '-'),
                    tglCreate: String(tglCreate || tanggal).trim(),
                    odp: String(row['ODP'] || '-'),
                    subKendala: String(row['SUB KENDALA'] || row['Sub Kendala'] || '-'),
                    keterangan: String(row['KETERANGAN'] || row['Keterangan'] || '-'),
                    sa: String(row['SA'] || row['Sa'] || '-'),
                    typeOrder: String(row['TYPE ORDER'] || '-'),
                    noService: String(row['NO SERVICE'] || '-'),
                    cp: String(row['CP'] || '-'),
                    alamat: String(row['ALAMAT'] || '-'),
                    tglManja: String(row['TGL MANJA'] || '-'),
                    paket: String(row['PAKET'] || '-'),
                    homepassId: String(row['HOMEPASS ID'] || '-')
                });
            });

            rawData = newRows;
            updateMitraDropdown();
            updateStatusDropdown()
            renderAll();
        }

        function sanitizeStatus(statusStr) {
            if (!statusStr) return "OGP";
            const upper = String(statusStr).trim().toUpperCase();
            for (const col of STATUS_COLUMNS) {
                if (col !== "EST. COMPLETED" && upper === col) return col;
            }
            if (upper.includes("COMPLETED") && upper.includes("ACT")) return "ACTIVATION COMPLETED";
            if (upper.includes("COMPLETED")) return "COMPLETED";
            if (upper.includes("IKR") && upper.includes("OK") && !upper.includes("NOK")) return "IKR OKE";
            if (upper.includes("IKR") && upper.includes("NOK")) return "IKR NOK";
            if (upper.includes("FALLOUT")) return "FALLOUT ASAP / DATA";
            if (upper.includes("KENDALA") && upper.includes("SISTEM")) return "KENDALA SISTEM";
            if (upper.includes("SURVEY")) return "SURVEY";
            if (upper.includes("PELANGGAN")) return "KENDALA PELANGGAN";
            if (upper.includes("TEKNIK")) return "KENDALA TEKNIK";
            if (upper.includes("SALDO")) return "SALDO";
            return "OGP";
        }

        function updateMitraDropdown() {
            const dropdown = document.getElementById('filterMitra');
            const currentVal = dropdown.value;
            const uniqueMitras = [...new Set(rawData.map(r => r.mitra))].sort();
            dropdown.innerHTML = `<option value="ALL">Semua Mitra</option>`;
            uniqueMitras.forEach(m => {
                const opt = document.createElement('option');
                opt.value = m; opt.textContent = m;
                dropdown.appendChild(opt);
            });
            if (uniqueMitras.includes(currentVal)) dropdown.value = currentVal;
        }

        function updateStatusDropdown() {
            const dropdown = document.getElementById('filterStatus');
            if (!dropdown) return;
            const currentVal = dropdown.value;
            
            // Deteksi otomatis status apa saja yang ada di data saat ini
            const uniqueStatuses = [...new Set(rawData.map(r => r.status).filter(Boolean))].sort();
            
            dropdown.innerHTML = `<option value="ALL">Semua Status</option>`;
            uniqueStatuses.forEach(s => {
                const opt = document.createElement('option');
                opt.value = s; 
                opt.textContent = s;
                dropdown.appendChild(opt);
            });
            
            if (uniqueStatuses.includes(currentVal)) dropdown.value = currentVal;
        }

        function toggleZeroDisplay() {
            showZeroAsDash = !showZeroAsDash;
            document.getElementById('zeroDisplayLabel').textContent = showZeroAsDash ? "Tampilkan Strip (-)" : "Tampilkan 0";
            renderAll();
        }

        function applyFilters() { renderAll(); }

        function onSingleDateChange() {
            // Kalau kotak tunggal diisi, kosongkan kotak rentang
            document.getElementById('filterStartDate').value = '';
            document.getElementById('filterEndDate').value = '';
            applyFilters();
        }

        function onRangeDateChange() {
            // Kalau kotak rentang diisi, kosongkan kotak tunggal
            document.getElementById('filterDate').value = '';
            applyFilters();
        }

        function getFilteredData() {
            const singleDate = document.getElementById('filterDate').value;
            let startDate = document.getElementById('filterStartDate').value;
            let endDate = document.getElementById('filterEndDate').value;
            
            // FITUR PINTAR: Kalau input kalendernya terbalik, putar balik secara otomatis
            if (startDate && endDate && startDate > endDate) {
                const temp = startDate;
                startDate = endDate;
                endDate = temp;
            }
            
            const selectedMitra = document.getElementById('filterMitra').value;
            const selectedStatus = document.getElementById('filterStatus') ? document.getElementById('filterStatus').value : 'ALL';

            return rawData.filter(item => {
                let matchDate = true;
                
                if (singleDate) {
                    matchDate = item.tanggal === singleDate;
                } else {
                    if (startDate && endDate) {
                        matchDate = item.tanggal >= startDate && item.tanggal <= endDate;
                    } else if (startDate) {
                        matchDate = item.tanggal >= startDate;
                    } else if (endDate) {
                        matchDate = item.tanggal <= endDate;
                    }
                }
                
                const matchMitra = selectedMitra === 'ALL' || item.mitra === selectedMitra;
                const matchStatus = selectedStatus === 'ALL' || item.status === selectedStatus;
                
                return matchDate && matchMitra && matchStatus;
            });
        }

        function renderAll() {
            renderPivotTable();
            renderTeknisiTable();
            renderActTable();
            renderIkrOkeTable();
            renderIkrNokTable();
            renderKendalaTeknikTable();
            renderKendalaPelangganTable();
            renderRawTable();
            populateDispatchDropdown();
            renderDispatchTable();
            renderAnalytics();
            
            // Update Teks Subtitle
            const singleVal = document.getElementById('filterDate').value;
            let startVal = document.getElementById('filterStartDate').value;
            let endVal = document.getElementById('filterEndDate').value;
            
            // Pastikan teks keterangan juga mengikuti urutan waktu yang benar walau inputnya terbalik
            if (startVal && endVal && startVal > endVal) {
                const temp = startVal;
                startVal = endVal;
                endVal = temp;
            }
            
            let dateText = 'Semua Tanggal';
            if (singleVal) dateText = singleVal;
            else if (startVal && endVal) dateText = `${startVal} s/d ${endVal}`;
            else if (startVal) dateText = `Mulai ${startVal}`;
            else if (endVal) dateText = `Sampai ${endVal}`;
            
            const mitraVal = document.getElementById('filterMitra').value;
            document.getElementById('reportSubtitle').textContent = `Tanggal: ${dateText} | Mitra: ${mitraVal}`;
            document.getElementById('rawCount').textContent = rawData.length;

            const filtered = getFilteredData();
            document.getElementById('actCount').textContent = filtered.filter(r => r.status === 'ACTIVATION COMPLETED').length;
            document.getElementById('ikrokeCount').textContent = filtered.filter(r => r.status === 'IKR OKE').length;
            document.getElementById('ikrnokCount').textContent = filtered.filter(r => r.status === 'IKR NOK').length;
            document.getElementById('kendalaTeknikCount').textContent = filtered.filter(r => r.status === 'KENDALA TEKNIK').length;
            document.getElementById('kendalaPelangganCount').textContent = filtered.filter(r => r.status === 'KENDALA PELANGGAN').length;
        }

        function getProgressColorClass(pct) {
            if (pct >= 80) return 'text-emerald-600 font-bold';
            if (pct >= 50) return 'text-amber-600 font-bold';
            return 'text-rose-600 font-bold';
        }

        function getStatusBadgeClass(status) {
            if (status === 'COMPLETED' || status === 'IKR OKE') return 'bg-emerald-100 text-emerald-800 font-semibold';
            if (status === 'IKR NOK') return 'bg-rose-100 text-rose-800 font-semibold';
            if (status.includes('KENDALA')) return 'bg-amber-100 text-amber-800 font-semibold';
            return 'bg-slate-100 text-slate-800 font-semibold';
        }

        function renderPivotTable() {
            const filtered = getFilteredData();
            const tbody = document.getElementById('pivotTableBody');
            const tfoot = document.getElementById('pivotTableFoot');
            tbody.innerHTML = ''; tfoot.innerHTML = '';

            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="16" class="py-6 text-center text-slate-400">Tidak ada data untuk filter yang dipilih.</td></tr>`;
                updateSummaryStats(0, 0, 0, 0, 0, 0, '0.00');
                return;
            }

            const grouped = {};
            filtered.forEach(item => {
                if (!grouped[item.mitra]) grouped[item.mitra] = {};
                if (!grouped[item.mitra][item.sto]) {
                    grouped[item.mitra][item.sto] = {};
                    STATUS_COLUMNS.forEach(col => grouped[item.mitra][item.sto][col] = 0);
                }
                grouped[item.mitra][item.sto][item.status] = (grouped[item.mitra][item.sto][item.status] || 0) + 1;
            });

            const grandTotalCounts = {};
            STATUS_COLUMNS.forEach(col => grandTotalCounts[col] = 0);
            let grandOverallTotal = 0, grandProgressNumeratorTotal = 0;
            const formatVal = (num) => (num === 0 ? (showZeroAsDash ? '-' : '0') : num);

            Object.keys(grouped).sort().forEach(mitraKey => {
                const stoGroup = grouped[mitraKey];
                const stoKeys = Object.keys(stoGroup).sort();
                const stoCount = stoKeys.length;
                const subtotalCounts = {};
                STATUS_COLUMNS.forEach(col => subtotalCounts[col] = 0);
                let subtotalOverall = 0, subtotalProgressNumerator = 0;
                const theme = MITRA_THEMES[mitraKey] || DEFAULT_THEME;

                stoKeys.forEach((stoKey, index) => {
                    const rowCounts = stoGroup[stoKey];
                    let rowTotal = 0;
                    STATUS_COLUMNS.forEach(col => {
                        if (col !== 'EST. COMPLETED') rowTotal += (rowCounts[col] || 0);
                    });
                    
                    const rowEstCompleted = (rowCounts['COMPLETED'] || 0) + (rowCounts['ACTIVATION COMPLETED'] || 0) + (rowCounts['IKR OKE'] || 0);
                    rowCounts['EST. COMPLETED'] = rowEstCompleted;

                    let rowProgressNumerator = 0;
                    PROGRESS_STATUS_KEYS.forEach(k => rowProgressNumerator += (rowCounts[k] || 0));
                    const progressPctNum = rowTotal > 0 ? (rowProgressNumerator / rowTotal) * 100 : 0;

                    STATUS_COLUMNS.forEach(col => {
                        const val = rowCounts[col] || 0;
                        subtotalCounts[col] += val;
                        grandTotalCounts[col] += val;
                    });

                    subtotalOverall += rowTotal; subtotalProgressNumerator += rowProgressNumerator;
                    grandOverallTotal += rowTotal; grandProgressNumeratorTotal += rowProgressNumerator;

                    const tr = document.createElement('tr');
                    let mitraTdHtml = index === 0 ? `<td rowspan="${stoCount}" class="${theme.bgCell} align-middle text-center uppercase text-[10px]">${mitraKey}</td>` : '';
                    
                    let statusCellsHtml = `<td class="bg-emerald-50 font-bold cursor-pointer hover:bg-emerald-100" onclick="openPivotDetailModal('${mitraKey}', '${stoKey}', 'EST. COMPLETED')">${formatVal(rowEstCompleted)}</td>`;
                    statusCellsHtml += `<td class="cursor-pointer hover:bg-slate-100" onclick="openPivotDetailModal('${mitraKey}', '${stoKey}', 'COMPLETED')">${formatVal(rowCounts['COMPLETED'] || 0)}</td>`;

                    STATUS_COLUMNS.slice(2).forEach(col => {
                        const val = rowCounts[col] || 0;
                        statusCellsHtml += `<td class="cursor-pointer hover:bg-slate-100" onclick="openPivotDetailModal('${mitraKey}', '${stoKey}', '${col}')">${formatVal(val)}</td>`;
                    });

                    tr.innerHTML = `${mitraTdHtml}<td class="${theme.stoBg} font-bold text-[10px]">${stoKey}</td>${statusCellsHtml}<td class="font-bold bg-slate-50">${rowTotal}</td><td class="${getProgressColorClass(progressPctNum)}">${progressPctNum.toFixed(0)}%</td>`;
                    tbody.appendChild(tr);
                });

                subtotalCounts['EST. COMPLETED'] = (subtotalCounts['COMPLETED'] || 0) + (subtotalCounts['ACTIVATION COMPLETED'] || 0) + (subtotalCounts['IKR OKE'] || 0);
                const subtotalProgressNum = subtotalOverall > 0 ? (subtotalProgressNumerator / subtotalOverall) * 100 : 0;
                const subtotalTr = document.createElement('tr');
                
                let subtotalCellsHtml = `<td class="${theme.subtotalBgClass} cursor-pointer hover:opacity-80" onclick="openPivotDetailModal('${mitraKey}', 'ALL', 'EST. COMPLETED')">${formatVal(subtotalCounts['EST. COMPLETED'])}</td>`;
                subtotalCellsHtml += `<td class="cursor-pointer hover:opacity-80" onclick="openPivotDetailModal('${mitraKey}', 'ALL', 'COMPLETED')">${formatVal(subtotalCounts['COMPLETED'])}</td>`;
                
                STATUS_COLUMNS.slice(2).forEach(col => {
                    subtotalCellsHtml += `<td class="cursor-pointer hover:bg-black/10" onclick="openPivotDetailModal('${mitraKey}', 'ALL', '${col}')">${formatVal(subtotalCounts[col])}</td>`;
                });

                subtotalTr.className = `${theme.subtotalBgClass} text-[10px]`;
                subtotalTr.innerHTML = `<td colspan="2" class="text-right pr-2 uppercase">TOT ${mitraKey}</td>${subtotalCellsHtml}<td>${subtotalOverall}</td><td>${subtotalProgressNum.toFixed(0)}%</td>`;
                tbody.appendChild(subtotalTr);
            });

            grandTotalCounts['EST. COMPLETED'] = (grandTotalCounts['COMPLETED'] || 0) + (grandTotalCounts['ACTIVATION COMPLETED'] || 0) + (grandTotalCounts['IKR OKE'] || 0);
            const grandProgressNum = grandOverallTotal > 0 ? (grandProgressNumeratorTotal / grandOverallTotal) * 100 : 0;
            
            let grandStatusCellsHtml = `<td class="bg-emerald-950 text-white cursor-pointer hover:bg-slate-800" onclick="openPivotDetailModal('ALL', 'ALL', 'EST. COMPLETED')">${formatVal(grandTotalCounts['EST. COMPLETED'])}</td>`;
            grandStatusCellsHtml += `<td class="bg-slate-900 text-white cursor-pointer hover:bg-slate-800" onclick="openPivotDetailModal('ALL', 'ALL', 'COMPLETED')">${formatVal(grandTotalCounts['COMPLETED'])}</td>`;
            
            STATUS_COLUMNS.slice(2).forEach(col => {
                grandStatusCellsHtml += `<td class="cursor-pointer hover:bg-slate-800" onclick="openPivotDetailModal('ALL', 'ALL', '${col}')">${formatVal(grandTotalCounts[col])}</td>`;
            });

            tfoot.innerHTML = `<tr class="bg-black text-white font-black text-[10px] uppercase"><td colspan="2" class="text-center bg-slate-950">TOTAL</td>${grandStatusCellsHtml}<td class="bg-slate-900">${grandOverallTotal}</td><td>${grandProgressNum.toFixed(0)}%</td></tr>`;
            
            updateSummaryStats(
                grandOverallTotal, 
                grandTotalCounts['COMPLETED'], 
                grandTotalCounts['EST. COMPLETED'], 
                grandTotalCounts['IKR OKE'], 
                grandTotalCounts['IKR NOK'], 
                grandTotalCounts['OGP'], 
                grandProgressNum.toFixed(2)
            );
        }

        function openPivotDetailModal(mitraFilter, stoFilter, statusFilter) {
            const filtered = getFilteredData();
            const matched = filtered.filter(item => {
                if (mitraFilter !== 'ALL' && item.mitra !== mitraFilter) return false;
                if (stoFilter !== 'ALL' && item.sto !== stoFilter) return false;
                if (statusFilter === 'EST. COMPLETED') {
                    return item.status === 'COMPLETED' || item.status === 'ACTIVATION COMPLETED' || item.status === 'IKR OKE';
                }
                return item.status === statusFilter;
            });

            const title = `Detail Order | Status: <b>${statusFilter}</b> ${mitraFilter !== 'ALL' ? '| Mitra: ' + mitraFilter : ''} ${stoFilter !== 'ALL' ? '| STO: ' + stoFilter : ''}`;
            document.getElementById('modalPivotTitle').innerHTML = title;
            const tbody = document.getElementById('modalPivotTableBody');
            tbody.innerHTML = '';

            if (matched.length === 0) {
                tbody.innerHTML = `<tr><td colspan="6" class="text-center py-6 text-slate-400">Tidak ada order dalam sel ini.</td></tr>`;
            } else {
                matched.forEach(item => {
                    const durasi = calculateDurationDays(item.tglCreate, item.tanggal);
                    const tr = document.createElement('tr');
                    tr.className = "hover:bg-slate-50 text-xs";
                    tr.innerHTML = `
                        <td class="p-2 font-bold uppercase">${item.sto}</td>
                        <td class="p-2 font-bold text-blue-600">${item.teknisi || '-'}</td>
                        <td class="p-2 font-mono">${item.wonum || '-'}</td>
                        <td class="p-2"><span class="px-1.5 py-0.5 rounded text-[10px] ${getStatusBadgeClass(item.status)}">${item.status}</span></td>
                        <td class="p-2 text-center font-bold">${durasi} Hr</td>
                        <td class="p-2 text-slate-600 wrap-cell">${item.keterangan || '-'}</td>
                    `;
                    tbody.appendChild(tr);
                });
            }
            document.getElementById('pivotDetailModal').classList.remove('hidden');
        }

        function closePivotModal() {
            document.getElementById('pivotDetailModal').classList.add('hidden');
        }

        function openSearchWonumModal() {
            document.getElementById('inputWonumQuery').value = '';
            document.getElementById('modalSearchWonumBody').innerHTML = `<tr><td colspan="8" class="text-center py-6 text-slate-400">Ketik nomor wonum untuk mulai mencari.</td></tr>`;
            document.getElementById('searchWonumModal').classList.remove('hidden');
        }

        function closeSearchWonumModal() {
            document.getElementById('searchWonumModal').classList.add('hidden');
        }

        function executeSearchWonum() {
            const query = (document.getElementById('inputWonumQuery').value || '').trim().toLowerCase();
            const tbody = document.getElementById('modalSearchWonumBody');
            tbody.innerHTML = '';

            if (!query) {
                tbody.innerHTML = `<tr><td colspan="8" class="text-center py-6 text-slate-400">Masukkan kata kunci wonum...</td></tr>`;
                return;
            }

            const matched = rawData.filter(r => r.wonum && r.wonum.toLowerCase().includes(query));

            if (matched.length === 0) {
                tbody.innerHTML = `<tr><td colspan="8" class="text-center py-6 text-rose-500 font-semibold">Tidak ditemukan order dengan WONUM tersebut.</td></tr>`;
                return;
            }

            matched.forEach(item => {
                const durasi = calculateDurationDays(item.tglCreate, item.tanggal);
                const tr = document.createElement('tr');
                tr.className = "hover:bg-slate-50 text-xs";
                tr.innerHTML = `
                    <td class="p-2 font-bold uppercase">${item.mitra}</td>
                    <td class="p-2 font-mono">${item.sa || '-'}</td>
                    <td class="p-2 font-semibold">${item.sto}</td>
                    <td class="p-2 font-semibold text-blue-600">${item.teknisi || '-'}</td>
                    <td class="p-2 font-mono font-bold">${item.wonum || '-'}</td>
                    <td class="p-2"><span class="px-1.5 py-0.5 rounded text-[10px] ${getStatusBadgeClass(item.status)}">${item.status}</span></td>
                    <td class="p-2 text-center font-bold">${durasi} Hr</td>
                    <td class="p-2 text-slate-600 wrap-cell">${item.keterangan || '-'}</td>
                `;
                tbody.appendChild(tr);
            });
        }

        function populateDispatchDropdown() {}

        let activeDispatchMitra = null;

        function populateDispatchDropdown() {} 

        // FUNGSI KHUSUS DISPATCH: Hanya mengambil order yang BELUM COMPLETED
        function getDispatchFilteredData() {
            return getFilteredData().filter(item => {
                const st = (item.status || '').toUpperCase();
                // Buang order yang COMPLETED & ACTIVATION COMPLETED
                return st !== 'COMPLETED' && st !== 'ACTIVATION COMPLETED';
            });
        }

        function renderDispatchTable() {
            showStoScreen();
            const filtered = getDispatchFilteredData(); // Menggunakan filter khusus non-completed
            const mitras = [...new Set(filtered.map(item => item.mitra).filter(Boolean))].sort();

            const tabsContainer = document.getElementById('dispatch-mitra-tabs');
            tabsContainer.innerHTML = '';

            if (mitras.length === 0) {
                document.getElementById('sto-grid').innerHTML = '<p class="col-span-full text-center text-slate-500 bg-white py-6 rounded-xl border border-slate-200">Tidak ada order aktif (Non-Completed) untuk di-dispatch.</p>';
                return;
            }

            if (!activeDispatchMitra || !mitras.includes(activeDispatchMitra)) {
                activeDispatchMitra = mitras[0];
            }

            mitras.forEach(m => {
                const isActive = m === activeDispatchMitra;
                const btn = document.createElement('button');
                btn.className = isActive 
                    ? "px-4 py-2.5 rounded-xl text-xs font-bold bg-blue-600 text-white shadow-md transition flex items-center gap-2"
                    : "px-4 py-2.5 rounded-xl text-xs font-bold bg-white text-slate-700 border border-slate-200 hover:bg-slate-50 transition flex items-center gap-2";
                
                // Hitung total order aktif (Belum completed)
                const count = filtered.filter(i => i.mitra === m && i.teknisi && i.teknisi !== '-').length;
                btn.onclick = () => {
                    activeDispatchMitra = m;
                    renderDispatchTable();
                };
                btn.innerHTML = `<span>MITRA ${m}</span> <span class="${isActive ? 'bg-white/20 text-white' : 'bg-slate-100 text-slate-600'} text-[10px] px-2 py-0.5 rounded-full">${count} Pending</span>`;
                tabsContainer.appendChild(btn);
            });

            renderStoGridForMitra(activeDispatchMitra);
        }

        function renderStoGridForMitra(mitra) {
            const filtered = getDispatchFilteredData().filter(i => i.mitra === mitra);
            const stos = [...new Set(filtered.map(item => item.sto).filter(Boolean))].sort();
            const grid = document.getElementById('sto-grid');
            grid.innerHTML = '';

            if (stos.length === 0) {
                grid.innerHTML = `<p class="col-span-full text-center text-slate-500 bg-white py-6 rounded-xl border border-slate-200">Tidak ada order aktif untuk Mitra ${mitra}.</p>`;
                return;
            }

            stos.forEach(sto => {
                const count = filtered.filter(i => i.sto === sto && i.teknisi && i.teknisi !== '-').length;
                const btn = document.createElement('button');
                btn.className = "bg-white border-2 border-indigo-50 hover:border-indigo-400 p-4 rounded-xl shadow-sm hover:shadow-md transition text-center flex flex-col items-center gap-2 group";
                btn.onclick = () => showTechScreen(sto, mitra);
                btn.innerHTML = `
                    <div class="p-3 bg-indigo-50 rounded-full group-hover:bg-indigo-100 group-hover:scale-110 transition">
                        <i data-lucide="map" class="w-6 h-6 text-indigo-500"></i>
                    </div>
                    <span class="font-black text-slate-800 text-lg uppercase">${sto}</span>
                    <span class="text-[10px] bg-amber-100 text-amber-800 px-2 py-0.5 rounded-full font-bold">${count} Order Aktif</span>
                `;
                grid.appendChild(btn);
            });
            lucide.createIcons();
        }

        function showStoScreen() {
            document.getElementById('dispatch-sto-screen').classList.remove('hidden');
            document.getElementById('dispatch-tech-screen').classList.add('hidden');
        }

        function showTechScreen(sto, mitra) {
            document.getElementById('dispatch-sto-screen').classList.add('hidden');
            document.getElementById('dispatch-tech-screen').classList.remove('hidden');
            document.getElementById('selected-mitra-label').textContent = mitra;
            document.getElementById('selected-sto-label').textContent = sto;

            const filtered = getDispatchFilteredData().filter(i => i.sto === sto && i.mitra === mitra);
            const tbody = document.getElementById('dispatchTableBody');
            tbody.innerHTML = '';

            const grouped = {};
            filtered.forEach(item => {
                const tName = item.teknisi;
                if (!tName || tName === '-') return;
                if (!grouped[tName]) grouped[tName] = { total: 0, orders: [] };
                grouped[tName].total += 1;
                grouped[tName].orders.push(item);
            });

            const techList = Object.keys(grouped).sort();
            
            if (techList.length === 0) {
                tbody.innerHTML = '<tr><td colspan="4" class="text-center py-6 text-slate-400 font-medium">Tidak ada order aktif (Non-Completed) untuk teknisi di area STO ini.</td></tr>';
                return;
            }

            techList.forEach(tName => {
                // KUNCI PERBAIKAN: Melindungi nama yang ada tanda petiknya agar tidak merusak tombol
                const escapedTechName = tName.replace(/'/g, "\\'"); 
                
                const tr = document.createElement('tr');
                tr.className = "hover:bg-slate-50";
                
                tr.innerHTML = `
                    <td class="px-3 py-2 text-center">
                        <button onclick="sendDirectTelegramBot('${escapedTechName}', '${sto}', '${mitra}')" class="inline-flex items-center justify-center gap-1.5 bg-[#0088cc] hover:bg-[#0077b3] text-white px-4 py-2 rounded-lg font-bold transition shadow-sm w-full">
                            <i data-lucide="send" class="w-3.5 h-3.5"></i> Kirim
                        </button>
                    </td>
                    <td class="px-3 py-2 font-bold text-slate-800 text-[11px]">${tName}</td>
                    <td class="px-3 py-2 text-center font-black text-indigo-600 text-sm">${grouped[tName].total}</td>
                    <td class="px-3 py-2 text-center">
                        <!-- Panggil fungsi openDispatchDetail yang baru -->
                        <button onclick="openDispatchDetail('${escapedTechName}', '${sto}', '${mitra}')" class="text-xs font-semibold bg-slate-100 hover:bg-slate-200 text-slate-700 px-3 py-1.5 rounded-lg border border-slate-300 transition">
                            Lihat Detail
                        </button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
            lucide.createIcons();
        }

        // FUNGSI KHUSUS UNTUK TOMBOL DETAIL DI MENU DISPATCH
        function openDispatchDetail(techName, sto, mitra) {
            // Isi Judul Modal
            document.getElementById('modalTechName').innerHTML = `<i data-lucide="user" class="w-5 h-5 text-blue-600"></i> Tugas Aktif (Dispatch): ${techName} - Area ${sto}`;
            
            // Ambil data yang khusus belum selesai (non-completed)
            const filtered = getDispatchFilteredData();
            
            // Saring HANYA order di STO tersebut dan untuk teknisi tersebut
            const techOrders = filtered.filter(r => r.mitra === mitra && r.teknisi === techName && r.sto === sto);
            const tbody = document.getElementById('modalTechTableBody');
            tbody.innerHTML = '';

            if (techOrders.length === 0) {
                tbody.innerHTML = `<tr><td colspan="6" class="text-center py-4 text-slate-400">Tidak ada order aktif.</td></tr>`;
            } else {
                techOrders.forEach(item => {
                    const durasi = calculateDurationDays(item.tglCreate, item.tanggal);
                    const isOver3Days = durasi > 3;
                    const durasiHtml = isOver3Days 
                        ? `<span class="inline-flex items-center gap-1 text-rose-600 font-bold"><i data-lucide="alert-triangle" class="w-3.5 h-3.5"></i> ${durasi} Hr</span>`
                        : `<span class="font-bold text-slate-700">${durasi} Hr</span>`;

                    const tr = document.createElement('tr');
                    tr.className = "hover:bg-slate-50 text-xs";
                    tr.innerHTML = `
                        <td class="p-2 font-semibold">${item.sto}</td>
                        <td class="p-2 font-bold text-blue-600">${item.teknisi || '-'}</td>
                        <td class="p-2 font-mono">${item.wonum || '-'}</td>
                        <td class="p-2"><span class="px-1.5 py-0.5 rounded text-[10px] ${getStatusBadgeClass(item.status)}">${item.status}</span></td>
                        <td class="p-2 text-center">${durasiHtml}</td>
                        <td class="p-2 text-slate-600 wrap-cell">${item.keterangan || '-'}</td>
                    `;
                    tbody.appendChild(tr);
                });
            }
            
            lucide.createIcons();
            document.getElementById('techDetailModal').classList.remove('hidden');
        }

        // FUNGSI API TELEGRAM (Anti-Error "Message Too Long" - Sistem Cicil per Order)
        async function sendDirectTelegramBot(techName, sto, mitra) {
            
            const BOT_TOKEN = '8761040444:AAHuoy0HpbbhxantB5nJ9kHVSaelyw8LikQ'; 
            const TARGET_CHAT_ID = '5356190617'; 

            const filtered = getDispatchFilteredData().filter(i => i.sto === sto && i.mitra === mitra && i.teknisi === techName);
            
            if (filtered.length === 0) {
                alert('Tidak ada tugas aktif untuk dikirim.');
                return;
            }

            const url = `https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`;
            showToast('Mulai mengirim pesan ke Telegram...');

            try {

                // 2. KIRIM ORDER SATU PER SATU (Di-looping agar tidak terlalu panjang)
                for (let index = 0; index < filtered.length; index++) {
                    const order = filtered[index];
                    const typeOrder = order.typeOrder && order.typeOrder !== '-' ? order.typeOrder : 'CREATE';
                    const phone = order.cp && order.cp !== '-' ? order.cp.replace(/\D/g, '') : '';
                    const waLink = phone ? `https://wa.me/${phone}` : '#';
                    const mapsLink = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(order.alamat || '')}`;

                    const safeNama = (order.nama || '-').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
                    const safeAlamat = (order.alamat || '-').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

                    let orderText = `📦 <b>ORDER ${index + 1} DARI ${filtered.length}</b>\n`;
                    orderText += `--------------------------------\n`;
                    orderText += `<b>ORDER PSB SA ${order.sto}</b>\n`;
                    orderText += `#${typeOrder.replace(/\s+/g, '_')}_${order.sto}\n\n`;
                    
                    orderText += `<b>TGL ORDER :</b> <code>${order.tanggal}</code>\n`;
                    orderText += `<b>HOMEPASS ID :</b> <code>${order.homepassId}</code>\n`;
                    orderText += `<b>ODP :</b> <code>${order.odp}</code>\n\n`;
                    
                    orderText += `<b>TYPE | STO | PAKET :</b>\n`;
                    orderText += `<code>${typeOrder} | ${order.sto} | ${order.paket}</code>\n\n`;
                    
                    orderText += `<b>WONUM | NO SERVICE :</b>\n`;
                    orderText += `<code>${order.wonum} | ${order.noService}</code>\n\n`;
                    
                    orderText += `<b>NO ORDER :</b>\n`;
                    orderText += `<code>${order.noOrder}</code>\n\n`;
                    
                    orderText += `<b>SUMMARY :</b>\n`;
                    orderText += `${safeNama} // ${safeAlamat} // ${order.cp}\n\n`;
                    
                    orderText += `<b>CREATE :</b> <code>${order.tglCreate}</code>\n`;
                    orderText += `<b>MANJA :</b> <code>${order.tglManja}</code>\n\n`;
                    
                    orderText += `💬 <a href="${waLink}">WA</a> | 📍 <a href="${mapsLink}">Maps</a>\n\n`;
                    
                    orderText += `<b>TEKNISI :</b>\n`;
                    orderText += `<b>${order.teknisi}</b>`;

                    // Kirim pesan per order
                    await fetch(url, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            chat_id: TARGET_CHAT_ID,
                            text: orderText,
                            parse_mode: 'HTML',
                            disable_web_page_preview: true
                        })
                    });

                    // Jeda 300ms agar urutan masuk di Telegram tidak terbalik dan tidak dianggap spam
                    await new Promise(resolve => setTimeout(resolve, 300));
                }

                showToast(`✅ Berhasil! ${filtered.length} order untuk ${techName} terkirim.`);

            } catch (error) {
                alert('Gagal mengirim sebagian/seluruh pesan! Pastikan ada koneksi internet.');
                console.error(error);
            }
        }

        function renderTeknisiTable() {
            const filtered = getFilteredData();
            const container = document.getElementById('teknisiTablesContainer');
            const search = (document.getElementById('teknisiSearch').value || '').toLowerCase();
            container.innerHTML = '';

            if (filtered.length === 0) {
                container.innerHTML = `<div class="py-6 text-center text-slate-400">Tidak ada data teknisi.</div>`;
                return;
            }

            const grouped = {};
            filtered.forEach(item => {
                const m = item.mitra || 'UNKNOWN';
                const t = (item.teknisi && item.teknisi !== '-') ? item.teknisi : 'Tanpa Nama';
                if (search && !t.toLowerCase().includes(search)) return;

                if (!grouped[m]) grouped[m] = {};
                if (!grouped[m][t]) {
                    grouped[m][t] = { stos: new Set(), total: 0, progressNumerator: 0 };
                    STATUS_COLUMNS.forEach(col => grouped[m][t][col] = 0);
                }
                grouped[m][t].total += 1;
                grouped[m][t].stos.add(item.sto);
                grouped[m][t][item.status] = (grouped[m][t][item.status] || 0) + 1;
                if (PROGRESS_STATUS_KEYS.includes(item.status)) grouped[m][t].progressNumerator += 1;
            });

            Object.keys(grouped).sort().forEach(mitraKey => {
                const theme = MITRA_THEMES[mitraKey] || DEFAULT_THEME;
                const techs = grouped[mitraKey];
                const techList = Object.keys(techs).sort();
                const section = document.createElement('div');
                section.className = 'border border-slate-200 rounded-xl overflow-hidden shadow-sm';

                let rowsHtml = '';
                const subtotalCounts = {};
                STATUS_COLUMNS.forEach(col => subtotalCounts[col] = 0);

                // Tambahkan alat penerjemah format 0 vs Strip
                const formatVal = (num) => (num === 0 ? (showZeroAsDash ? '-' : '0') : num);

                techList.forEach((tName, idx) => {
                    const data = techs[tName];
                    const estComp = (data['COMPLETED'] || 0) + (data['ACTIVATION COMPLETED'] || 0) + (data['IKR OKE'] || 0);

                    // Terapkan alat penerjemah ke semua angka
                    let statusCellsHtml = `<td class="bg-emerald-50 font-bold">${formatVal(estComp)}</td><td>${formatVal(data['COMPLETED'] || 0)}</td>`;
                    STATUS_COLUMNS.slice(2).forEach(col => {
                        const val = data[col] || 0;
                        subtotalCounts[col] += val;
                        statusCellsHtml += `<td>${formatVal(val)}</td>`;
                    });

                    rowsHtml += `
                        <tr class="hover:bg-slate-50 text-[11px]">
                            <td class="text-center text-slate-400">${idx + 1}</td>
                            <td class="font-bold text-blue-600 cursor-pointer truncate" onclick="openTechModal('${tName}', '${mitraKey}')">${tName}</td>
                            <td class="truncate">${Array.from(data.stos).join(',')}</td>
                            ${statusCellsHtml}
                            <td class="text-center font-bold">${data.total}</td>
                            <td class="text-center ${getProgressColorClass(data.total > 0 ? (data.progressNumerator/data.total)*100 : 0)}">${(data.total > 0 ? (data.progressNumerator/data.total)*100 : 0).toFixed(0)}%</td>
                        </tr>
                    `;
                });

                section.innerHTML = `
                    <div class="${theme.subtotalBgClass} px-3 py-2 text-xs flex justify-between">
                        <span>MITRA ${mitraKey} (${techList.length} Teknisi)</span>
                    </div>
                    <div class="w-full overflow-hidden">
                        <table class="w-full text-left report-table table-fixed">
                            <thead>
                                <tr class="bg-black text-white text-[9px] uppercase">
                                    <th class="w-[5%]">No</th>
                                    <th class="w-[18%]">Teknisi</th>
                                    <th class="w-[10%]">STO</th>
                                    <th class="w-[5%] bg-emerald-950">EST</th>
                                    <th class="w-[5%]">COM</th>
                                    <th class="w-[5%]">ACT</th>
                                    <th class="w-[5%]">IOK</th>
                                    <th class="w-[5%]">INOK</th>
                                    <th class="w-[5%]">FALL</th>
                                    <th class="w-[5%]">KSYS</th>
                                    <th class="w-[5%]">SURV</th>
                                    <th class="w-[5%]">KPEL</th>
                                    <th class="w-[5%]">KTEK</th>
                                    <th class="w-[5%]">OGP</th>
                                    <th class="w-[5%]">SALD</th>
                                    <th class="w-[6%]">TOT</th>
                                    <th class="w-[7%]">PROG</th>
                                </tr>
                            </thead>
                            <tbody>${rowsHtml}</tbody>
                        </table>
                    </div>
                `;
                container.appendChild(section);
            });
            lucide.createIcons();
        }

        function openTechModal(techName, mitraKey) {
            document.getElementById('modalTechName').textContent = `Teknisi: ${techName} (${mitraKey})`;
            const filtered = getFilteredData();
            const techOrders = filtered.filter(r => r.mitra === mitraKey && r.teknisi === techName);
            const tbody = document.getElementById('modalTechTableBody');
            tbody.innerHTML = '';

            if (techOrders.length === 0) {
                tbody.innerHTML = `<tr><td colspan="6" class="text-center py-4 text-slate-400">Tidak ada order.</td></tr>`;
            } else {
                techOrders.forEach(item => {
                    const durasi = calculateDurationDays(item.tglCreate, item.tanggal);
                    const isOver3Days = durasi > 3;
                    const durasiHtml = isOver3Days 
                        ? `<span class="inline-flex items-center gap-1 text-rose-600 font-bold" title="Sudah lewat TTI (> 3 Hari)"><i data-lucide="alert-triangle" class="w-3.5 h-3.5"></i> ${durasi} Hr</span>`
                        : `<span class="font-bold text-slate-700">${durasi} Hr</span>`;

                    const tr = document.createElement('tr');
                    tr.className = "hover:bg-slate-50 text-xs";
                    tr.innerHTML = `
                        <td class="p-2 font-semibold">${item.sto}</td>
                        <td class="p-2 font-bold text-blue-600">${item.teknisi || '-'}</td>
                        <td class="p-2 font-mono">${item.wonum || '-'}</td>
                        <td class="p-2"><span class="px-1.5 py-0.5 rounded text-[10px] ${getStatusBadgeClass(item.status)}">${item.status}</span></td>
                        <td class="p-2 text-center">${durasiHtml}</td>
                        <td class="p-2 text-slate-600 wrap-cell">${item.keterangan || '-'}</td>
                    `;
                    tbody.appendChild(tr);
                });
            }
            lucide.createIcons();
            document.getElementById('techDetailModal').classList.remove('hidden');
        }
        function closeTechModal() { document.getElementById('techDetailModal').classList.add('hidden'); }

        // Fungsi Modal Detail Khusus untuk halaman Dispatch (Filter berdasarkan STO, bukan Mitra)
        function openTechModalDispatch(techName, stoName) {
            document.getElementById('modalTechName').textContent = `Teknisi: ${techName} (Area: ${stoName})`;
            const filtered = getFilteredData();
            const techOrders = filtered.filter(r => r.sto === stoName && r.teknisi === techName);
            const tbody = document.getElementById('modalTechTableBody');
            tbody.innerHTML = '';

            if (techOrders.length === 0) {
                tbody.innerHTML = `<tr><td colspan="6" class="text-center py-4 text-slate-400">Tidak ada order.</td></tr>`;
            } else {
                techOrders.forEach(item => {
                    const durasi = calculateDurationDays(item.tglCreate, item.tanggal);
                    const isOver3Days = durasi > 3;
                    const durasiHtml = isOver3Days 
                        ? `<span class="inline-flex items-center gap-1 text-rose-600 font-bold" title="Sudah lewat TTI (> 3 Hari)"><i data-lucide="alert-triangle" class="w-3.5 h-3.5"></i> ${durasi} Hr</span>`
                        : `<span class="font-bold text-slate-700">${durasi} Hr</span>`;

                    const tr = document.createElement('tr');
                    tr.className = "hover:bg-slate-50 text-xs";
                    tr.innerHTML = `
                        <td class="p-2 font-semibold">${item.sto}</td>
                        <td class="p-2 font-bold text-blue-600">${item.teknisi || '-'}</td>
                        <td class="p-2 font-mono">${item.wonum || '-'}</td>
                        <td class="p-2"><span class="px-1.5 py-0.5 rounded text-[10px] ${getStatusBadgeClass(item.status)}">${item.status}</span></td>
                        <td class="p-2 text-center">${durasiHtml}</td>
                        <td class="p-2 text-slate-600 wrap-cell">${item.keterangan || '-'}</td>
                    `;
                    tbody.appendChild(tr);
                });
            }
            lucide.createIcons();
            document.getElementById('techDetailModal').classList.remove('hidden');
        }

        function renderActTable() {
            renderGenericTable('ACTIVATION COMPLETED', 'actSearch', 'actTableBody', 'badgeActTotal');
        }
        function renderIkrOkeTable() {
            renderGenericTable('IKR OKE', 'ikrOkeSearch', 'ikrOkeTableBody', 'badgeIkrOkeTotal');
        }
        function renderIkrNokTable() {
            renderGenericTable('IKR NOK', 'ikrNokSearch', 'ikrNokTableBody', 'badgeIkrNokTotal');
        }
        function renderKendalaTeknikTable() {
            renderGenericTable('KENDALA TEKNIK', 'kendalaTeknikSearch', 'kendalaTeknikTableBody', 'badgeKendalaTeknikTotal');
        }
        function renderKendalaPelangganTable() {
            renderGenericTable('KENDALA PELANGGAN', 'kendalaPelangganSearch', 'kendalaPelangganTableBody', 'badgeKendalaPelangganTotal');
        }

        function renderGenericTable(statusName, searchId, tbodyId, badgeId) {
            const filtered = getFilteredData();
            const search = (document.getElementById(searchId).value || '').toLowerCase();
            const tbody = document.getElementById(tbodyId);
            const list = filtered.filter(item => {
                if (item.status !== statusName) return false;
                if (!search) return true;
                return (item.teknisi.toLowerCase().includes(search) || item.noOrder.toLowerCase().includes(search) || item.mitra.toLowerCase().includes(search) || item.keterangan.toLowerCase().includes(search));
            });
            document.getElementById(badgeId).textContent = list.length;
            tbody.innerHTML = '';
            if (list.length === 0) {
                tbody.innerHTML = `<tr><td colspan="8" class="text-center py-6 text-slate-400">Tidak ada data ${statusName}.</td></tr>`;
                return;
            }
            list.forEach((item, index) => {
                const durasi = calculateDurationDays(item.tglCreate, item.tanggal);
                const tr = document.createElement('tr');
                tr.className = "hover:bg-slate-50 text-xs";
                if (statusName === 'IKR OKE' || statusName === 'ACTIVATION COMPLETED') {
                    tr.innerHTML = `
                        <td class="px-3 py-2 text-center">${index + 1}</td>
                        <td class="px-3 py-2 font-bold uppercase">${item.mitra}</td>
                        <td class="px-3 py-2 font-semibold">${item.sto}</td>
                        <td class="px-3 py-2 font-semibold">${item.teknisi}</td>
                        <td class="px-3 py-2 font-mono text-blue-600">${item.noOrder}</td>
                        <td class="px-3 py-2 text-center font-bold">${durasi} Hr</td>
                        <td class="px-3 py-2 text-slate-700 wrap-cell">${item.keterangan || '-'}</td>
                    `;
                } else if (statusName === 'IKR NOK') {
                    tr.innerHTML = `
                        <td class="px-3 py-2 text-center">${index + 1}</td>
                        <td class="px-3 py-2 font-bold uppercase">${item.mitra}</td>
                        <td class="px-3 py-2 font-semibold">${item.sto}</td>
                        <td class="px-3 py-2 font-mono text-blue-600">${item.noOrder}</td>
                        <td class="px-3 py-2">${item.nama}</td>
                        <td class="px-3 py-2">${item.teknisi}</td>
                        <td class="px-3 py-2 text-center font-bold">${durasi} Hr</td>
                        <td class="px-3 py-2 text-slate-700 wrap-cell">${item.keterangan || '-'}</td>
                    `;
                } else {
                    tr.innerHTML = `
                        <td class="px-3 py-2 text-center">${index + 1}</td>
                        <td class="px-3 py-2 font-bold uppercase">${item.mitra}</td>
                        <td class="px-3 py-2 font-semibold">${item.sto}</td>
                        <td class="px-3 py-2 font-mono text-blue-600">${item.noOrder}</td>
                        <td class="px-3 py-2">${item.teknisi}</td>
                        <td class="px-3 py-2 text-center font-bold">${durasi} Hr</td>
                        <td class="px-3 py-2 font-semibold text-amber-700">${item.subKendala || '-'}</td>
                        <td class="px-3 py-2 text-slate-700 wrap-cell">${item.keterangan || '-'}</td>
                    `;
                }
                tbody.appendChild(tr);
            });
        }

        function renderAnalytics() {
            const filtered = getFilteredData();
            
            // ==========================================
            // 1. MENGHIDUPKAN 4 KOTAK KPI ATAS
            // ==========================================
            let totalProgressNumSum = 0;
            let totalEstCompAnalytics = 0;
            let totalNokAnalytics = 0;

            filtered.forEach(item => {
                if (PROGRESS_STATUS_KEYS.includes(item.status)) totalProgressNumSum += 1;
                if (item.status === 'COMPLETED' || item.status === 'ACTIVATION COMPLETED' || item.status === 'IKR OKE') totalEstCompAnalytics += 1;
                if (item.status === 'IKR NOK' || item.status === 'KENDALA TEKNIK' || item.status === 'KENDALA PELANGGAN' || item.status === 'KENDALA SISTEM') totalNokAnalytics += 1;
            });

            const overallProgressPct = filtered.length > 0 ? ((totalProgressNumSum / filtered.length) * 100).toFixed(1) : 0;
            
            const elTotal = document.getElementById('statAnalyticsTotal');
            const elProg = document.getElementById('statAnalyticsProgress');
            const elEst = document.getElementById('statAnalyticsEstComp');
            const elNok = document.getElementById('statAnalyticsNok');

            if (elTotal) elTotal.textContent = filtered.length;
            if (elProg) elProg.textContent = overallProgressPct + '%';
            if (elEst) elEst.textContent = totalEstCompAnalytics;
            if (elNok) elNok.textContent = totalNokAnalytics;

            // Bersihkan instance chart lama
            if (window.chartBarMitra) window.chartBarMitra.destroy();
            if (window.chartTrenHarian) window.chartTrenHarian.destroy();

            // ==========================================
            // 2. GRAFIK BAR CHART (FIX NABRAK LEGEND)
            // ==========================================
            const ctxBar = document.getElementById('barMitraChart');
            if (ctxBar) {
                const ctx2dBar = ctxBar.getContext('2d');
                const mitraList = [...new Set(filtered.map(item => item.mitra).filter(m => m && m !== 'UNKNOWN'))].sort();
                
                const dataCompleted = mitraList.map(m => filtered.filter(i => i.mitra === m && (i.status === 'COMPLETED' || i.status === 'ACTIVATION COMPLETED' || i.status === 'IKR OKE')).length);
                const dataKP = mitraList.map(m => filtered.filter(i => i.mitra === m && i.status === 'KENDALA PELANGGAN').length);
                const dataKT = mitraList.map(m => filtered.filter(i => i.mitra === m && i.status === 'KENDALA TEKNIK').length);

                const getColors = (type) => mitraList.map(m => {
                    if (m === 'BISTEL') return type === 'COMP' ? '#0051ff' : (type === 'KP' ? '#60a5fa' : '#c3ddfa'); // Biru Tua -> Biru Sedang -> Biru Muda
                    if (m === 'TA') return type === 'COMP' ? '#ff0000' : (type === 'KP' ? '#ff6b6b' : '#fcc3c3'); // Merah Tua -> Merah Sedang -> Merah Muda
                    if (m === 'BST') return type === 'COMP' ? '#ffcc00' : (type === 'KP' ? '#ffe16a' : '#f8efc3'); // Oranye -> Kuning -> Kuning Muda
                    return type === 'COMP' ? '#475569' : (type === 'KP' ? '#94a3b8' : '#cbd5e1'); // Default
                });

                window.chartBarMitra = new Chart(ctx2dBar, {
                    type: 'bar',
                    data: {
                        labels: mitraList,
                        datasets: [
                            { label: 'Completed', data: dataCompleted, backgroundColor: getColors('COMP'), borderRadius: 6, borderWidth: 0, categoryPercentage: 0.6, barPercentage: 0.8 },
                            { label: 'Kendala Pelanggan', data: dataKP, backgroundColor: getColors('KP'), borderRadius: 6, borderWidth: 0, categoryPercentage: 0.6, barPercentage: 0.8 },
                            { label: 'Kendala Teknik', data: dataKT, backgroundColor: getColors('KT'), borderRadius: 6, borderWidth: 0, categoryPercentage: 0.6, barPercentage: 0.8 }
                        ]
                    },
                    options: {
                        responsive: true, maintainAspectRatio: false,
                        layout: {
                            padding: { top: 25 } // SOLUSI: Menambah jarak plafon biar angka gak numpuk legend
                        },
                        plugins: {
                            legend: { position: 'top', align: 'end', labels: { usePointStyle: true, pointStyle: 'circle', boxWidth: 8, font: {family: 'Inter', size: 11} } },
                            datalabels: {
                                display: function(context) { return context.dataset.data[context.dataIndex] > 0; },
                                anchor: 'end', align: 'end', offset: 2,
                                backgroundColor: '#1e293b', borderRadius: 4, color: '#ffffff', 
                                font: { weight: 'bold', size: 11, family: 'Inter' },
                                padding: { top: 3, bottom: 3, left: 6, right: 6 }
                            }
                        },
                        scales: {
                            y: { 
                                beginAtZero: true, 
                                grace: '15%', // SOLUSI: Tiang tertinggi gak akan menyentuh batas atas
                                grid: { borderDash: [4, 4], color: '#f1f5f9' }, 
                                border: { display: false } 
                            },
                            x: { grid: { display: false }, border: { display: false }, ticks: { font: { weight: 'bold', family: 'Inter', size: 12 } } }
                        }
                    },
                    plugins: [ChartDataLabels]
                });
            }

            // ==========================================
            // 3. GRAFIK LINE CHART (TREN HARIAN - TOTAL KESELURUHAN)
            // ==========================================
            const ctxLine = document.getElementById('trenHarianChart');
            if (ctxLine) {
                const ctx2d = ctxLine.getContext('2d');
                const selectedMitra = document.getElementById('filterMitra').value;
                const trendData = selectedMitra === 'ALL' ? rawData : rawData.filter(item => item.mitra === selectedMitra);

                const rawDates = [...new Set(trendData.map(item => item.tanggal).filter(Boolean))].sort();
                const formattedDates = rawDates.map(d => {
                    const dt = new Date(d);
                    return dt.toLocaleDateString('id-ID', { day: '2-digit', month: 'short' });
                });

                // Hitung TOTAL completed harian (Gabungan semua mitra)
                const dailyCompleted = rawDates.map(date => {
                    return trendData.filter(i => i.tanggal === date && (i.status === 'COMPLETED' || i.status === 'ACTIVATION COMPLETED' || i.status === 'IKR OKE')).length;
                });

                // Efek Biru Transparan (Lebih pekat di atas, memudar di bawah)
                let gradientBiru = ctx2d.createLinearGradient(0, 0, 0, 300);
                gradientBiru.addColorStop(0, 'rgba(0, 102, 255, 0.8)'); 
                gradientBiru.addColorStop(1, 'rgba(0, 94, 245, 0)'); 

                window.chartTrenHarian = new Chart(ctx2d, {
                    type: 'line',
                    data: {
                        labels: formattedDates,
                        datasets: [
                            {
                                label: 'Total Completed Harian',
                                data: dailyCompleted,
                                borderColor: 'transparent', // Garis pinggir dihilangkan
                                backgroundColor: gradientBiru,
                                borderWidth: 0, // <--- KUNCI: Garis tebal dihapus jadi 0
                                fill: true,
                                pointRadius: 0, // Titik sembunyi
                                pointHoverRadius: 6, // Muncul saat di-hover
                                pointBackgroundColor: '#ffffff',
                                pointBorderColor: '#2543eb',
                                pointBorderWidth: 2,
                                tension: 0 // Kaku / Tajam bersudut
                            }
                        ]
                    },
                    options: {
                        responsive: true, maintainAspectRatio: false,
                        interaction: { mode: 'index', intersect: false }, 
                        plugins: {
                            legend: { position: 'top', align: 'end', labels: { usePointStyle: true, boxWidth: 10, font: {family: 'Inter', size: 11} } },
                            datalabels: { display: false } 
                        },
                        scales: {
                            y: { beginAtZero: true, grid: { borderDash: [5, 5], color: '#ffffff' }, border: { display: false }, ticks: { font: { size: 10, family: 'Inter' } } },
                            x: { grid: { display: false }, border: { display: false }, ticks: { font: { size: 10, family: 'Inter' }, maxRotation: 0, minRotation: 0, maxTicksLimit: 12 } }
                        }
                    }
                });
            }
        }

        function updateSummaryStats(total, completed, estCompleted, ikrOke, ikrNok, ogp, progress) {
            document.getElementById('statTotalOrders').textContent = total;
            document.getElementById('statCompleted').textContent = completed;
            document.getElementById('statEstCompleted').textContent = estCompleted;
            document.getElementById('statIkrOke').textContent = ikrOke;
            document.getElementById('statIkrNok').textContent = ikrNok;
            document.getElementById('statOGP').textContent = ogp;
            document.getElementById('statProgress').textContent = progress + '%';
        }

        function renderRawTable() {
            const tbody = document.getElementById('rawTableBody');
            const search = (document.getElementById('rawSearch').value || '').toLowerCase();
            const filtered = getFilteredData(); // Menggunakan filter tanggal, mitra, & status

            const list = filtered.filter(r => 
                r.mitra.toLowerCase().includes(search) || 
                r.noOrder.toLowerCase().includes(search) ||
                r.nama.toLowerCase().includes(search) ||
                r.sto.toLowerCase().includes(search) ||
                r.status.toLowerCase().includes(search)
            );

            tbody.innerHTML = '';
            if (list.length === 0) {
                tbody.innerHTML = '<tr><td colspan="7" class="text-center py-6 text-slate-400">Tidak ada data mentah sesuai filter.</td></tr>';
                return;
            }

            list.forEach((item, index) => {
                const tr = document.createElement('tr');
                tr.className = "hover:bg-slate-50 text-xs";
                tr.innerHTML = `
                    <td class="px-2 py-2 text-center">${index + 1}</td>
                    <td class="px-2 py-2 truncate">${item.tanggal}</td>
                    <td class="px-2 py-2 font-bold truncate">${item.mitra}</td>
                    <td class="px-2 py-2 truncate">${item.sto}</td>
                    <td class="px-2 py-2 font-mono text-blue-600 truncate">${item.noOrder}</td>
                    <td class="px-2 py-2 truncate">${item.nama}</td>
                    <td class="px-2 py-2 truncate"><span class="px-1.5 py-0.5 rounded text-[10px] ${getStatusBadgeClass(item.status)}">${item.status}</span></td>
                `;
                tbody.appendChild(tr);
            });
        }

        function switchTab(tabName) {
            ['table', 'teknisi','dispatch', 'act', 'ikroke', 'ikrnok', 'kendalateknik', 'kendalapelanggan', 'analytics', 'raw'].forEach(t => {
                const btn = document.getElementById('tab-' + t);
                const view = document.getElementById('view-' + t);
                if (btn) {
                    if (t === tabName) {
                        btn.className = t === 'ai' 
                            ? "w-full flex items-center space-x-3 px-3 py-2.5 rounded-xl text-xs font-medium text-indigo-300 bg-indigo-900/80 transition border border-indigo-500/30"
                            : "w-full flex items-center space-x-3 px-3 py-2.5 rounded-xl text-xs font-semibold bg-blue-600 text-white shadow-sm transition";
                    } else {
                        btn.className = t === 'ai'
                            ? "w-full flex items-center space-x-3 px-3 py-2.5 rounded-xl text-xs font-medium text-indigo-300 bg-indigo-950/50 hover:bg-indigo-900/60 transition border border-indigo-500/20"
                            : "w-full flex items-center space-x-3 px-3 py-2.5 rounded-xl text-xs font-medium text-slate-300 hover:bg-slate-800 hover:text-white transition";
                    }
                }
                if (view) {
                    if (t === tabName) view.classList.remove('hidden');
                    else view.classList.add('hidden');
                }
            });
            if (window.innerWidth < 1024) {
                const sidebar = document.getElementById('appSidebar');
                const backdrop = document.getElementById('sidebarBackdrop');
                sidebar.classList.add('-translate-x-full');
                backdrop.classList.add('hidden');
            }
            if (tabName === 'analytics') renderAnalytics();
        }

        function openGoogleSheetModal() { document.getElementById('googleSheetModal').classList.remove('hidden'); }
        function closeGoogleSheetModal() { document.getElementById('googleSheetModal').classList.add('hidden'); }

        async function saveAndFetchGSheet() {
            const url = document.getElementById('gsheetUrl').value.trim();
            const sheetName = document.getElementById('gsheetName').value.trim() || DEFAULT_SHEET_NAME;
            localStorage.setItem('saved_gsheet_url', url);
            localStorage.setItem('saved_gsheet_name', sheetName);
            closeGoogleSheetModal();
            await triggerRefreshData();
        }

        window.addEventListener('DOMContentLoaded', () => {
            lucide.createIcons();
            const todayStr = new Date().toISOString().split('T')[0];
            
            // Setel default ke tanggal HARI INI (pakai filter tunggal)
            document.getElementById('filterDate').value = todayStr;
            document.getElementById('filterStartDate').value = '';
            document.getElementById('filterEndDate').value = '';
            
            fetchDatabaseData(false);
        });

        function handleFileUpload(event) {
            const file = event.target.files[0];
            if (!file) return;

            const formData = new FormData();
            formData.append('file', file);
            
            showToast("Sedang mengupload dan menyinkronkan Database...");

            fetch('upload.php', { method: 'POST', body: formData })
            .then(response => response.text())
            .then(data => {
                if(data.includes("Berhasil")) {
                    showToast("Upload Selesai! Merender data...");
                    fetchDatabaseData(false); // Refresh tabel otomatis
                } else {
                    showToast("Gagal: " + data);
                }
            })
            .catch(err => showToast("Error koneksi upload!"));
            
            event.target.value = ''; // Reset input form
        }

        async function triggerRefreshData() {
            const icon = document.getElementById('refreshIcon');
            icon.classList.add('animate-spin');
            await fetchDatabaseData(false);
            setTimeout(() => icon.classList.remove('animate-spin'), 600);
            showToast("Data dari Database MySQL berhasil direfresh!");
        }

        // Pengganti Google Sheets: Tarik JSON dari api.php lokalmu
        async function fetchDatabaseData(isBackground = false) {
            try {
                const response = await fetch('api.php');
                const json = await response.json();
                
                if (Array.isArray(json)) {
                    parseAndLoadRows(json); // Kirim JSON dari PHP ke logika tabel Margo
                    if (!isBackground) showToast("Berhasil terhubung ke Database MySQL!");
                }
            } catch (err) {
                console.warn("Fetch DB error:", err);
                if (!isBackground) showToast("Gagal mengambil data dari Database lokal.");
            }
        }

        function exportToExcel() {
            XLSX.writeFile(XLSX.utils.table_to_book(document.getElementById('mainPivotTable'), { sheet: "Pivot" }), "Laporan_WOC_Samarinda.xlsx");
            showToast("Export Excel berhasil!");
        }

        function copyToClipboardWA() {
            const filtered = getFilteredData();
            let total = filtered.length;
            let completed = filtered.filter(r => r.status === 'COMPLETED').length;
            let estCompleted = filtered.filter(r => r.status === 'COMPLETED' || r.status === 'ACTIVATION COMPLETED' || r.status === 'IKR OKE').length;
            let ikrOke = filtered.filter(r => r.status === 'IKR OKE').length;
            let ikrNok = filtered.filter(r => r.status === 'IKR NOK').length;

            let text = `📊 *LAPORAN HARIAN STATUS ORDER WOC SAMARINDA*\n\n`;
            text += `• Total Order: *${total}*\n`;
            text += `• Est. Completed: *${estCompleted}*\n`;
            text += `• Completed: *${completed}*\n`;
            text += `• IKR OKE: *${ikrOke}*\n`;
            text += `• IKR NOK: *${ikrNok}*\n`;

            const textArea = document.createElement("textarea");
            textArea.value = text;
            document.body.appendChild(textArea);
            textArea.select();
            document.execCommand('copy');
            document.body.removeChild(textArea);
            showToast("Ringkasan disalin ke Clipboard (Format WA)!");
        }

        function showToast(message) {
            const toast = document.getElementById('toast');
            document.getElementById('toastMessage').textContent = message;
            toast.classList.remove('translate-y-20', 'opacity-0');
            setTimeout(() => toast.classList.add('translate-y-20', 'opacity-0'), 3500);
        }

        // =====================================
        // FUNGSI MODAL UNTUK 4 KOTAK KPI
        // =====================================
        function openKpiModal(type) {
            const filtered = getFilteredData(); // Ambil data sesuai filter kalender/mitra saat ini
            let matchedData = [];
            let title = '';

            // Saring data berdasarkan kotak mana yang diklik
            if (type === 'total') {
                matchedData = filtered;
                title = 'Detail: Total WO Aktif';
            } else if (type === 'progress') {
                matchedData = filtered.filter(item => PROGRESS_STATUS_KEYS.includes(item.status));
                title = 'Detail: Data Progress Lapangan';
            } else if (type === 'est_completed') {
                matchedData = filtered.filter(item => ['COMPLETED', 'ACTIVATION COMPLETED', 'IKR OKE'].includes(item.status));
                title = 'Detail: Total Est. Completed';
            } else if (type === 'nok') {
                matchedData = filtered.filter(item => ['IKR NOK', 'KENDALA TEKNIK', 'KENDALA PELANGGAN', 'KENDALA SISTEM'].includes(item.status));
                title = 'Detail: Total Kendala / NOK';
            }

            // Ubah Judul Modal
            document.getElementById('modalKpiTitle').innerHTML = `<i data-lucide="layers" class="w-5 h-5 text-indigo-500"></i> ${title} <span class="bg-indigo-100 text-indigo-700 px-2 py-0.5 rounded-full text-xs ml-2">${matchedData.length} Order</span>`;
            
            const tbody = document.getElementById('modalKpiTableBody');
            tbody.innerHTML = '';

            if (matchedData.length === 0) {
                tbody.innerHTML = `<tr><td colspan="6" class="text-center py-6 text-slate-400">Tidak ada data untuk kategori ini.</td></tr>`;
            } else {
                matchedData.forEach(item => {
                    const durasi = calculateDurationDays(item.tglCreate, item.tanggal);
                    const isOver3Days = durasi > 3;
                    const durasiHtml = isOver3Days 
                        ? `<span class="inline-flex items-center gap-1 text-rose-600 font-bold"><i data-lucide="alert-triangle" class="w-3.5 h-3.5"></i> ${durasi} Hr</span>`
                        : `<span class="font-bold text-slate-700">${durasi} Hr</span>`;

                    const tr = document.createElement('tr');
                    tr.className = "hover:bg-slate-50 text-xs";
                    tr.innerHTML = `
                        <td class="p-2 font-bold text-slate-800 uppercase">${item.sto}</td>
                        <td class="p-2 font-bold text-blue-600">${item.teknisi || '-'}</td>
                        <td class="p-2 font-mono">${item.wonum || '-'}</td>
                        <td class="p-2"><span class="px-1.5 py-0.5 rounded text-[10px] ${getStatusBadgeClass(item.status)}">${item.status}</span></td>
                        <td class="p-2 text-center">${durasiHtml}</td>
                        <td class="p-2 text-slate-600 wrap-cell">${item.keterangan || '-'}</td>
                    `;
                    tbody.appendChild(tr);
                });
            }
            
            lucide.createIcons();
            document.getElementById('kpiDetailModal').classList.remove('hidden');
        }

        function closeKpiModal() {
            document.getElementById('kpiDetailModal').classList.add('hidden');
        }
    </script>
</body>
</html>