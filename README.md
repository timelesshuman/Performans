<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Reazy | VCT Performans & Gözlem Aracı</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- SheetJS (Excel dosyalarını okumak için eklendi) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <!-- Google Fonts -->
    <link href="https://fonts.googleapis.com/css2?family=Oswald:wght@500;700&family=Inter:wght@400;600&display=swap" rel="stylesheet">
    
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        valDark: '#0F1923',
                        valCard: '#1F2E3A',
                        valRed: '#FF4655',
                        valLight: '#ECE8E1',
                        valGray: '#8B978F',
                        valGreen: '#00B5B8'
                    },
                    fontFamily: {
                        heading: ['Oswald', 'sans-serif'],
                        body: ['Inter', 'sans-serif']
                    }
                }
            }
        }
    </script>
    <style>
        body { background-color: #0F1923; color: #ECE8E1; overflow-x: hidden; }
        .val-border { border-top: 4px solid #FF4655; }
        .glass-card { background: rgba(31, 46, 58, 0.9); backdrop-filter: blur(10px); }
        .tab-btn { transition: all 0.3s; border-bottom: 2px solid transparent; }
        .tab-btn.active { color: #FF4655; border-bottom: 2px solid #FF4655; font-weight: bold; }
        .tab-btn:hover:not(.active) { color: #ECE8E1; border-bottom: 2px solid #8B978F; }
        select { background-image: none; outline: none; }
        ::-webkit-scrollbar { width: 8px; }
        ::-webkit-scrollbar-track { background: #0F1923; }
        ::-webkit-scrollbar-thumb { background: #1F2E3A; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #FF4655; }
    </style>
</head>
<body class="font-body min-h-screen flex flex-col">

    <!-- Navbar & Header -->
    <header class="bg-valCard border-b border-gray-700 sticky top-0 z-50">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div class="flex flex-col md:flex-row justify-between items-center py-4">
                <div class="flex items-center gap-3">
                    <div class="w-10 h-10 bg-valRed flex items-center justify-center rounded transform rotate-45">
                        <span class="text-white font-heading font-bold text-xl transform -rotate-45">R</span>
                    </div>
                    <div>
                        <h1 class="text-2xl font-heading text-valLight uppercase tracking-wide">REAZY ANALYTICS</h1>
                        <p class="text-valGray text-xs uppercase tracking-wider">Scouting & Performance Tool v2.1</p>
                    </div>
                </div>
                <div class="mt-4 md:mt-0 flex items-center gap-3">
                    <label class="text-xs text-valGray uppercase">Veri Yükle (CSV / XLSX):</label>
                    <input type="file" id="csvFileInput" accept=".csv, .xlsx, .xls" class="text-sm bg-valDark text-valLight rounded px-2 py-1 border border-gray-600 focus:border-valRed outline-none cursor-pointer">
                </div>
            </div>
            
            <!-- Dinamik Menü -->
            <nav class="flex space-x-6 overflow-x-auto mt-2" id="navMenu"></nav>
        </div>
    </header>

    <!-- Main Content Area -->
    <main class="flex-grow max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 w-full">
        
        <!-- GENEL SAYFASI (Default) -->
        <div id="view-general" class="space-y-8 block animate-fade-in">
            <!-- Özet Kartları -->
            <div class="grid grid-cols-1 md:grid-cols-4 gap-4" id="genSummaryCards"></div>
            
            <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                <!-- Trend Grafiği (Line) -->
                <div class="lg:col-span-2 glass-card rounded-lg p-5 val-border shadow-lg">
                    <h2 class="text-xl font-heading mb-4 text-valLight uppercase">Tüm Maçlar - Performans Trendi</h2>
                    <div class="relative h-72"><canvas id="genTrendChart"></canvas></div>
                </div>
                
                <!-- Top 3 Listesi -->
                <div class="glass-card rounded-lg p-5 border-l-4 border-valGreen shadow-lg flex flex-col gap-6">
                    <div>
                        <h2 class="text-lg font-heading mb-3 text-valGreen uppercase">En İyi 3 Harita (Rating)</h2>
                        <ul id="topMapsList" class="space-y-2"></ul>
                    </div>
                    <div>
                        <h2 class="text-lg font-heading mb-3 text-valRed uppercase">En Verimli 3 Ajan</h2>
                        <ul id="topAgentsList" class="space-y-2"></ul>
                    </div>
                </div>
            </div>

            <!-- Detaylı Ajan İstatistikleri Tablosu -->
            <div class="glass-card rounded-lg p-5 val-border shadow-lg">
                <h2 class="text-xl font-heading mb-4 text-valLight uppercase">Detaylı Ajan İstatistikleri</h2>
                <div class="overflow-x-auto">
                    <table class="w-full text-left border-collapse whitespace-nowrap">
                        <thead>
                            <tr class="text-valGray border-b border-gray-700 text-sm uppercase tracking-wide">
                                <th class="pb-3 pl-2">Ajan</th>
                                <th class="pb-3 text-center">Maç</th>
                                <th class="pb-3 text-center">Ort. Rating <span class="text-xs normal-case">(ACS*)</span></th>
                                <th class="pb-3 text-center">Toplam K / D / A</th>
                                <th class="pb-3 text-center">FK / FD</th>
                                <th class="pb-3 text-center text-valRed">Saldırı Kill</th>
                                <th class="pb-3 text-center text-blue-400">Savunma Kill</th>
                            </tr>
                        </thead>
                        <tbody id="agentSummaryTableBody"></tbody>
                    </table>
                </div>
                <p class="text-xs text-valGray mt-3">* Veri setinde ACS (Ortalama Çatışma Skoru) olmadığı için Rating üzerinden yaklaşık hesaplanmıştır (Rating x 210).</p>
            </div>
        </div>

        <!-- HARİTA SPESİFİK SAYFA -->
        <div id="view-map" class="space-y-8 hidden animate-fade-in">
            <!-- Filtreleme Çubuğu -->
            <div class="glass-card p-4 rounded-lg flex flex-col md:flex-row justify-between items-center gap-4 border border-gray-700">
                <div>
                    <h2 class="text-2xl font-heading text-valRed uppercase" id="currentMapTitle">HARİTA: -</h2>
                    <p class="text-sm text-valGray">Özel istatistikler ve rakip analizi</p>
                </div>
                <div class="flex items-center gap-3">
                    <svg class="w-5 h-5 text-valGray" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 4a1 1 0 011-1h16a1 1 0 011 1v2.586a1 1 0 01-.293.707l-6.414 6.414a1 1 0 00-.293.707V17l-4 4v-6.586a1 1 0 00-.293-.707L3.293 7.293A1 1 0 013 6.586V4z"></path></svg>
                    <label class="text-sm text-valLight uppercase">Rakip Filtresi:</label>
                    <select id="opponentFilter" class="bg-valDark text-valLight border border-gray-600 rounded px-3 py-2 text-sm focus:border-valRed transition-colors"></select>
                </div>
            </div>

            <!-- Harita Özet Kartları -->
            <div class="grid grid-cols-2 md:grid-cols-4 gap-4" id="mapSummaryCards"></div>

            <!-- YENİ: Ajan Bazlı Trend Grafiği -->
            <div class="glass-card rounded-lg p-5 val-border shadow-lg">
                <h2 class="text-xl font-heading mb-4 text-valLight uppercase">Bu Haritada Ajanlara Göre Performans Trendi</h2>
                <div class="relative h-72"><canvas id="mapAgentTrendChart"></canvas></div>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                <!-- Taraf Analizi (Atk vs Def) -->
                <div class="glass-card rounded-lg p-5 val-border shadow-lg">
                    <h2 class="text-xl font-heading mb-4 text-valLight uppercase">Saldırı vs Savunma Başarısı</h2>
                    <div class="relative h-64"><canvas id="mapSideChart"></canvas></div>
                </div>

                <!-- Ajan Verimliliği (Bu Haritada) -->
                <div class="glass-card rounded-lg p-5 val-border shadow-lg">
                    <h2 class="text-xl font-heading mb-4 text-valLight uppercase">Bu Haritada Ajan Kullanımı</h2>
                    <div class="relative h-64"><canvas id="mapAgentChart"></canvas></div>
                </div>
            </div>
        </div>

        <!-- MAÇ DETAY MODALI (POPUP SAYFASI) -->
        <div id="matchDetailModal" class="fixed inset-0 z-[100] hidden items-center justify-center bg-black/80 backdrop-blur-sm p-4 opacity-0 transition-opacity duration-300">
            <div class="glass-card val-border rounded-lg max-w-4xl w-full max-h-[90vh] overflow-y-auto relative shadow-2xl">
                <button onclick="closeMatchModal()" class="absolute top-4 right-4 text-valGray hover:text-valRed transition-colors">
                    <svg class="w-8 h-8" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
                </button>
                <div class="p-6 md:p-8">
                    <h2 id="modalTitle" class="text-3xl md:text-4xl font-heading text-valRed uppercase mb-1">Rakip - Harita</h2>
                    <p id="modalSubtitle" class="text-valGray text-sm font-semibold uppercase tracking-wider mb-6">Ajan: Unknown | Tarih: -</p>

                    <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                        <!-- Toplam -->
                        <div class="bg-valDark rounded-lg p-5 border border-gray-700 shadow-inner">
                            <h3 class="text-valLight font-heading text-xl mb-4 border-b border-gray-600 pb-2">GENEL İSTATİSTİKLER</h3>
                            <ul class="space-y-3 text-sm md:text-base">
                                <li class="flex justify-between items-center"><span class="text-valGray">Rating:</span> <span id="mod_tr" class="text-valGreen font-bold text-lg">-</span></li>
                                <li class="flex justify-between items-center"><span class="text-valGray">K / D / A:</span> <span id="mod_tkda" class="font-bold">-</span></li>
                                <li class="flex justify-between items-center"><span class="text-valGray">KAST:</span> <span id="mod_tkast" class="font-bold">-</span></li>
                                <li class="flex justify-between items-center"><span class="text-valGray">FK / FD:</span> <span id="mod_tfkfd" class="font-bold">-</span></li>
                            </ul>
                        </div>
                        <!-- Saldırı -->
                        <div class="bg-valDark rounded-lg p-5 border border-valRed/30 shadow-inner">
                            <h3 class="text-valRed font-heading text-xl mb-4 border-b border-gray-600 pb-2">SALDIRI (ATK)</h3>
                            <ul class="space-y-3 text-sm md:text-base">
                                <li class="flex justify-between items-center"><span class="text-valGray">Rating:</span> <span id="mod_ar" class="text-valLight font-bold text-lg">-</span></li>
                                <li class="flex justify-between items-center"><span class="text-valGray">K / D / A:</span> <span id="mod_akda" class="font-bold">-</span></li>
                                <li class="flex justify-between items-center"><span class="text-valGray">KAST:</span> <span id="mod_akast" class="font-bold">-</span></li>
                                <li class="flex justify-between items-center"><span class="text-valGray">FK / FD:</span> <span id="mod_afkfd" class="font-bold">-</span></li>
                            </ul>
                        </div>
                        <!-- Savunma -->
                        <div class="bg-valDark rounded-lg p-5 border border-valGreen/30 shadow-inner">
                            <h3 class="text-valGreen font-heading text-xl mb-4 border-b border-gray-600 pb-2">SAVUNMA (DEF)</h3>
                            <ul class="space-y-3 text-sm md:text-base">
                                <li class="flex justify-between items-center"><span class="text-valGray">Rating:</span> <span id="mod_dr" class="text-valLight font-bold text-lg">-</span></li>
                                <li class="flex justify-between items-center"><span class="text-valGray">K / D / A:</span> <span id="mod_dkda" class="font-bold">-</span></li>
                                <li class="flex justify-between items-center"><span class="text-valGray">KAST:</span> <span id="mod_dkast" class="font-bold">-</span></li>
                                <li class="flex justify-between items-center"><span class="text-valGray">FK / FD:</span> <span id="mod_dfkfd" class="font-bold">-</span></li>
                            </ul>
                        </div>
                    </div>
                </div>
            </div>
        </div>

    </main>

    <script>
        // --- 1. VERİ SAKLAMA VE YÖNETİMİ ---
        let parsedData = [];
        let uniqueMaps = [];
        let charts = {};
        let currentActiveView = 'GENEL';
        let currentMapName = '';

        // Valorant Ajan Renk Paleti (Trend grafiğinde farklı ajanlar için)
        const agentColors = [
            '#FF4655', // Kırmızı
            '#00B5B8', // Turkuaz
            '#EAB308', // Sarı
            '#A855F7', // Mor
            '#3B82F6', // Mavi
            '#22C55E', // Yeşil
            '#F97316', // Turuncu
            '#EC4899', // Pembe
            '#ECE8E1'  // Beyazımsı
        ];

        const defaultCSV = `Maç Bilgileri,Harita,Ajan,Toplam R,Toplam K,Toplam D,Toplam A,Toplam KAST,Toplam FK,Toplam FD,Saldırı R,Saldırı K,Saldırı D,Saldırı A,Saldırı KAST,Saldırı FK,Saldırı FD,Savunma R,Savunma K,Savunma D,Savunma A,Savunma KAST,Savunma FK,Savunma FD
Beşiktaş - 10 Ocak,Abyss,Viper,-,22,19,7,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-
Beşiktaş - 10 Ocak,Haven,Cypher,1.28,13,7,5,0.8,3,2,0.72,2,2,0,0.67,1,1,1.41,11,5,5,0.83,2,1
Bushido - 18 Ocak,Breeze,Viper,1.21,22,18,13,0.85,2,2,0.93,8,9,7,0.85,0,1,1.48,14,9,6,0.85,2,1
Bushido - 18 Ocak,Abyss,Viper,1.33,16,10,0,0.82,3,1,0.90,8,7,0,0.75,2,1,2.36,8,3,0,1,1,0
Fire Flux - 24 Ocak,Abyss,Viper,2026-03-01 00:00:00,9,9,4,0.81,0,0,2026-06-01 00:00:00,8,7,3,0.83,0,0,0.94,1,2,1,0.75,0,0
Fire Flux - 24 Ocak,Breeze,Viper,1.45,25,16,10,0.69,2,3,2026-12-01 00:00:00,11,9,3,0.62,1,2,1.77,14,7,7,0.77,1,1
Misa - 26 Ocak,Pearl,Killjoy,2026-02-01 00:00:00,14,16,1,0.58,4,0,0.52,3,10,0,0.33,0,0,1.53,11,6,1,0.83,4,0
Misa - 26 Ocak,Haven,Cypher,0.94,16,16,3,0.61,2,4,1.18,10,6,1,0.67,2,1,0.68,6,10,2,0.55,0,3
Galatasaray - 31 Ocak,Bind,Brimstone,0.62,11,19,11,0.65,2,2,0.83,5,9,7,0.58,1,2,0.38,6,10,4,0.73,1,0
Galatasaray - 31 Ocak,Split,Viper,0.75,12,15,5,0.78,0,2,0.81,6,7,2,0.75,0,0,0.68,6,8,3,0.82,0,2
Eternal Fire - 2 Şubat,Corrode,Viper,0.88,11,15,6,0.65,3,2,0.78,5,8,2,0.5,1,0,0.99,6,7,4,0.82,2,2
Eternal Fire - 2 Şubat,Haven,Cypher,0.52,10,17,5,0.61,1,4,0.31,4,10,2,0.55,0,2,0.72,6,7,3,0.67,1,2
Eternal Fire - 2 Şubat,Abyss,Viper,2026-05-01 00:00:00,17,19,8,0.65,5,3,1.21,11,9,3,0.69,3,1,0.89,6,10,5,0.62,2,2
FUT Academy - 8 Şubat,Corrode,Viper,0.84,16,17,2,0.61,1,1,0.81,8,9,1,0.42,1,0,0.88,8,8,1,0.82,0,1
FUT Academy - 8 Şubat,Breeze,Viper,0.99,16,18,6,0.7,3,2,0.77,7,10,4,0.73,2,0,1.20,9,8,2,0.67,1,2
Misa PO - 28 Şubat,Breeze,Viper,0.88,13,14,5,0.67,3,4,0.86,7,8,3,0.67,1,2,0.91,6,6,2,0.67,2,2
Misa PO - 28 Şubat,Split,Viper,0.73,10,14,2,0.61,1,1,0.89,7,8,2,0.67,0,0,0.41,3,6,0,0.5,1,1
Beşiktaş PO - 2 Mart,Split,Viper,0.62,11,20,8,0.71,1,2,0.86,7,10,4,0.83,0,2,0.38,4,10,4,0.58,1,0
Beşiktaş PO - 2 Mart,Abyss,Viper,0.95,13,16,11,0.79,0,3,0.87,6,7,3,0.75,0,2,2026-03-01 00:00:00,7,9,8,0.83,0,1
Beşiktaş PO - 2 Mart,Corrode,Viper,1.54,21,13,4,0.81,5,2,1.40,11,9,3,0.75,3,1,1.72,10,4,1,0.89,2,1
Beşiktaş PO - 2 Mart,Breeze,Viper,1.23,11,7,3,1,1,1,1.55,5,3,1,1,0,0,2026-12-01 00:00:00,6,4,2,1,1,1
Beşiktaş PO - 2 Mart,Pearl,Sova,0.97,14,15,5,0.83,0,0,0.99,7,5,2,0.91,0,0,0.95,7,10,3,0.75,0,0
MIR (EMEA) - 20 Mart,Split,Omen,0.97,9,8,4,0.85,2,2,0.94,8,7,3,0.83,2,2,1.36,1,1,1,1,0,0
MIR (EMEA) - 20 Mart,Bind,Viper,1.23,5,5,9,0.88,0,0,0.62,8,13,10,0.6,1,2,0.21,3,8,1,0.42,1,2
Mandatory - 22 Mart,Breeze,Omen,0.99,13,15,5,0.71,2,1,1.15,8,7,4,0.75,1,0,0.79,5,8,1,0.67,1,1
Mandatory - 22 Mart,Pearl,Killjoy,0.78,11,12,2,0.74,0,3,0.82,6,6,1,0.75,0,2,0.73,5,6,1,0.73,0,1
GnG (EMEA) - 23 Mart,Split,Omen,0.90,11,13,9,0.9,1,0,0.84,6,7,6,0.92,1,0,0.98,5,6,3,0.88,0,0
GnG (EMEA) - 23 Mart,Haven,Omen,2026-05-01 00:00:00,11,10,7,0.84,2,4,0.89,7,8,3,0.83,0,2,1.31,4,2,4,0.86,2,2
Ghosts (S2) - 5 Nisan,Bind,Viper,0.78,9,12,4,0.74,0,1,0.79,5,7,1,0.58,0,1,0.74,4,5,3,1,0,0
Ghosts (S2) - 5 Nisan,Split,Omen,1.34,15,8,6,0.72,1,3,1.17,3,2,2,0.83,0,1,1.43,12,6,4,0.67,1,2
Fire Flux (S2) - 7 Nisan,Split,Omen,2.42,8,1,2,1,1,0,2026-05-01 00:00:00,9,9,5,0.83,1,1,1.45,17,10,7,0.88,2,1
Fire Flux (S2) - 7 Nisan,Lotus,Omen,0.90,8,10,8,0.79,1,0,0.88,7,7,4,0.79,1,0,1.33,12,10,4,0.93,2,1
FUT Aca (S2) - 12 Nisan,Split,Omen,1.40,20,12,6,0.92,3,1,1.81,16,6,4,0.92,3,1,0.70,4,6,2,0.92,0,0
FUT Aca (S2) - 12 Nisan,Lotus,Omen,2026-04-01 00:00:00,11,10,6,0.83,2,1,1.37,5,2,0,1,0,0,0.87,6,8,6,0.75,2,1
REBORN vs Galatasaray Esports - 19 Nisan,Split,Omen,1.55,14,7,12,0.94,1,2,1.20,8,5,8,0.92,0,2,2.61,6,2,4,1,1,0
Galatasaray Esports - 19 Nisan,Bind,Omen,1.23,22,15,3,0.71,3,4,1.34,12,7,1,0.58,3,2,2026-12-01 00:00:00,10,8,2,0.83,0,2
Misa Esports - 27 Nisan,Haven,Omen,1.34,18,11,9,0.74,4,3,1.53,8,4,0,0.71,2,1,1.23,10,7,9,0.75,2,2
Misa Esports - 27 Nisan,Split,Omen,1.28,23,15,6,0.81,4,4,1.15,11,9,1,0.85,2,3,1.41,12,6,5,0.77,2,1
Beşiktaş Esports - 3 Mayıs,Split,Omen,1.00,14,10,8,0.83,1,0,0.77,7,7,4,0.75,1,0,1.47,7,3,4,1,0,0
Beşiktaş Esports - 3 Mayıs,Lotus,Omen,1.21,24,19,14,0.72,1,3,0.86,10,12,10,0.69,0,2,1.55,14,7,4,0.75,1,1
Beşiktaş Esports - 3 Mayıs,Breeze,Viper,1.28,16,9,4,0.9,1,1,1.19,6,2,2,0.89,0,0,1.35,10,7,2,0.92,1,1






















































































































































































































































































































































































































































































































































































































































































































































































































































































































`;

        // --- 2. VERİ TEMİZLEME VE NORMALİZASYON ---
        function normalizeRating(val, k, d, a) {
            if (val === '-' || val === '' || val === undefined) return (k + (a * 0.3)) / Math.max(1, d);
            let num = parseFloat(val.toString().replace(',', '.'));
            if (isNaN(num)) return (k + (a * 0.3)) / Math.max(1, d);
            if (num > 0 && num <= 3) return num;
            let proxyRating = (k + (a * 0.3)) / Math.max(1, d);
            return Math.min(2.0, Math.max(0.4, proxyRating));
        }

        function normalizeKAST(val) {
            if (val === '-' || val === '' || val === undefined) return 70;
            let num = parseFloat(val.toString().replace(',', '.'));
            if (isNaN(num)) return 70;
            if (num <= 1) return num * 100;
            return num;
        }

        function extractOpponent(matchInfo) {
            if(!matchInfo) return "Bilinmeyen";
            let parts = matchInfo.split('-');
            return parts[0].trim();
        }

        function processCSV(csvText) {
            const lines = csvText.trim().split('\n');
            const headers = lines[0].split(',').map(h => h.trim());
            
            const idx = {
                match: headers.findIndex(h => h.includes('Maç Bilgileri')),
                map: headers.findIndex(h => h.includes('Harita')),
                agent: headers.findIndex(h => h.includes('Ajan')),
                tr: headers.findIndex(h => h === 'Toplam R'),
                tk: headers.findIndex(h => h === 'Toplam K'),
                td: headers.findIndex(h => h === 'Toplam D'),
                ta: headers.findIndex(h => h === 'Toplam A'),
                kast: headers.findIndex(h => h.includes('KAST') && h.includes('Toplam')),
                fk: headers.findIndex(h => h === 'Toplam FK'),
                fd: headers.findIndex(h => h === 'Toplam FD'),
                atkR: headers.findIndex(h => h === 'Saldırı R'),
                defR: headers.findIndex(h => h === 'Savunma R'),
                atkK: headers.findIndex(h => h === 'Saldırı K'),
                atkD: headers.findIndex(h => h === 'Saldırı D'),
                atkA: headers.findIndex(h => h === 'Saldırı A'),
                atkKast: headers.findIndex(h => h === 'Saldırı KAST'),
                atkFk: headers.findIndex(h => h === 'Saldırı FK'),
                atkFd: headers.findIndex(h => h === 'Saldırı FD'),
                defK: headers.findIndex(h => h === 'Savunma K'),
                defD: headers.findIndex(h => h === 'Savunma D'),
                defA: headers.findIndex(h => h === 'Savunma A'),
                defKast: headers.findIndex(h => h === 'Savunma KAST'),
                defFk: headers.findIndex(h => h === 'Savunma FK'),
                defFd: headers.findIndex(h => h === 'Savunma FD')
            };

            let data = [];
            let mapsSet = new Set();

            for (let i = 1; i < lines.length; i++) {
                const lineStr = lines[i].trim();
                // Sadece virgüllerden oluşan boş satırları atla
                if (!lineStr || lineStr.replace(/,/g, '').trim() === '') continue;

                const row = lineStr.split(',');
                if (row.length < 5) continue;

                // Maç veya Harita hücresi boş olan (hayalet) satırları atla
                const rawMatch = row[idx.match];
                const rawMap = row[idx.map];
                if (!rawMatch || rawMatch.trim() === '' || !rawMap || rawMap.trim() === '') continue;

                const k = parseInt(row[idx.tk]) || 0;
                const d = parseInt(row[idx.td]) || 0;
                const a = parseInt(row[idx.ta]) || 0;

                const rating = normalizeRating(row[idx.tr], k, d, a);
                const matchFull = rawMatch.trim();
                const mapStr = rawMap.trim();
                
                mapsSet.add(mapStr);

                data.push({
                    rawMatch: matchFull,
                    opponent: extractOpponent(matchFull),
                    date: matchFull.split('-')[1]?.trim() || '',
                    map: mapStr,
                    agent: row[idx.agent] || 'Unknown',
                    rating: parseFloat(rating.toFixed(2)),
                    k: k, d: d, a: a,
                    kast: Math.round(normalizeKAST(row[idx.kast])),
                    fk: parseInt(row[idx.fk]) || 0,
                    fd: parseInt(row[idx.fd]) || 0,
                    atkR: parseFloat(normalizeRating(row[idx.atkR], k/2, d/2, a/2).toFixed(2)),
                    defR: parseFloat(normalizeRating(row[idx.defR], k/2, d/2, a/2).toFixed(2)),
                    atkK: idx.atkK !== -1 ? parseInt(row[idx.atkK]) || 0 : 0,
                    atkD: idx.atkD !== -1 ? parseInt(row[idx.atkD]) || 0 : 0,
                    atkA: idx.atkA !== -1 ? parseInt(row[idx.atkA]) || 0 : 0,
                    atkKast: idx.atkKast !== -1 ? Math.round(normalizeKAST(row[idx.atkKast])) : 0,
                    atkFk: idx.atkFk !== -1 ? parseInt(row[idx.atkFk]) || 0 : 0,
                    atkFd: idx.atkFd !== -1 ? parseInt(row[idx.atkFd]) || 0 : 0,
                    defK: idx.defK !== -1 ? parseInt(row[idx.defK]) || 0 : 0,
                    defD: idx.defD !== -1 ? parseInt(row[idx.defD]) || 0 : 0,
                    defA: idx.defA !== -1 ? parseInt(row[idx.defA]) || 0 : 0,
                    defKast: idx.defKast !== -1 ? Math.round(normalizeKAST(row[idx.defKast])) : 0,
                    defFk: idx.defFk !== -1 ? parseInt(row[idx.defFk]) || 0 : 0,
                    defFd: idx.defFd !== -1 ? parseInt(row[idx.defFd]) || 0 : 0
                });
            }
            parsedData = data;
            uniqueMaps = Array.from(mapsSet).sort();
        }

        // --- 3. UI VE YÖNLENDİRME ---
        function initApp(csvText) {
            processCSV(csvText);
            Chart.defaults.color = '#ECE8E1';
            Chart.defaults.font.family = "'Inter', sans-serif";
            
            buildNavigation();
            switchView('GENEL');
        }

        function buildNavigation() {
            const nav = document.getElementById('navMenu');
            let html = `<button onclick="switchView('GENEL')" class="tab-btn active pb-2 px-1 text-sm font-heading tracking-wide uppercase" data-target="GENEL">GENEL ÖZET</button>`;
            uniqueMaps.forEach(map => {
                html += `<button onclick="switchView('${map}')" class="tab-btn pb-2 px-1 text-sm text-valGray font-heading tracking-wide uppercase" data-target="${map}">${map}</button>`;
            });
            nav.innerHTML = html;
        }

        function switchView(targetView) {
            currentActiveView = targetView;

            document.querySelectorAll('.tab-btn').forEach(btn => {
                if(btn.dataset.target === targetView) {
                    btn.classList.add('active'); btn.classList.remove('text-valGray');
                } else {
                    btn.classList.remove('active'); btn.classList.add('text-valGray');
                }
            });

            if (targetView === 'GENEL') {
                document.getElementById('view-general').classList.remove('hidden');
                document.getElementById('view-map').classList.add('hidden');
                renderGeneralView();
            } else {
                document.getElementById('view-general').classList.add('hidden');
                document.getElementById('view-map').classList.remove('hidden');
                currentMapName = targetView;
                renderMapView(targetView);
            }
        }

        function destroyChart(chartName) {
            if(charts[chartName]) { charts[chartName].destroy(); }
        }

        // --- 4. GENEL SAYFA RENDER ---
        function renderGeneralView() {
            const data = parsedData;
            const totalMatches = data.length;
            if(totalMatches === 0) return;

            const avgRating = (data.reduce((s, m) => s + m.rating, 0) / totalMatches).toFixed(2);
            const avgKast = Math.round(data.reduce((s, m) => s + m.kast, 0) / totalMatches);
            const totalFK = data.reduce((s, m) => s + m.fk, 0);
            const totalFD = data.reduce((s, m) => s + m.fd, 0);
            const totalK = data.reduce((s, m) => s + m.k, 0);
            const totalD = data.reduce((s, m) => s + m.d, 0);
            const kdRatio = (totalK / Math.max(1, totalD)).toFixed(2);

            document.getElementById('genSummaryCards').innerHTML = 
                createCardHTML('Genel Ortalama Rating', avgRating) +
                createCardHTML('K/D Oranı', kdRatio) +
                createCardHTML('Ortalama KAST', `%${avgKast}`) +
                createCardHTML('Açılış Skoru (FK-FD)', `${totalFK - totalFD > 0 ? '+' : ''}${totalFK - totalFD}`, totalFK >= totalFD ? 'text-valGreen' : 'text-valRed');

            destroyChart('genTrend');
            const ctxTrend = document.getElementById('genTrendChart').getContext('2d');
            charts.genTrend = new Chart(ctxTrend, {
                type: 'line',
                data: {
                    labels: data.map(m => `${m.opponent} (${m.map})`),
                    datasets: [{
                        label: 'Maç Ratingi',
                        data: data.map(m => m.rating),
                        borderColor: '#FF4655',
                        backgroundColor: 'rgba(255, 70, 85, 0.1)',
                        borderWidth: 3,
                        pointBackgroundColor: '#00B5B8',
                        pointRadius: 5, fill: true, tension: 0.3
                    }]
                },
                options: { 
                    responsive: true, 
                    maintainAspectRatio: false, 
                    scales: { y: { beginAtZero: false, suggestedMin: 0.5, suggestedMax: 1.5 } },
                    onClick: (e, activeEls) => {
                        if (activeEls.length > 0) {
                            const index = activeEls[0].index;
                            openMatchModal(data[index]);
                        }
                    },
                    plugins: {
                        tooltip: {
                            callbacks: {
                                footer: () => '👉 Detaylı analiz için grafiğe tıkla!'
                            }
                        }
                    }
                }
            });

            renderTopLists(data);
            renderAgentTable(data);
        }

        function renderTopLists(data) {
            let mapStats = {}, agentStats = {};
            data.forEach(m => {
                if(!mapStats[m.map]) mapStats[m.map] = {r:0, c:0};
                mapStats[m.map].r += m.rating; mapStats[m.map].c++;

                if(!agentStats[m.agent]) agentStats[m.agent] = {r:0, c:0};
                agentStats[m.agent].r += m.rating; agentStats[m.agent].c++;
            });

            const topMaps = Object.keys(mapStats)
                .map(k => ({ name: k, val: (mapStats[k].r / mapStats[k].c).toFixed(2), count: mapStats[k].c }))
                .sort((a,b) => b.val - a.val).slice(0,3);

            const topAgents = Object.keys(agentStats)
                .map(k => ({ name: k, val: (agentStats[k].r / agentStats[k].c).toFixed(2), count: agentStats[k].c }))
                .sort((a,b) => b.val - a.val).slice(0,3);

            document.getElementById('topMapsList').innerHTML = topMaps.map((m, i) => `
                <li class="flex justify-between items-center bg-valDark p-2 rounded cursor-pointer group hover:bg-valCard transition-colors" onclick="openMapMatchesModal('${m.name}')" title="${m.name} maçlarını görmek için tıkla">
                    <span class="text-valLight font-bold group-hover:text-valRed transition-colors">#${i+1} ${m.name} <span class="text-xs text-valGray font-normal">(${m.count} Maç)</span></span>
                    <span class="text-valGreen font-heading">${m.val}</span>
                </li>`).join('') || '<li class="text-valGray">Yeterli veri yok</li>';
            
            document.getElementById('topAgentsList').innerHTML = topAgents.map((a, i) => `
                <li class="flex justify-between items-center bg-valDark p-2 rounded cursor-pointer group hover:bg-valCard transition-colors" onclick="openAgentMatchesModal('${a.name}')" title="${a.name} maçlarını görmek için tıkla">
                    <span class="text-valLight font-bold group-hover:text-valRed transition-colors">#${i+1} ${a.name} <span class="text-xs text-valGray font-normal">(${a.count} Maç)</span></span>
                    <span class="text-valRed font-heading">${a.val}</span>
                </li>`).join('') || '<li class="text-valGray">Yeterli veri yok</li>';
        }

        function renderAgentTable(data) {
            const agentStats = {};
            data.forEach(m => {
                if(!agentStats[m.agent]) {
                    agentStats[m.agent] = { c: 0, r: 0, k: 0, d: 0, a: 0, fk: 0, fd: 0, atkK: 0, defK: 0 };
                }
                agentStats[m.agent].c++; agentStats[m.agent].r += m.rating;
                agentStats[m.agent].k += m.k; agentStats[m.agent].d += m.d; agentStats[m.agent].a += m.a;
                agentStats[m.agent].fk += m.fk; agentStats[m.agent].fd += m.fd;
                agentStats[m.agent].atkK += m.atkK; agentStats[m.agent].defK += m.defK;
            });

            let html = '';
            Object.keys(agentStats).sort((a,b) => agentStats[b].c - agentStats[a].c).forEach(agent => {
                const s = agentStats[agent];
                const avgR = (s.r / s.c).toFixed(2);
                const proxyAcs = Math.round((s.r / s.c) * 210); 
                html += `
                    <tr class="border-b border-gray-800 hover:bg-valCard transition-colors">
                        <td class="py-3 pl-2 font-bold text-valLight">${agent}</td>
                        <td class="py-3 text-center text-valGray font-heading text-lg">${s.c}</td>
                        <td class="py-3 text-center"><span class="text-valGreen font-bold">${avgR}</span> <span class="text-xs text-valGray">(${proxyAcs})</span></td>
                        <td class="py-3 text-center tracking-wider">${s.k} / ${s.d} / ${s.a}</td>
                        <td class="py-3 text-center">${s.fk} / ${s.fd}</td>
                        <td class="py-3 text-center text-valRed font-bold">${s.atkK}</td>
                        <td class="py-3 text-center text-blue-400 font-bold">${s.defK}</td>
                    </tr>`;
            });
            document.getElementById('agentSummaryTableBody').innerHTML = html || '<tr><td colspan="7" class="text-center text-valGray py-4">Veri bulunamadı</td></tr>';
        }

        // --- 5. HARİTA (MAP) SAYFASI RENDER VE FİLTRELEME ---
        function renderMapView(mapName) {
            document.getElementById('currentMapTitle').innerText = `HARİTA: ${mapName}`;
            
            const mapDataAll = parsedData.filter(m => m.map === mapName);
            const opponents = [...new Set(mapDataAll.map(m => m.opponent))].sort();
            
            const filterSelect = document.getElementById('opponentFilter');
            let optionsHtml = `<option value="ALL">Tüm Rakipler (Genel)</option>`;
            opponents.forEach(opp => { optionsHtml += `<option value="${opp}">${opp}</option>`; });
            filterSelect.innerHTML = optionsHtml;
            filterSelect.value = "ALL";

            filterSelect.onchange = () => updateMapCharts(mapName, filterSelect.value);
            updateMapCharts(mapName, "ALL");
        }

        function updateMapCharts(mapName, opponentFilter) {
            let mapData = parsedData.filter(m => m.map === mapName);
            if (opponentFilter !== "ALL") {
                mapData = mapData.filter(m => m.opponent === opponentFilter);
            }

            if(mapData.length === 0) {
                document.getElementById('mapSummaryCards').innerHTML = '<p class="text-valRed col-span-4">Bu filtreye uygun maç bulunamadı.</p>';
                destroyChart('mapSide'); destroyChart('mapAgent'); destroyChart('mapAgentTrend');
                return;
            }

            const avgR = (mapData.reduce((s,m) => s + m.rating, 0) / mapData.length).toFixed(2);
            const avgAtk = (mapData.reduce((s,m) => s + m.atkR, 0) / mapData.length).toFixed(2);
            const avgDef = (mapData.reduce((s,m) => s + m.defR, 0) / mapData.length).toFixed(2);
            const winCondition = avgAtk > avgDef ? 'Saldırı Odaklı' : 'Savunma Odaklı';

            document.getElementById('mapSummaryCards').innerHTML = 
                createCardHTML('Filtrelenen Maç', mapData.length) +
                createCardHTML('Ortalama Rating', avgR) +
                createCardHTML('En Güçlü Taraf', winCondition, 'text-valGreen') +
                createCardHTML('Atk / Def Rating', `${avgAtk} / ${avgDef}`);

            // YENİ: Ajanlara Göre Trend Grafiği (Line Chart with Multiple Lines)
            destroyChart('mapAgentTrend');
            const uniqueAgentsInMap = [...new Set(mapData.map(m => m.agent))];
            
            const datasets = uniqueAgentsInMap.map((agent, index) => {
                // Her ajan için sadece kendi oynandığı maçlarda rating gösterilir, diğerlerinde 'null' döner ki çizgi kopsun/ayrılsın
                const agentDataPoints = mapData.map(m => m.agent === agent ? m.rating : null);
                const color = agentColors[index % agentColors.length];
                
                return {
                    label: agent,
                    data: agentDataPoints,
                    borderColor: color,
                    backgroundColor: color,
                    borderWidth: 3,
                    pointBackgroundColor: color,
                    pointRadius: 6,
                    pointHoverRadius: 8,
                    spanGaps: true, // Eğer araya başka ajan girerse çizgiyi birleştirir
                    tension: 0.2
                };
            });

            const ctxAgentTrend = document.getElementById('mapAgentTrendChart').getContext('2d');
            charts.mapAgentTrend = new Chart(ctxAgentTrend, {
                type: 'line',
                data: {
                    labels: mapData.map(m => opponentFilter === "ALL" ? m.opponent : m.date || 'Tarihsiz'),
                    datasets: datasets
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: { y: { suggestedMin: 0.5, suggestedMax: 1.5 } },
                    onClick: (e, activeEls) => {
                        if (activeEls.length > 0) {
                            const index = activeEls[0].index;
                            openMatchModal(mapData[index]);
                        }
                    },
                    plugins: {
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    return `${context.dataset.label} Rating: ${context.parsed.y}`;
                                },
                                footer: () => '👉 Detaylı analiz için grafiğe tıkla!'
                            }
                        }
                    }
                }
            });

            // Taraf Analizi Grafiği
            destroyChart('mapSide');
            const ctxSide = document.getElementById('mapSideChart').getContext('2d');
            charts.mapSide = new Chart(ctxSide, {
                type: 'bar',
                data: {
                    labels: mapData.map(m => opponentFilter === "ALL" ? m.opponent : m.date || 'Tarihsiz'),
                    datasets: [
                        { label: 'Saldırı', data: mapData.map(m => m.atkR), backgroundColor: '#FF4655', borderRadius: 4 },
                        { label: 'Savunma', data: mapData.map(m => m.defR), backgroundColor: '#00B5B8', borderRadius: 4 }
                    ]
                },
                options: { 
                    responsive: true, 
                    maintainAspectRatio: false,
                    onClick: (e, activeEls) => {
                        if (activeEls.length > 0) {
                            const index = activeEls[0].index;
                            openMatchModal(mapData[index]);
                        }
                    },
                    plugins: {
                        tooltip: {
                            callbacks: {
                                footer: () => '👉 Detaylı analiz için grafiğe tıkla!'
                            }
                        }
                    }
                }
            });

            // Ajan Verimliliği Grafiği
            destroyChart('mapAgent');
            let agentMapStats = {};
            mapData.forEach(m => {
                if(!agentMapStats[m.agent]) agentMapStats[m.agent] = {r:0, c:0};
                agentMapStats[m.agent].r += m.rating;
                agentMapStats[m.agent].c++;
            });
            
            const agentLabels = Object.keys(agentMapStats);
            const agentData = agentLabels.map(a => (agentMapStats[a].r / agentMapStats[a].c).toFixed(2));

            const ctxAgent = document.getElementById('mapAgentChart').getContext('2d');
            charts.mapAgent = new Chart(ctxAgent, {
                type: 'polarArea',
                data: {
                    labels: agentLabels,
                    datasets: [{
                        data: agentData,
                        backgroundColor: ['rgba(255, 70, 85, 0.7)', 'rgba(0, 181, 184, 0.7)', 'rgba(236, 232, 225, 0.7)', 'rgba(139, 151, 143, 0.7)'],
                        borderColor: '#0F1923',
                        borderWidth: 2
                    }]
                },
                options: { responsive: true, maintainAspectRatio: false, scales: { r: { grid: { color: 'rgba(139, 151, 143, 0.2)' }, ticks: { backdropColor: 'transparent', color: '#8B978F' } } } }
            });
        }

        function createCardHTML(title, value, colorClass = 'text-white') {
            return `
                <div class="glass-card p-4 rounded-lg val-border">
                    <p class="text-valGray text-xs font-semibold uppercase tracking-wider">${title}</p>
                    <p class="text-2xl md:text-3xl font-heading ${colorClass} mt-2">${value}</p>
                </div>`;
        }

        // --- MAÇ DETAY MODALI FONKSİYONLARI ---
        function openMatchModal(match) {
            if (!match) return;
            
            // Başlıklar
            document.getElementById('modalTitle').innerText = `${match.opponent} - ${match.map}`;
            document.getElementById('modalSubtitle').innerText = `AJAN: ${match.agent} | TARİH: ${match.date || 'Belirtilmedi'}`;

            // Genel İstatistikler
            document.getElementById('mod_tr').innerText = match.rating;
            document.getElementById('mod_tkda').innerText = `${match.k} / ${match.d} / ${match.a}`;
            document.getElementById('mod_tkast').innerText = `%${match.kast}`;
            document.getElementById('mod_tfkfd').innerText = `${match.fk} / ${match.fd}`;

            // Saldırı İstatistikleri
            document.getElementById('mod_ar').innerText = match.atkR;
            document.getElementById('mod_akda').innerText = `${match.atkK} / ${match.atkD} / ${match.atkA}`;
            document.getElementById('mod_akast').innerText = match.atkKast ? `%${match.atkKast}` : '-';
            document.getElementById('mod_afkfd').innerText = `${match.atkFk} / ${match.atkFd}`;

            // Savunma İstatistikleri
            document.getElementById('mod_dr').innerText = match.defR;
            document.getElementById('mod_dkda').innerText = `${match.defK} / ${match.defD} / ${match.defA}`;
            document.getElementById('mod_dkast').innerText = match.defKast ? `%${match.defKast}` : '-';
            document.getElementById('mod_dfkfd').innerText = `${match.defFk} / ${match.defFd}`;

            // Modalı Göster
            const modal = document.getElementById('matchDetailModal');
            modal.classList.remove('hidden');
            modal.classList.add('flex');
            setTimeout(() => modal.classList.remove('opacity-0'), 10);
        }

        function closeMatchModal() {
            const modal = document.getElementById('matchDetailModal');
            modal.classList.add('opacity-0');
            setTimeout(() => {
                modal.classList.add('hidden');
                modal.classList.remove('flex');
            }, 300);
        }

        // --- YENİ: AJAN MAÇLARI MODALI FONKSİYONLARI ---
        function openAgentMatchesModal(agentName) {
            document.getElementById('agentModalTitle').innerText = `${agentName} MAÇ GEÇMİŞİ`;
            
            // Tüm maçlardan sadece tıklanan ajanın oynandığı maçları ayırıyoruz
            const agentMatches = parsedData.filter(m => m.agent === agentName);
            document.getElementById('agentModalSubtitle').innerText = `TOPLAM ${agentMatches.length} MAÇ BULUNDU`;

            let html = '';
            agentMatches.forEach(m => {
                // Tıklandığında direkt o maçı açabilmek için global ID'sini buluyoruz
                const globalIndex = parsedData.indexOf(m);
                const isWin = m.rating >= 1; 
                
                html += `
                    <tr class="border-b border-gray-800 hover:bg-valCard transition-colors cursor-pointer group" onclick="openMatchModalByDataIndex(${globalIndex})">
                        <td class="p-3 text-valLight font-bold">${m.opponent} <br><span class="text-xs text-valGray font-normal">${m.date || 'Tarih Yok'}</span></td>
                        <td class="p-3 text-valLight tracking-wider">${m.map}</td>
                        <td class="p-3 text-center font-heading text-lg ${isWin ? 'text-valGreen' : 'text-valRed'}">${m.rating}</td>
                        <td class="p-3 text-center text-valLight">${m.k} / ${m.d} / ${m.a}</td>
                        <td class="p-3 text-center text-valGray group-hover:text-valRed transition-colors">
                            <svg class="w-6 h-6 mx-auto" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 5l7 7-7 7M5 5l7 7-7 7"></path></svg>
                        </td>
                    </tr>
                `;
            });

            document.getElementById('agentMatchesListBody').innerHTML = html;

            const modal = document.getElementById('agentMatchesModal');
            modal.classList.remove('hidden');
            modal.classList.add('flex');
            setTimeout(() => modal.classList.remove('opacity-0'), 10);
        }

        function closeAgentMatchesModal() {
            const modal = document.getElementById('agentMatchesModal');
            modal.classList.add('opacity-0');
            setTimeout(() => {
                modal.classList.add('hidden');
                modal.classList.remove('flex');
            }, 300);
        }

        function openMatchModalByDataIndex(index) {
            // Ajan listesinin içinden spesifik bir maça tıklandığında üstüne (z-index 100 ile) Maç Detay modunu açar
            openMatchModal(parsedData[index]);
        }

        // --- BAŞLATMA ---
        window.onload = () => {
            initApp(defaultCSV);

            document.getElementById('csvFileInput').addEventListener('change', function(e) {
                const file = e.target.files[0];
                if (!file) return;

                const fileName = file.name.toLowerCase();
                const reader = new FileReader();

                if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
                    reader.onload = function(e) {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, {type: 'array'});
                        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                        const csvText = XLSX.utils.sheet_to_csv(worksheet);
                        initApp(csvText);
                    };
                    reader.readAsArrayBuffer(file);
                } else {
                    reader.onload = function(e) { initApp(e.target.result); };
                    reader.readAsText(file);
                }
            });
        };
    </script>
</body>
</html>







```
