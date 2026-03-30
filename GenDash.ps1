# ============================================================================
# Generate-Dashboard.ps1
# ============================================================================
# SYNOPSIS: 
#   Reads CSV logs, queries Active Directory, and generates an HTML dashboard.
#   POLLING MODE: Checks for log file changes every 5 seconds.
# ============================================================================

# --- PREREQUISITES ---
try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Warning "ActiveDirectory module not found. The 'Active Lockouts' table may be empty."
}

# --- CONFIGURATION ---
$LogPath = "\\vm-isserver\toolkit\[IT Toolkit]\HDCompanion\Logs"
$OutputHtml = "\\vm-simplify\custom\fileuploads\acctdashboard.html" 
$IgnoredUsers = @("support", "admbrian", "guest")

# --- HTML TEMPLATE ---
$HtmlTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="refresh" content="120">
    <title>Pelican CU Account Monitor Dashboard</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://unpkg.com/@phosphor-icons/web"></script>
    <style>
        .custom-scrollbar::-webkit-scrollbar { width: 8px; height: 8px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #f1f1f1; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #c1c1c1; border-radius: 4px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #a8a8a8; }
        tr.cursor-pointer:hover { background-color: #f8fafc; }
    </style>
</head>
<body class="bg-slate-100 text-slate-800 font-sans h-screen flex flex-col overflow-hidden">
    <nav class="bg-[#005596] text-white shadow-md z-10 shrink-0">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div class="flex items-center justify-between h-16">
                <div class="flex items-center gap-3">
                    <i class="ph ph-shield-check text-2xl"></i>
                    <span class="font-bold text-lg tracking-wide">PELICAN CREDIT UNION</span>
                    <span class="bg-blue-800 text-xs px-2 py-1 rounded text-blue-200">Account Monitor</span>
                </div>
                <div class="flex items-center gap-4">
                    <div class="flex items-center gap-2 bg-blue-700/50 px-3 py-1 rounded text-xs font-mono text-blue-100 border border-blue-500/30">
                        <i class="ph ph-timer"></i>
                        <span>Refreshing in <span id="refresh-timer" class="font-bold text-white">120</span>s</span>
                    </div>
                    
                    <div class="text-sm opacity-80" id="last-updated">Report Generated: {{GENERATED_TIME}}</div>
                    <button id="yearToggleBtn" onclick="toggleYearView()" class="bg-white/10 hover:bg-white/20 px-3 py-1 rounded text-xs font-medium border border-white/20 transition-colors flex items-center gap-2">
                        <i class="ph ph-calendar-plus"></i> Show All History
                    </button>
                    <button onclick="resetFilters()" class="bg-white/10 hover:bg-white/20 px-3 py-1 rounded text-xs font-medium border border-white/20 transition-colors flex items-center gap-2">
                        <i class="ph ph-arrows-counter-clockwise"></i> Reset View
                    </button>
                </div>
            </div>
        </div>
    </nav>

    <div class="flex-1 overflow-auto custom-scrollbar p-6">
        <div class="max-w-7xl mx-auto space-y-6">
            
            <!-- KPI Cards -->
            <div class="grid grid-cols-1 md:grid-cols-4 gap-6">
                <div class="bg-white p-6 rounded-lg shadow-sm border border-slate-200 flex items-center justify-between">
                    <div>
                        <p class="text-xs font-medium text-slate-500 uppercase tracking-wider">Currently Locked</p>
                        <p class="text-3xl font-bold text-red-600 mt-1" id="kpi-current-locked">0</p>
                    </div>
                    <div class="p-3 bg-red-50 rounded-full text-red-600"><i class="ph ph-lock-key text-2xl"></i></div>
                </div>
                <div class="bg-white p-6 rounded-lg shadow-sm border border-slate-200 flex items-center justify-between">
                    <div>
                        <p id="kpi-total-label" class="text-xs font-medium text-slate-500 uppercase tracking-wider">YTD Lockouts</p>
                        <p class="text-3xl font-bold text-[#005596] mt-1" id="kpi-total-events">0</p>
                        <p class="text-[10px] text-slate-400 mt-1">{{DATE_RANGE}}</p>
                    </div>
                    <div class="p-3 bg-blue-50 rounded-full text-[#005596]"><i class="ph ph-list-dashes text-2xl"></i></div>
                </div>
                <div class="bg-white p-6 rounded-lg shadow-sm border border-slate-200 flex items-center justify-between">
                    <div>
                        <p class="text-xs font-medium text-slate-500 uppercase tracking-wider">Agent Unlocks</p>
                        <p class="text-3xl font-bold text-emerald-600 mt-1" id="kpi-manual-unlocks">0</p>
                        <p class="text-[10px] text-slate-400 mt-1">vs <span id="kpi-auto-unlocks">0</span> Auto</p>
                    </div>
                    <div class="p-3 bg-emerald-50 rounded-full text-emerald-600"><i class="ph ph-wrench text-2xl"></i></div>
                </div>
                <div class="bg-white p-6 rounded-lg shadow-sm border border-slate-200 flex items-center justify-between">
                    <div>
                        <p class="text-xs font-medium text-slate-500 uppercase tracking-wider">Agent Time Spent</p>
                        <p class="text-3xl font-bold text-orange-600 mt-1" id="kpi-agent-time">0m</p>
                        <p class="text-[10px] text-slate-400 mt-1">Est. 5 min per unlock</p>
                    </div>
                    <div class="p-3 bg-orange-50 rounded-full text-orange-600"><i class="ph ph-clock text-2xl"></i></div>
                </div>
            </div>

            <!-- Current Lockouts Table -->
            <div class="bg-white rounded-lg shadow-sm border border-slate-200 overflow-hidden">
                <div class="px-6 py-4 border-b border-slate-200 bg-slate-50 flex justify-between items-center">
                    <h3 class="font-bold text-slate-700 flex items-center gap-2">
                        <i class="ph ph-siren text-red-500"></i> Active Lockouts (AD Query)
                    </h3>
                    <span class="text-xs text-slate-500 bg-white border px-2 py-1 rounded">Real-time status</span>
                </div>
                <div class="overflow-x-auto">
                    <table class="w-full text-left text-sm">
                        <thead class="bg-slate-50 text-slate-500 font-medium border-b border-slate-200">
                            <tr>
                                <th class="px-6 py-3">Username</th>
                                <th class="px-6 py-3">Lockout Time</th>
                                <th class="px-6 py-3">Source / Computer</th>
                                <th class="px-6 py-3">Details</th>
                                <th class="px-6 py-3">Status</th>
                            </tr>
                        </thead>
                        <tbody id="active-lockouts-body" class="divide-y divide-slate-100"></tbody>
                    </table>
                </div>
            </div>

            <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                <div class="bg-white p-6 rounded-lg shadow-sm border border-slate-200 lg:col-span-2">
                    <div class="flex justify-between items-center mb-4">
                        <h3 class="text-lg font-bold text-slate-700">Lockout Activity Trend</h3>
                        <span class="text-xs text-slate-400 italic">Click point to filter logs</span>
                    </div>
                    <div class="relative h-64 w-full"><canvas id="lockoutTrendChart"></canvas></div>
                </div>
                <div class="bg-white p-6 rounded-lg shadow-sm border border-slate-200">
                    <div class="flex justify-between items-center mb-4">
                        <h3 class="text-lg font-bold text-slate-700">Top Offenders</h3>
                        <span class="text-xs text-slate-400 italic">Click bar to filter logs</span>
                    </div>
                    <div class="relative h-64 w-full"><canvas id="userBarChart"></canvas></div>
                </div>
            </div>

            <div class="bg-white rounded-lg shadow-sm border border-slate-200 overflow-hidden" id="logSection">
                <div class="px-6 py-4 border-b border-slate-200 bg-slate-50 flex justify-between items-center">
                    <h3 class="font-bold text-slate-700">Recent Log Activity</h3>
                    <div id="activeFilterBadge" class="hidden bg-blue-100 text-blue-800 text-xs px-2 py-1 rounded-full font-medium border border-blue-200 flex items-center gap-1">
                        <span>Filtered: <b id="filterLabel"></b></span>
                        <button onclick="resetFilters()" class="hover:text-blue-600"><i class="ph ph-x"></i></button>
                    </div>
                </div>
                <div class="overflow-x-auto max-h-96 custom-scrollbar">
                    <table class="w-full text-left text-sm">
                        <thead class="bg-slate-50 text-slate-500 font-medium border-b border-slate-200 sticky top-0">
                            <tr>
                                <th class="px-6 py-3">Timestamp</th>
                                <th class="px-6 py-3">Event</th>
                                <th class="px-6 py-3">Username</th>
                                <th class="px-6 py-3">Operator</th>
                            </tr>
                        </thead>
                        <tbody id="raw-log-body" class="divide-y divide-slate-100 font-mono text-xs"></tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <script>
        // Data injected via PowerShell. 
        // Using Base64 prevents ANY special characters/backticks from breaking the Javascript Engine
        const base64CsvData = `{{CSV_BASE64}}`;
        const ignoredUsers = ['{{IGNORED_USERS}}'].map(u => u.toLowerCase());
        
        let rawGlobalLogData = [];
        let globalLogData = []; 
        let currentFilter = null; 
        let showAllYears = false;
        
        let trendChartInstance = null;
        let barChartInstance = null;
        
        let rawAdData = {{AD_LOCKOUTS_JSON}};
        if (!rawAdData) rawAdData = [];
        const adLockouts = Array.isArray(rawAdData) ? rawAdData : [rawAdData];

        let timeLeft = 120;
        const timerEl = document.getElementById('refresh-timer');
        setInterval(() => {
            if (timeLeft > 0) {
                timeLeft--;
                if(timerEl) timerEl.innerText = timeLeft;
            }
        }, 1000);

        function toggleYearView() {
            showAllYears = !showAllYears;
            const btn = document.getElementById('yearToggleBtn');
            const label = document.getElementById('kpi-total-label');
            
            if (showAllYears) {
                btn.innerHTML = '<i class="ph ph-calendar-blank"></i> Current Year Only';
                btn.classList.add('bg-blue-600/50', 'border-blue-400');
                label.innerText = 'TOTAL LOCKOUTS';
            } else {
                btn.innerHTML = '<i class="ph ph-calendar-plus"></i> Show All History';
                btn.classList.remove('bg-blue-600/50', 'border-blue-400');
                label.innerText = 'YTD LOCKOUTS';
            }
            
            resetFilters(); // Reset filters when changing view modes
            renderDashboard(false); // Re-render without re-parsing CSV
        }

        function parseCSVLine(line) {
            const regex = /(?:^|,)(\s*"(?:[^"]|"")*"|[^,]*)/g;
            const matches = [];
            let match;
            while (match = regex.exec(line)) {
                let val = match[1];
                if (val.trim().startsWith('"') && val.trim().endsWith('"')) { 
                    val = val.trim().slice(1, -1).replace(/""/g, '"'); 
                }
                matches.push(val.trim());
            }
            return matches;
        }

        function parseCSVData() {
            if(!base64CsvData) return [];
            
            let csvText = "";
            try {
                // Decode Base64 string to original UTF-8 text safely
                const binString = window.atob(base64CsvData);
                const bytes = new Uint8Array(binString.length);
                for (let i = 0; i < binString.length; i++) {
                    bytes[i] = binString.charCodeAt(i);
                }
                csvText = new TextDecoder('utf-8').decode(bytes);
            } catch(e) {
                console.error("Failed to decode base64 data", e);
                return [];
            }

            const lines = csvText.trim().split('\n');
            const results = [];
            
            for (let i = 1; i < lines.length; i++) {
                const line = lines[i].trim();
                if (!line) continue;
                const cols = parseCSVLine(line);
                if(cols.length < 3) continue; 
                
                // Exclude ignored users locally to save UI memory
                const username = (cols[2] || "").trim().toLowerCase();
                if (ignoredUsers.includes(username)) continue;

                const entry = {
                    timestamp: cols[0], event: cols[1], username: cols[2],
                    details: cols[3] || "", status: cols[4] || "", operator: cols[5] || ""
                };
                let d = new Date(entry.timestamp);
                if (isNaN(d)) { d = new Date(entry.timestamp.replace(' ', 'T')); }
                entry.dateObj = isNaN(d) ? new Date() : d;
                results.push(entry);
            }
            return results;
        }

        function processData(data) {
            const dailyCounts = {}; 
            const userCounts = {}; 
            let totalLockouts = 0;
            let manualUnlocks = 0;
            let autoUnlocks = 0;

            data.sort((a, b) => a.dateObj - b.dateObj);

            data.forEach(row => {
                // Extract local date parts to prevent UTC shifting
                const y = row.dateObj.getFullYear();
                const m = String(row.dateObj.getMonth() + 1).padStart(2, '0');
                const d = String(row.dateObj.getDate()).padStart(2, '0');
                const dateKey = `${y}-${m}-${d}`;
                
                const user = row.username;
                const evt = row.event.toLowerCase();

                if (evt.includes("locked out") || evt.includes("account locked") || evt.includes("lockout detected")) {
                    totalLockouts++;
                    dailyCounts[dateKey] = (dailyCounts[dateKey] || 0) + 1;
                    userCounts[user] = (userCounts[user] || 0) + 1;
                }
                else if (evt.includes("unlock") || evt.includes("lockout cleared")) {
                    if (evt.includes("auto") || evt.includes("cleared")) {
                        autoUnlocks++;
                    } else {
                        manualUnlocks++;
                    }
                }
            });
            return { dailyCounts, userCounts, totalLockouts, manualUnlocks, autoUnlocks, raw: data };
        }

        function filterLogs(type, value) {
            currentFilter = { type, value };
            const badge = document.getElementById('activeFilterBadge');
            const label = document.getElementById('filterLabel');
            badge.classList.remove('hidden');
            label.innerText = (type === 'date') ? `Date: ${value}` : `User: ${value}`;
            renderLogTable();
            document.getElementById('logSection').scrollIntoView({ behavior: 'smooth' });
        }

        function resetFilters() {
            currentFilter = null;
            document.getElementById('activeFilterBadge').classList.add('hidden');
            renderLogTable();
        }

        function renderLogTable() {
            const rawBody = document.getElementById('raw-log-body');
            rawBody.innerHTML = '';
            
            // 1. ALWAYS filter down to only lock/unlock events first
            let displayData = [...globalLogData].filter(log => {
                const evtLower = log.event.toLowerCase();
                return evtLower.includes('account locked') || 
                       evtLower.includes('locked out') || 
                       evtLower.includes('lockout detected') ||
                       evtLower.includes('unlock') ||
                       evtLower.includes('lockout cleared');
            }).reverse();

            // 2. Apply user/date filters if active
            if (currentFilter) {
                if (currentFilter.type === 'user') {
                    displayData = displayData.filter(log => log.username === currentFilter.value);
                } else if (currentFilter.type === 'date') {
                    displayData = displayData.filter(log => {
                        const y = log.dateObj.getFullYear();
                        const m = String(log.dateObj.getMonth() + 1).padStart(2, '0');
                        const d = String(log.dateObj.getDate()).padStart(2, '0');
                        return `${y}-${m}-${d}` === currentFilter.value;
                    });
                }
            }

            const limit = currentFilter ? 500 : 100;

            const htmlParts = displayData.slice(0, limit).map(log => {
                let color = "bg-slate-100 text-slate-600";
                const evtLower = log.event.toLowerCase();
                
                if (evtLower.includes('locked') || evtLower.includes('lockout detected')) color = "bg-red-50 text-red-600 font-bold";
                else if (evtLower.includes('unlock') || evtLower.includes('lockout cleared')) color = "bg-green-50 text-green-600";
                
                return `<tr><td class="px-6 py-2 whitespace-nowrap text-slate-500">${log.dateObj.toLocaleString()}</td><td class="px-6 py-2"><span class="${color} px-2 py-0.5 rounded text-[10px] uppercase border border-opacity-20 border-current">${log.event}</span></td><td class="px-6 py-2 font-bold">${log.username}</td><td class="px-6 py-2 text-slate-500">${log.operator}</td></tr>`;
            });
            rawBody.innerHTML = htmlParts.join('');
            
            if (displayData.length === 0) {
                rawBody.innerHTML = `<tr><td colspan="4" class="px-6 py-8 text-center text-slate-400">No records found for this filter.</td></tr>`;
            }
        }

        function renderDashboard(isInit = true) {
            // Only parse the giant base64 string once
            if (isInit) {
                rawGlobalLogData = parseCSVData();
            }
            
            // Apply Year Filter globally before processing
            const currentYear = new Date().getFullYear();
            globalLogData = showAllYears 
                ? rawGlobalLogData 
                : rawGlobalLogData.filter(log => log.dateObj.getFullYear() === currentYear);

            const stats = processData(globalLogData);

            const activeCount = adLockouts ? adLockouts.length : 0;
            document.getElementById('kpi-current-locked').innerText = activeCount;
            document.getElementById('kpi-total-events').innerText = stats.totalLockouts;
            document.getElementById('kpi-manual-unlocks').innerText = stats.manualUnlocks;
            document.getElementById('kpi-auto-unlocks').innerText = stats.autoUnlocks;
            
            const totalMinutes = stats.manualUnlocks * 5;
            let timeString = totalMinutes + "m";
            if (totalMinutes > 60) {
               const hrs = Math.floor(totalMinutes / 60);
               const mins = totalMinutes % 60;
               timeString = hrs + "h " + mins + "m";
            }
            document.getElementById('kpi-agent-time').innerText = timeString;

            const activeTableBody = document.getElementById('active-lockouts-body');
            activeTableBody.innerHTML = '';
            if (activeCount === 0) {
                activeTableBody.innerHTML = `<tr><td colspan="5" class="px-6 py-8 text-center text-slate-400 bg-green-50">No active lockouts in Active Directory.</td></tr>`;
            } else {
                adLockouts.forEach(item => {
                    const lockoutTime = item.LockoutTime || "Unknown";
                    
                    const userLogs = globalLogData.filter(l => {
                        const evt = l.event.toLowerCase();
                        return l.username.toLowerCase() === item.SamAccountName.toLowerCase() && 
                        (evt.includes('locked out') || evt.includes('account locked') || evt.includes('lockout detected'));
                    });
                    userLogs.sort((a,b) => b.dateObj - a.dateObj);
                    
                    let lockoutSource = "Log Not Found";
                    if(userLogs.length > 0) {
                        lockoutSource = userLogs[0].details || "N/A";
                    }

                    activeTableBody.innerHTML += `<tr class="hover:bg-slate-50">
                        <td class="px-6 py-4 font-bold text-slate-800">${item.SamAccountName}</td>
                        <td class="px-6 py-4 text-slate-600">${lockoutTime}</td>
                        <td class="px-6 py-4">
                            <span class="font-mono text-xs bg-slate-100 border border-slate-200 px-2 py-1 rounded text-slate-600">
                                ${lockoutSource}
                            </span>
                        </td>
                        <td class="px-6 py-4 text-xs text-slate-400 italic">Real-time AD Status</td>
                        <td class="px-6 py-4"><span class="bg-red-100 text-red-700 px-2 py-1 rounded text-xs font-bold">LOCKED</span></td>
                    </tr>`;
                });
            }

            renderLogTable();

            const trendCtx = document.getElementById('lockoutTrendChart');
            if(trendCtx) {
                if (trendChartInstance) trendChartInstance.destroy();
                trendChartInstance = new Chart(trendCtx, {
                    type: 'line',
                    data: {
                        labels: Object.keys(stats.dailyCounts).sort(),
                        datasets: [{ label: 'Lockouts', data: Object.keys(stats.dailyCounts).sort().map(d => stats.dailyCounts[d]), borderColor: '#005596', backgroundColor: 'rgba(0, 85, 150, 0.1)', fill: true, tension: 0.3 }]
                    },
                    options: { 
                        responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } },
                        scales: { x: { ticks: { maxTicksLimit: 15 } } },
                        elements: { point: { radius: Object.keys(stats.dailyCounts).length > 30 ? 2 : 4 } },
                        onClick: (evt, activeElements, chart) => {
                            if (activeElements.length > 0) {
                                filterLogs('date', chart.data.labels[activeElements[0].index]);
                            }
                        },
                        onHover: (event, chartElement) => {
                            event.native.target.style.cursor = chartElement[0] ? 'pointer' : 'default';
                        }
                    }
                });
            }

            const barCtx = document.getElementById('userBarChart');
            if(barCtx) {
                if (barChartInstance) barChartInstance.destroy();
                barChartInstance = new Chart(barCtx, {
                    type: 'bar',
                    data: {
                        labels: Object.keys(stats.userCounts).sort((a,b) => stats.userCounts[b] - stats.userCounts[a]).slice(0, 10),
                        datasets: [{ label: 'Lockouts', data: Object.values(stats.userCounts).sort((a,b) => b - a).slice(0, 10), backgroundColor: '#00A9CE', borderRadius: 4 }]
                    },
                    options: { 
                        responsive: true, maintainAspectRatio: false, indexAxis: 'y', plugins: { legend: { display: false } },
                        onClick: (evt, activeElements, chart) => {
                            if (activeElements.length > 0) {
                                filterLogs('user', chart.data.labels[activeElements[0].index]);
                            }
                        },
                        onHover: (event, chartElement) => {
                            event.native.target.style.cursor = chartElement[0] ? 'pointer' : 'default';
                        }
                    }
                });
            }
        }
        
        try {
            renderDashboard(true);
        } catch(e) {
            console.error("Dashboard Rendering Error:", e);
            document.body.innerHTML += `<div class='fixed bottom-0 w-full bg-red-600 text-white p-2 text-center'>Error rendering dashboard: ${e.message}</div>`;
        }
    </script>
</body>
</html>
'@

# --- UPDATE FUNCTION ---
function Update-Dashboard {
    try {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Change Detected. Regenerating..." -NoNewline

        $Base64CsvString = ""
        $DateRangeString = "No logs found"
        $AllCsvLines = New-Object System.Collections.Generic.List[string]
        $IsFirstFile = $true

        if (Test-Path -LiteralPath $LogPath) {
            # Bypassing PowerShell Path Bug while maintaining performance:
            # When a LiteralPath contains brackets (e.g., [toolkit]), the -Filter parameter breaks and returns nothing.
            # Using an escaped path with -Path allows -Filter to work, drastically improving performance over Where-Object.
            $EscapedPath = [System.Management.Automation.WildcardPattern]::Escape($LogPath)
            $CsvFiles = Get-ChildItem -Path $EscapedPath -Filter "UnlockLog_*.csv" -Recurse -ErrorAction SilentlyContinue |
                        Sort-Object LastWriteTime
            
            if ($CsvFiles) {
                $EarliestFile = $CsvFiles | Select-Object -First 1
                $EarliestDate = $EarliestFile.LastWriteTime.ToString("MM/dd/yyyy")
                if ($EarliestFile.CreationTime -lt $EarliestFile.LastWriteTime) {
                    $EarliestDate = $EarliestFile.CreationTime.ToString("MM/dd/yyyy")
                }
                $DateRangeString = "Since $EarliestDate"

                # Read files robustly with retries
                foreach ($file in $CsvFiles) {
                    $retry = 3
                    $success = $false
                    
                    while ($retry -gt 0 -and -not $success) {
                        try {
                            $stream = [System.IO.File]::Open($file.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                            $reader = New-Object System.IO.StreamReader($stream)
                            
                            $lineCount = 0
                            while ($null -ne ($line = $reader.ReadLine())) {
                                if (-not [string]::IsNullOrWhiteSpace($line)) {
                                    if ($lineCount -eq 0) {
                                        # Only add the header line for the very first file
                                        if ($IsFirstFile) { $AllCsvLines.Add($line) }
                                    } else {
                                        $AllCsvLines.Add($line)
                                    }
                                }
                                $lineCount++
                            }
                            
                            $reader.Close()
                            $stream.Close()
                            $IsFirstFile = $false
                            $success = $true
                        } catch {
                            $retry--
                            if ($retry -gt 0) { Start-Sleep -Milliseconds 500 }
                        }
                    }
                    
                    if (-not $success) { Write-Warning "Skipped file due to persistent lock: $($file.Name)" }
                }

                # Bypass PowerShell's pipeline limits by using raw string handling and Base64
                if ($AllCsvLines.Count -gt 1) {
                    $FullCsvText = $AllCsvLines -join "`n"
                    $Bytes = [System.Text.Encoding]::UTF8.GetBytes($FullCsvText)
                    $Base64CsvString = [Convert]::ToBase64String($Bytes)
                }
            }
        }

        # Query AD
        $AdLockoutsJson = "[]"
        try {
            # Build server-side filter to exclude ignored users at the AD query level
            $Filter = "LockedOut -eq `$true"
            if ($IgnoredUsers) {
                foreach ($user in $IgnoredUsers) {
                    $Filter += " -and SamAccountName -notlike '$user'"
                }
            }

            $lockedAccounts = Get-ADUser -Filter $Filter -Properties LockoutTime -ErrorAction Stop |
                              Select-Object SamAccountName, Name, @{N='LockoutTime';E={
                                  if ($_.LockoutTime) { [DateTime]::FromFileTime($_.LockoutTime).ToString("MM/dd/yyyy HH:mm:ss") } else { "" }
                              }}
            if ($lockedAccounts) {
                $AdLockoutsJson = ConvertTo-Json -InputObject @($lockedAccounts) -Compress
            }
        } catch {}

        # Inject and Save
        $CurrentHtml = $HtmlTemplate.Replace("{{CSV_BASE64}}", $Base64CsvString)
        $CurrentHtml = $CurrentHtml.Replace("{{AD_LOCKOUTS_JSON}}", $AdLockoutsJson)
        $CurrentHtml = $CurrentHtml.Replace("{{IGNORED_USERS}}", ($IgnoredUsers -join "','"))
        $CurrentHtml = $CurrentHtml.Replace("{{GENERATED_TIME}}", (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
        $CurrentHtml = $CurrentHtml.Replace("{{DATE_RANGE}}", $DateRangeString)

        $ParentDir = Split-Path -Parent $OutputHtml
        if (-not (Test-Path -LiteralPath $ParentDir)) { New-Item -ItemType Directory -LiteralPath $ParentDir -Force | Out-Null }
        
        $CurrentHtml | Out-File -LiteralPath $OutputHtml -Encoding UTF8
        Write-Host " DONE." -ForegroundColor Green
    } catch { Write-Error "`n[Error] $_" }
}

# --- POLLING MANAGEMENT ---
$Global:LastSignature = ""

function Check-For-Updates {
    if (Test-Path -LiteralPath $LogPath) {
        # Bypassing PowerShell Path Bug here as well using escaped path and -Filter for better performance
        $EscapedPath = [System.Management.Automation.WildcardPattern]::Escape($LogPath)
        $CurrentSignature = Get-ChildItem -Path $EscapedPath -Filter "UnlockLog_*.csv" -Recurse -ErrorAction SilentlyContinue |
                            Sort-Object Name |
                            ForEach-Object { "$($_.Name)|$($_.LastWriteTime.Ticks)|$($_.Length)" } |
                            Out-String
        
        if ($CurrentSignature -ne $Global:LastSignature) {
            $Global:LastSignature = $CurrentSignature
            return $true
        }
    }
    return $false
}

# --- EXECUTION LOOP ---
Write-Host "Monitoring started on: $LogPath" -ForegroundColor Cyan
Write-Host "Output HTML: $OutputHtml" -ForegroundColor Gray
Write-Host "Mode: Smart Polling (5s interval)" -ForegroundColor Gray
Write-Host "Press Ctrl+C to stop." -ForegroundColor Yellow

if (Check-For-Updates) { Update-Dashboard } else { Update-Dashboard }

while ($true) {
    if (Check-For-Updates) { Update-Dashboard }
    Start-Sleep -Seconds 5
}