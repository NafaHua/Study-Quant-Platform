<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>数学量化 · Self-Quantified Dashboard</title>

    <script src="https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js">
    </script>
    <script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js">
    </script>

    <style>
        :root {
            --bg-body: #0b0f1c;
            --bg-sidebar: #0f1526;
            --bg-card: #111827;
            --bg-card-hover: #161e30;
            --border: rgba(255, 255, 255, 0.07);
            --border-active: rgba(99, 130, 255, 0.35);
            --text: #e2e8f0;
            --text-secondary: #94a3b8;
            --text-muted: #64748b;
            --accent: #6d8aff;
            --accent-glow: rgba(109, 138, 255, 0.3);
            --success: #34d399;
            --warning: #fbbf24;
            --danger: #f87171;
            --purple: #a78bfa;
            --radius: 14px;
            --transition: 0.2s ease;
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        html {
            font-size: 15px;
            height: 100%;
        }

        body {
            font-family: 'Inter', 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif;
            background: var(--bg-body);
            color: var(--text);

            /* ★ 核心：CSS Grid 横版布局 */
            display: grid;
            grid-template-columns: 240px 1fr;
            grid-template-rows: 1fr;
            height: 100vh;
            height: 100dvh;
            overflow: hidden;

            -webkit-font-smoothing: antialiased;
        }

        /* ========== 侧边栏 ========== */
        .sidebar {
            background: var(--bg-sidebar);
            border-right: 1px solid var(--border);
            display: flex;
            flex-direction: column;
            padding: 1.6rem 0;
            overflow-y: auto;
            overflow-x: hidden;
            z-index: 10;
        }
        .sidebar-logo {
            padding: 0 1.4rem 1.2rem;
            margin-bottom: 0.6rem;
            border-bottom: 1px solid var(--border);
            display: flex;
            align-items: center;
            gap: 0.6rem;
        }
        .sidebar-logo .icon {
            width: 36px;
            height: 36px;
            border-radius: 10px;
            background: linear-gradient(135deg, var(--accent), var(--purple));
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.1rem;
            flex-shrink: 0;
            box-shadow: 0 0 18px var(--accent-glow);
        }
        .sidebar-logo .title {
            font-weight: 700;
            font-size: 1rem;
            line-height: 1.2;
        }
        .sidebar-logo .subtitle {
            font-size: 0.68rem;
            color: var(--text-muted);
            font-weight: 500;
        }

        .nav-item {
            padding: 0.6rem 1.4rem;
            margin: 2px 0.6rem;
            cursor: pointer;
            border-radius: 8px;
            transition: var(--transition);
            color: var(--text-secondary);
            display: flex;
            align-items: center;
            gap: 0.5rem;
            font-weight: 500;
            font-size: 0.85rem;
            white-space: nowrap;
            user-select: none;
        }
        .nav-item:hover {
            background: rgba(255, 255, 255, 0.04);
            color: var(--text);
        }
        .nav-item.active {
            background: rgba(109, 138, 255, 0.12);
            color: var(--accent);
            font-weight: 600;
            box-shadow: inset 0 0 0 1px rgba(109, 138, 255, 0.15);
        }

        /* ========== 主区域 ========== */
        .main {
            display: flex;
            flex-direction: column;
            overflow-y: auto;
            overflow-x: hidden;
            padding: 1.3rem 1.6rem;
            gap: 1.1rem;
        }
        .main::-webkit-scrollbar {
            width: 5px;
        }
        .main::-webkit-scrollbar-thumb {
            background: rgba(255, 255, 255, 0.06);
            border-radius: 10px;
        }

        /* 顶栏 */
        .top-strip {
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 0.6rem;
            padding: 0.5rem 1rem;
            background: var(--bg-card);
            border-radius: var(--radius);
            border: 1px solid var(--border);
            flex-shrink: 0;
        }
        .pill {
            display: inline-flex;
            align-items: center;
            gap: 0.3rem;
            padding: 0.3rem 0.75rem;
            border-radius: 20px;
            font-size: 0.76rem;
            font-weight: 600;
            background: rgba(255, 255, 255, 0.04);
            color: var(--text-secondary);
            border: 1px solid var(--border);
        }
        .pill.accent {
            background: rgba(109, 138, 255, 0.12);
            color: var(--accent);
            border-color: rgba(109, 138, 255, 0.2);
        }
        .pill.warn {
            background: rgba(251, 191, 36, 0.1);
            color: var(--warning);
            border-color: rgba(251, 191, 36, 0.2);
        }

        /* 视图 */
        .view {
            display: none;
            flex-direction: column;
            gap: 1.1rem;
        }
        .view.active {
            display: flex;
        }

        /* 卡片 */
        .card {
            background: var(--bg-card);
            border-radius: var(--radius);
            padding: 1.3rem;
            border: 1px solid var(--border);
            transition: var(--transition);
        }
        .card:hover {
            border-color: var(--border-active);
            background: var(--bg-card-hover);
        }
        .card-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 0.9rem;
            flex-wrap: wrap;
            gap: 0.5rem;
        }
        .card-header h3 {
            font-size: 0.9rem;
            font-weight: 650;
        }

        /* 统计卡片网格 */
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(148px, 1fr));
            gap: 0.7rem;
        }
        .stat-card {
            background: var(--bg-card);
            border-radius: var(--radius);
            padding: 0.9rem 1rem;
            border: 1px solid var(--border);
            text-align: center;
            transition: var(--transition);
        }
        .stat-card:hover {
            border-color: var(--border-active);
            transform: translateY(-2px);
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
        }
        .stat-value {
            font-size: 1.5rem;
            font-weight: 750;
            background: linear-gradient(135deg, var(--accent), #22d3ee);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            line-height: 1.2;
        }
        .stat-label {
            font-size: 0.7rem;
            color: var(--text-muted);
            margin-top: 0.15rem;
            font-weight: 500;
        }

        /* 双列 */
        .grid-2 {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(360px, 1fr));
            gap: 1.1rem;
        }
        .chart-container {
            width: 100%;
            height: 260px;
        }

        /* 按钮 */
        .btn {
            border: 1px solid var(--border);
            padding: 0.45rem 0.9rem;
            border-radius: 8px;
            font-weight: 600;
            cursor: pointer;
            font-size: 0.78rem;
            transition: var(--transition);
            white-space: nowrap;
            display: inline-flex;
            align-items: center;
            gap: 0.3rem;
            background: rgba(255, 255, 255, 0.03);
            color: var(--text-secondary);
            user-select: none;
        }
        .btn:hover {
            background: rgba(255, 255, 255, 0.07);
            color: var(--text);
            border-color: rgba(255, 255, 255, 0.15);
        }
        .btn:active {
            transform: scale(0.97);
        }
        .btn-primary {
            background: var(--accent);
            color: #fff;
            border: none;
            box-shadow: 0 4px 14px var(--accent-glow);
        }
        .btn-primary:hover {
            background: #7d97ff;
            color: #fff;
            box-shadow: 0 6px 20px rgba(109, 138, 255, 0.45);
        }
        .btn-danger {
            background: rgba(248, 113, 113, 0.12);
            color: var(--danger);
            border-color: rgba(248, 113, 113, 0.2);
        }
        .btn-danger:hover {
            background: rgba(248, 113, 113, 0.2);
        }
        .btn-sm {
            padding: 0.3rem 0.65rem;
            font-size: 0.68rem;
            border-radius: 6px;
        }

        /* 表单 */
        .form-row {
            display: flex;
            gap: 0.45rem;
            flex-wrap: wrap;
            align-items: flex-end;
        }
        .form-group {
            display: flex;
            flex-direction: column;
            min-width: 80px;
        }
        .form-group label {
            font-size: 0.65rem;
            margin-bottom: 0.15rem;
            color: var(--text-muted);
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.03em;
        }
        .form-group input,
        .form-group textarea {
            padding: 0.45rem 0.6rem;
            border: 1.5px solid var(--border);
            border-radius: 8px;
            font-size: 0.8rem;
            background: rgba(255, 255, 255, 0.03);
            color: var(--text);
            transition: var(--transition);
            font-family: inherit;
            outline: none;
        }
        .form-group input:focus,
        .form-group textarea:focus {
            border-color: var(--accent);
            box-shadow: 0 0 0 3px rgba(109, 138, 255, 0.08);
        }
        .form-group textarea {
            resize: vertical;
            min-height: 55px;
        }

        /* 上传区 */
        .upload-area {
            border: 2px dashed var(--border);
            border-radius: var(--radius);
            padding: 1.2rem;
            text-align: center;
            cursor: pointer;
            transition: var(--transition);
            color: var(--text-muted);
            font-size: 0.8rem;
        }
        .upload-area:hover {
            border-color: var(--accent);
            background: rgba(109, 138, 255, 0.04);
            color: var(--accent);
        }

        /* 表格 */
        .table-wrapper {
            max-height: 360px;
            overflow: auto;
            border-radius: 10px;
            border: 1px solid var(--border);
        }
        .table-wrapper::-webkit-scrollbar {
            height: 4px;
            width: 4px;
        }
        .table-wrapper::-webkit-scrollbar-thumb {
            background: rgba(255, 255, 255, 0.05);
            border-radius: 10px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.73rem;
            white-space: nowrap;
        }
        th,
        td {
            padding: 0.5rem 0.65rem;
            border-bottom: 1px solid var(--border);
            text-align: center;
        }
        th {
            background: rgba(255, 255, 255, 0.03);
            position: sticky;
            top: 0;
            z-index: 2;
            font-weight: 650;
            color: var(--text-secondary);
            font-size: 0.68rem;
            text-transform: uppercase;
            letter-spacing: 0.02em;
        }
        tr:hover td {
            background: rgba(255, 255, 255, 0.015);
        }

        /* 徽章 */
        .alert-badge {
            padding: 0.2rem 0.65rem;
            border-radius: 20px;
            font-size: 0.68rem;
            font-weight: 600;
            background: rgba(248, 113, 113, 0.12);
            color: var(--danger);
            display: inline-block;
            margin: 0.08rem;
        }
        .anomaly-badge {
            background: rgba(248, 113, 113, 0.18);
            color: #fca5a5;
            padding: 0.12rem 0.45rem;
            border-radius: 20px;
            font-weight: 700;
            font-size: 0.62rem;
            white-space: nowrap;
        }
        .info-row {
            padding: 0.45rem 0.75rem;
            background: rgba(109, 138, 255, 0.05);
            border-radius: 8px;
            font-size: 0.76rem;
            color: var(--text-secondary);
            margin-bottom: 0.5rem;
            border: 1px solid rgba(109, 138, 255, 0.08);
        }
        .info-row strong {
            color: var(--accent);
        }

        /* 折叠 */
        .collapse-header {
            cursor: pointer;
            padding: 0.3rem 0;
            display: flex;
            justify-content: space-between;
            align-items: center;
            user-select: none;
            color: var(--text-secondary);
            font-size: 0.82rem;
        }
        .collapse-header:hover {
            color: var(--text);
        }

        /* 记录 */
        .record-entry {
            background: rgba(255, 255, 255, 0.015);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 0.75rem;
            margin-bottom: 0.5rem;
            font-size: 0.78rem;
        }
        .record-entry img {
            max-width: 100%;
            max-height: 180px;
            margin: 0.35rem 0;
            border-radius: 8px;
            object-fit: contain;
        }

        /* 预测表格 */
        .prediction-table th {
            background: rgba(109, 138, 255, 0.18);
            color: var(--accent);
        }
        .prediction-table td {
            font-weight: 600;
            font-size: 0.78rem;
        }

        /* ========== 响应式：仅手机竖屏触发 ========== */
        @media (max-width: 640px) {
            body {
                grid-template-columns: 1fr;
                grid-template-rows: auto 1fr;
            }
            .sidebar {
                flex-direction: row;
                overflow-x: auto;
                padding: 0.4rem 0.5rem;
                border-right: none;
                border-bottom: 1px solid var(--border);
                gap: 0.15rem;
            }
            .sidebar-logo {
                display: none;
            }
            .nav-item {
                padding: 0.4rem 0.6rem;
                margin: 0;
                border-radius: 6px;
                font-size: 0.72rem;
                flex-shrink: 0;
                white-space: nowrap;
            }
            .main {
                padding: 0.7rem;
                gap: 0.7rem;
            }
            .stats-grid {
                grid-template-columns: repeat(2, 1fr);
                gap: 0.45rem;
            }
            .stat-value {
                font-size: 1.2rem;
            }
            .grid-2 {
                grid-template-columns: 1fr;
                gap: 0.7rem;
            }
            .chart-container {
                height: 200px;
            }
            .card {
                padding: 0.9rem;
            }
            .form-row {
                flex-direction: column;
            }
            .form-group {
                min-width: 100%;
            }
        }

        @media (min-width: 1400px) {
            .stats-grid {
                grid-template-columns: repeat(7, 1fr);
            }
            .chart-container {
                height: 300px;
            }
        }
    </style>

    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
</head>
<body>

    <!-- ========== 侧边栏 ========== -->
    <nav class="sidebar">
        <div class="sidebar-logo">
            <div class="icon">📐</div>
            <div>
                <div class="title">数学量化</div>
                <div class="subtitle">Self-Quantified</div>
            </div>
        </div>
        <div class="nav-item active" data-view="data">📁 数据管理</div>
        <div class="nav-item" data-view="dashboard">📈 量化仪表盘</div>
        <div class="nav-item" data-view="predict">🔮 预测中心</div>
        <div class="nav-item" data-view="records">📓 学习记录</div>
        <div class="nav-item" data-view="motivation">🏆 激励中心</div>
    </nav>

    <!-- ========== 主区域 ========== -->
    <main class="main">
        <div class="top-strip">
            <span class="pill accent">🔥 连续 <strong id="streakDays">0</strong> 天</span>
            <span class="pill">⭐ 经验 <strong id="xpValue">0</strong></span>
            <span class="pill warn" id="latestBadge"></span>
        </div>

        <!-- 数据管理 -->
        <section class="view active" id="view-data">
            <div class="card-header">
                <h3>📁 数据管理</h3>
                <div style="display:flex;gap:0.35rem;">
                    <button class="btn btn-primary" id="saveSampleBtn">📥 加载150天示例</button>
                    <button class="btn btn-danger" id="deleteAllBtn">🗑️ 清空</button>
                </div>
            </div>
            <div class="card">
                <div class="card-header"><h3>📝 实时录入</h3></div>
                <div class="form-row">
                    <div class="form-group"><label>学科</label><input type="text" id="inputSubject" value="数学" readonly></div>
                    <div class="form-group"><label>日期</label><input type="date" id="inputDate"></div>
                    <div class="form-group"><label>时长(分)</label><input type="number" id="inputDuration" value="60"></div>
                    <div class="form-group"><label>练习%</label><input type="number" id="inputPractice" value="75"></div>
                    <div class="form-group"><label>测试%</label><input type="number" id="inputTest" value="70"></div>
                    <div class="form-group"><label>复习%</label><input type="number" id="inputReview" value="65"></div>
                    <div class="form-group"><label>记忆%</label><input type="number" id="inputMemory" value="68"></div>
                    <div class="form-group"><label>权重</label><input type="number" id="inputWeight" value="20"></div>
                    <button class="btn btn-primary" id="addRecordBtn">✅ 添加</button>
                </div>
                <div class="upload-area" id="uploadArea" style="margin-top:0.5rem;">
                    📂 点击上传 Excel 文件（.xlsx / .xls / .csv）
                    <input type="file" id="fileInput" accept=".xlsx,.xls,.csv" hidden>
                </div>
                <div class="table-wrapper"><table id="dataTable"><thead></thead><tbody></tbody></table></div>
            </div>
        </section>

        <!-- 仪表盘 -->
        <section class="view" id="view-dashboard">
            <div class="card-header"><h3>📈 学习量化仪表盘</h3></div>
            <div class="stats-grid" id="statsCards"></div>
            <div class="grid-2">
                <div class="card"><div class="card-header"><h3>📊 学习时长 & 测试趋势</h3></div><div class="chart-container" id="trendChart"></div></div>
                <div class="card"><div class="card-header"><h3>📈 测试分数分布</h3></div><div class="chart-container" id="distChart"></div></div>
            </div>
            <div class="grid-2">
                <div class="card"><div class="card-header"><h3>🎯 各维度雷达图</h3></div><div class="chart-container" id="radarChart"></div></div>
                <div class="card"><div class="card-header"><h3>🔍 练习 vs 测试相关性</h3></div><div class="chart-container" id="scatterChart"></div></div>
            </div>
            <div class="card" id="alertContainer"></div>
        </section>

        <!-- 预测中心 -->
        <section class="view" id="view-predict">
            <div class="card-header"><h3>🔮 预测中心 · 数学</h3></div>
            <div class="card">
                <div class="collapse-header" id="settingsToggle">
                    <strong>⚙️ 预测参数设置（S型提升 / 遗忘底线）</strong>
                    <span id="settingsArrow">▼</span>
                </div>
                <div id="settingsBody">
                    <div class="form-row" style="margin-top:0.5rem;">
                        <div class="form-group"><label>目标正确率%</label><input type="number" id="targetScore" value="90" min="30" max="100"></div>
                        <div class="form-group"><label>S型中点(天)</label><input type="number" id="sMid" value="15" min="1"></div>
                        <div class="form-group"><label>陡峭度 k</label><input type="number" id="sK" value="0.3" step="0.05"></div>
                        <div class="form-group"><label>遗忘半衰期</label><input type="number" id="halfLife" value="12" min="1"></div>
                        <div class="form-group"><label>底线分%</label><input type="number" id="minScore" value="40" min="0" max="80"></div>
                        <button class="btn btn-primary btn-sm" id="savePredictSettingsBtn">💾 保存</button>
                        <span id="settingsSavedHint" style="color:var(--success);display:none;">✅ 已保存</span>
                    </div>
                </div>
            </div>
            <div class="card">
                <div class="card-header"><h3>📜 测试走势 <span class="anomaly-badge">红标=异常</span></h3></div>
                <div class="chart-container" id="historyChart"></div>
                <div class="table-wrapper" style="margin-top:0.5rem;"><table id="historyTable"></table></div>
            </div>
            <div class="card">
                <div class="card-header"><h3>⚠️ 异常点记录</h3></div>
                <div class="table-wrapper"><table id="anomalyTable"><thead><tr><th>日期</th><th>测试%</th><th>偏离 σ</th></tr></thead><tbody><tr><td colspan="3">暂无异常</td></tr></tbody></table></div>
            </div>
            <div class="card">
                <div class="card-header"><h3>🔮 未来30天预测</h3></div>
                <div class="info-row" id="forecastInfo"></div>
                <div class="table-wrapper"><table class="prediction-table" id="forecastTable"></table></div>
                <div class="chart-container" id="forecastChart" style="margin-top:0.5rem;"></div>
            </div>
        </section>

        <!-- 学习记录 -->
        <section class="view" id="view-records">
            <div class="card-header"><h3>📓 学习情况记录</h3></div>
            <div class="card">
                <div class="card-header"><h3>📝 添加新记录</h3></div>
                <div class="form-row">
                    <div class="form-group"><label>日期</label><input type="date" id="recordDate"></div>
                    <div class="form-group" style="flex:1;min-width:180px;"><label>文字记录</label><textarea id="recordText" rows="2" placeholder="今天学了什么…"></textarea></div>
                    <div class="form-group"><label>上传图片</label><input type="file" id="recordImage" accept="image/*"></div>
                    <button class="btn btn-primary" id="addRecordNoteBtn">➕ 添加</button>
                </div>
            </div>
            <div class="card">
                <div class="card-header"><h3>📋 历史记录</h3></div>
                <div id="recordNotesContainer"></div>
            </div>
        </section>

        <!-- 激励中心 -->
        <section class="view" id="view-motivation">
            <div class="card">
                <div class="card-header"><h3>🏅 成就徽章</h3></div>
                <div id="achievementList" class="stats-grid"></div>
            </div>
            <div class="card">
                <div class="card-header"><h3>📜 惩罚记录</h3></div>
                <div id="penaltyLog" style="font-size:0.78rem;color:var(--text-secondary);"></div>
            </div>
        </section>
    </main>

    <script>
        (function() {
            const STORAGE_DATA = 'sq_math_data_v5';
            const STORAGE_MOTIV = 'sq_math_motiv_v5';
            const STORAGE_SETTINGS = 'sq_math_settings_v5';
            const STORAGE_NOTES = 'sq_math_notes_v5';

            let records = [];
            let motivation = { xp: 0, streak: 0, achievements: {}, penaltyLog: [], lastCalcDate: '',
                lastPenaltyDate: '' };
            let forecastSettings = { target: 90, sMid: 15, sK: 0.3, halfLife: 12, minScore: 40 };
            let studyNotes = [];
            const chartInstances = {};

            const darkTheme = {
                textStyle: { color: '#94a3b8' },
                tooltip: { backgroundColor: 'rgba(17,24,39,0.96)', borderColor: 'rgba(109,138,255,0.25)',
                    textStyle: { color: '#e2e8f0' } },
            };

            function disposeChart(k) { if (chartInstances[k]) { chartInstances[k].dispose(); delete chartInstances[
                k]; } }

            function save(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} }

            function load(k) { try { const s = localStorage.getItem(k); return s ? JSON.parse(s) : null; } catch (
                e) { return null; } }

            function loadAll() {
                records = load(STORAGE_DATA) || [];
                motivation = load(STORAGE_MOTIV) || { xp: 0, streak: 0, achievements: {}, penaltyLog: [],
                    lastCalcDate: '', lastPenaltyDate: '' };
                forecastSettings = load(STORAGE_SETTINGS) || { target: 90, sMid: 15, sK: 0.3, halfLife: 12,
                    minScore: 40 };
                studyNotes = load(STORAGE_NOTES) || [];
            }

            function clamp(v, min, max) { return Math.max(min, Math.min(max, v || 0)); }

            // 导航
            document.querySelectorAll('.nav-item').forEach(item => {
                item.addEventListener('click', () => {
                    const vn = item.dataset.view;
                    document.querySelectorAll('.nav-item').forEach(n => n.classList.remove(
                        'active'));
                    item.classList.add('active');
                    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
                    const el = document.getElementById(`view-${vn}`);
                    if (el) el.classList.add('active');
                    if (vn === 'data') renderDataTable();
                    if (vn === 'dashboard') renderDashboard();
                    if (vn === 'predict') { renderHistory();
                        renderForecast();
                        renderAnomalies(); }
                    if (vn === 'records') renderNotes();
                    if (vn === 'motivation') renderMotivation();
                });
            });

            // 数据录入
            document.getElementById('inputDate').valueAsDate = new Date();
            document.getElementById('addRecordBtn').addEventListener('click', () => {
                records.unshift({
                    subject: '数学',
                    date: document.getElementById('inputDate').value,
                    duration: +document.getElementById('inputDuration').value,
                    practice: clamp(+document.getElementById('inputPractice').value, 0, 100),
                    test: clamp(+document.getElementById('inputTest').value, 0, 100),
                    review: clamp(+document.getElementById('inputReview').value, 0, 100),
                    memory: clamp(+document.getElementById('inputMemory').value, 0, 100),
                    weight: +document.getElementById('inputWeight').value || 20
                });
                save(STORAGE_DATA, records);
                updateAll();
            });
            window.deleteRecord = (idx) => { if (confirm('确认删除？')) { records.splice(idx, 1);
                    save(STORAGE_DATA, records);
                    updateAll(); } };
            document.getElementById('deleteAllBtn').addEventListener('click', () => { if (confirm(
                    '⚠️ 清空全部数据？')) { records = [];
                    save(STORAGE_DATA, records);
                    updateAll(); } });

            // Excel 导入
            const ua = document.getElementById('uploadArea'),
                fi = document.getElementById('fileInput');
            ua.addEventListener('click', () => fi.click());
            ua.addEventListener('dragover', e => { e.preventDefault();
                ua.style.borderColor = 'var(--accent)'; });
            ua.addEventListener('dragleave', () => { ua.style.borderColor = 'var(--border)'; });
            ua.addEventListener('drop', e => {
                e.preventDefault();
                ua.style.borderColor = 'var(--border)';
                if (e.dataTransfer.files[0]) processExcel(e.dataTransfer.files[0]);
            });
            fi.addEventListener('change', e => { if (e.target.files[0]) processExcel(e.target.files[0]); });

            function processExcel(file) {
                const reader = new FileReader();
                reader.onload = e => {
                    try {
                        const wb = XLSX.read(e.target.result, { type: 'binary' });
                        const ws = wb.Sheets[wb.SheetNames[0]];
                        const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
                        if (json.length < 2) return alert('需要标题行');
                        const headers = json[0];
                        const map = { '学科': 'subject', '日期': 'date', '学习时长(分钟)': 'duration',
                            '练习正确率%': 'practice', '测试正确率%': 'test', '复习完成率%': 'review',
                            '知识点记忆%': 'memory', '权重': 'weight' };
                        const parsed = json.slice(1).filter(r => r.some(c => c != null)).map(r => {
                            const obj = {};
                            headers.forEach((h, i) => { if (map[h]) obj[map[h]] = r[i]; });
                            obj.practice = clamp(+obj.practice, 0, 100);
                            obj.test = clamp(+obj.test, 0, 100);
                            obj.review = clamp(+obj.review, 0, 100);
                            obj.memory = clamp(+obj.memory, 0, 100);
                            obj.duration = +obj.duration || 0;
                            obj.weight = +obj.weight || 20;
                            obj.subject = obj.subject || '数学';
                            return obj;
                        });
                        records = [...records, ...parsed];
                        save(STORAGE_DATA, records);
                        updateAll();
                        alert(`✅ 导入 ${parsed.length} 条`);
                    } catch (ex) { alert('❌ 解析失败'); }
                };
                reader.readAsBinaryString(file);
            }

            // 150天示例
            function loadMathSample() {
                if (records.length && !confirm('替换为150天数据？')) return;
                const start = new Date();
                start.setDate(start.getDate() - 149);
                const sample = [];
                for (let i = 0; i < 150; i++) {
                    const d = new Date(start);
                    d.setDate(d.getDate() + i);
                    const p = i / 149;
                    const base = 45 + p * 40 + Math.sin(p * Math.PI * 10) * 15;
                    const t = clamp(Math.round(base + (Math.random() - 0.5) * 25), 0, 100);
                    sample.push({
                        subject: '数学',
                        date: d.toISOString().split('T')[0],
                        duration: 45 + Math.round(Math.random() * 40),
                        practice: clamp(t + Math.round((Math.random() - 0.5) * 12), 0, 100),
                        test: t,
                        review: clamp(t - 5 + Math.round((Math.random() - 0.5) * 15), 0, 100),
                        memory: clamp(t - 3 + Math.round((Math.random() - 0.5) * 10), 0, 100),
                        weight: 20
                    });
                }
                records = sample;
                save(STORAGE_DATA, records);
                updateAll();
            }
            document.getElementById('saveSampleBtn').addEventListener('click', loadMathSample);

            function renderDataTable() {
                const t = document.getElementById('dataTable');
                if (!records.length) { t.innerHTML =
                    '<tr><td colspan="9" style="padding:2rem;color:var(--text-muted);">📭 暂无数据</td></tr>'; return; }
                const cols = ['subject', 'date', 'duration', 'practice', 'test', 'review', 'memory', 'weight'];
                const hds = ['学科', '日期', '时长', '练习', '测试', '复习', '记忆', '权重'];
                let h = '<thead><tr>' + hds.map(c => `<th>${c}</th>`).join('') + '<th>操作</th></tr></thead><tbody>';
                records.forEach((r, i) => {
                    h +=
                        `<tr>${cols.map(c => `<td>${r[c]}</td>`).join('')}<td><button class="btn btn-danger btn-sm" onclick="deleteRecord(${i})">删</button></td></tr>`;
                });
                t.innerHTML = h + '</tbody>';
            }

            function renderDashboard() {
                const math = records.filter(r => r.subject === '数学');
                if (!math.length) {
                    document.getElementById('statsCards').innerHTML =
                        '<p style="grid-column:1/-1;text-align:center;color:var(--text-muted);padding:2rem;">📭 暂无数据</p>';
                    document.getElementById('alertContainer').innerHTML = '';
                    return;
                }
                const sorted = [...math].sort((a, b) => new Date(b.date) - new Date(a.date));
                const latest = sorted[0];
                const totalMin = math.reduce((s, r) => s + r.duration, 0);
                const days = new Set(math.map(r => r.date)).size;
                const avgTest = (math.reduce((s, r) => s + r.test, 0) / math.length).toFixed(1);
                const maxTest = Math.max(...math.map(r => r.test));
                const minTest = Math.min(...math.map(r => r.test));
                const avgDur = Math.round(totalMin / math.length);
                const overall = ((latest.practice + latest.test + latest.review + latest.memory) / 4).toFixed(1);

                document.getElementById('statsCards').innerHTML = [
                    { v: overall, l: '综合评分' }, { v: latest.test + '%', l: '最新测试' },
                    { v: Math.floor(totalMin / 60) + 'h ' + totalMin % 60 + 'm', l: '总时长' },
                    { v: days + ' 天', l: '学习天数' }, { v: avgTest + '%', l: '平均测试' },
                    { v: maxTest + '% / ' + minTest + '%', l: '最高 / 最低' },
                    { v: avgDur + ' 分钟', l: '日均时长' },
                ].map(s =>
                    `<div class="stat-card"><div class="stat-value">${s.v}</div><div class="stat-label">${s.l}</div></div>`
                    ).join('');

                const chrono = [...math].sort((a, b) => new Date(a.date) - new Date(b.date)).slice(-100);
                const dates = chrono.map(r => r.date);
                const tests = chrono.map(r => r.test);
                const durs = chrono.map(r => r.duration);

                disposeChart('trend');
                const tc = echarts.init(document.getElementById('trendChart'));
                chartInstances.trend = tc;
                tc.setOption({
                    ...darkTheme,
                    tooltip: { trigger: 'axis' },
                    legend: { bottom: 0, textStyle: { color: '#94a3b8', fontSize: 10 } },
                    grid: { left: 48, right: 48, top: 15, bottom: 40 },
                    xAxis: { type: 'category', data: dates, axisLabel: { rotate: 35, fontSize: 9,
                            color: '#64748b', interval: 'auto' }, axisLine: { lineStyle: { color: 'rgba(255,255,255,0.06)' } } },
                    yAxis: [
                        { type: 'value', name: '%', max: 100, axisLabel: { fontSize: 9,
                            color: '#64748b' }, splitLine: { lineStyle: { color: 'rgba(255,255,255,0.04)' } } },
                        { type: 'value', name: '分', axisLabel: { fontSize: 9,
                            color: '#64748b' }, splitLine: { show: false } }
                    ],
                    series: [
                        { name: '测试%', type: 'line', data: tests, smooth: true, color: '#6d8aff',
                            symbolSize: 3, lineStyle: { width: 2 } },
                        { name: '时长', type: 'bar', yAxisIndex: 1, data: durs, color: 'rgba(167,139,250,0.5)',
                            barWidth: '55%' }
                    ]
                });

                const bins = [0, 40, 50, 60, 70, 80, 90, 101];
                const counts = bins.slice(0, -1).map((low, i) => math.filter(r => r.test >= low && r.test < bins[i +
                    1]).length);
                disposeChart('dist');
                const dc = echarts.init(document.getElementById('distChart'));
                chartInstances.dist = dc;
                dc.setOption({
                    ...darkTheme,
                    tooltip: { trigger: 'axis' },
                    grid: { left: 40, right: 20, top: 10, bottom: 30 },
                    xAxis: { type: 'category', data: ['<40', '40-50', '50-60', '60-70', '70-80', '80-90',
                            '90+'
                        ], axisLabel: { fontSize: 9, color: '#64748b' }, axisLine: { lineStyle: { color: 'rgba(255,255,255,0.06)' } } },
                    yAxis: { type: 'value', name: '天', axisLabel: { fontSize: 9,
                        color: '#64748b' }, splitLine: { lineStyle: { color: 'rgba(255,255,255,0.04)' } } },
                    series: [{ type: 'bar', data: counts, color: '#34d399', barWidth: '55%',
                        itemStyle: { borderRadius: [6, 6, 0, 0] } }]
                });

                disposeChart('radar');
                const rc = echarts.init(document.getElementById('radarChart'));
                chartInstances.radar = rc;
                rc.setOption({
                    ...darkTheme,
                    radar: {
                        indicator: [{ name: '练习', max: 100 }, { name: '测试', max: 100 }, { name: '复习',
                            max: 100 }, { name: '记忆', max: 100 }],
                        radius: '62%',
                        axisName: { color: '#94a3b8', fontSize: 10 },
                        splitArea: { areaStyle: { color: ['rgba(109,138,255,0.02)',
                                'rgba(109,138,255,0.04)'
                            ] } }
                    },
                    series: [{ type: 'radar', data: [{ name: '数学', value: [latest.practice, latest.test,
                            latest.review, latest.memory
                        ], areaStyle: { color: 'rgba(109,138,255,0.15)' }, lineStyle: { color: '#6d8aff',
                            width: 2 } }], symbolSize: 4 }]
                });

                disposeChart('scatter');
                const sc = echarts.init(document.getElementById('scatterChart'));
                chartInstances.scatter = sc;
                sc.setOption({
                    ...darkTheme,
                    grid: { left: 50, right: 25, top: 15, bottom: 30 },
                    xAxis: { type: 'value', name: '练习%', max: 100, axisLabel: { fontSize: 9,
                            color: '#64748b' }, splitLine: { lineStyle: { color: 'rgba(255,255,255,0.04)' } } },
                    yAxis: { type: 'value', name: '测试%', max: 100, axisLabel: { fontSize: 9,
                            color: '#64748b' }, splitLine: { lineStyle: { color: 'rgba(255,255,255,0.04)' } } },
                    series: [{ type: 'scatter', data: math.map(r => [r.practice, r.test]), symbolSize: 5,
                        color: '#fbbf24', itemStyle: { borderColor: 'rgba(0,0,0,0.3)', borderWidth: 1 } }]
                });

                const alerts = [];
                if (latest.test < 60) alerts.push(`测试 ${latest.test}%`);
                if (latest.review < 50) alerts.push(`复习 ${latest.review}%`);
                if (latest.memory < 50) alerts.push(`记忆 ${latest.memory}%`);
                document.getElementById('alertContainer').innerHTML =
                    `<div class="card-header"><h3>⚠️ 预警</h3></div>${alerts.length ? alerts.map(a => `<span class="alert-badge">${a}</span>`).join(' ') : '<span style="color:var(--success);font-size:0.78rem;">✅ 一切正常</span>'}`;
            }

            function updateMotivation() {
                const dates = [...new Set(records.map(r => r.date))].sort();
                const today = new Date().toISOString().split('T')[0];
                let streak = 0,
                    set = new Set(dates),
                    d = new Date(today);
                if (!set.has(today)) d.setDate(d.getDate() - 1);
                while (set.has(d.toISOString().split('T')[0])) { streak++;
                    d.setDate(d.getDate() - 1); }
                motivation.streak = streak;
                const td = records.filter(r => r.date === today);
                if (td.length && today !== motivation.lastCalcDate) {
                    motivation.xp += Math.round(td.reduce((s, r) => s + (r.test / 100) * 10 + r.duration * 0.5, 0));
                    motivation.lastCalcDate = today;
                }
                if (streak >= 7) motivation.achievements.streak7 = true;
                if (records.reduce((s, r) => s + r.duration, 0) > 3000) motivation.achievements.total3k = true;
                const lastDate = dates.length ? new Date(dates[dates.length - 1]) : new Date(0);
                const diff = Math.floor((new Date() - lastDate) / 86400000);
                if (diff >= 3 && (!motivation.lastPenaltyDate || motivation.lastPenaltyDate !== today)) {
                    motivation.xp = Math.max(0, motivation.xp - 10);
                    motivation.penaltyLog.unshift({ date: today, reason: `连续 ${diff} 天未学习，扣除 10 经验` });
                    motivation.lastPenaltyDate = today;
                }
                save(STORAGE_MOTIV, motivation);
                document.getElementById('streakDays').textContent = streak;
                document.getElementById('xpValue').textContent = motivation.xp;
                document.getElementById('latestBadge').textContent = motivation.achievements.streak7 ? '🔥 连续7天' :
                    '';
            }

            function renderMotivation() {
                const badges = [{ id: 'streak7', icon: '🔥', label: '连续学习7天' }, { id: 'total3k', icon: '⏱️',
                    label: '总时长超3000分钟' }];
                document.getElementById('achievementList').innerHTML = badges.map(b => `
                <div class="stat-card" style="background:${motivation.achievements[b.id]?'rgba(52,211,153,0.1)':'rgba(255,255,255,0.02)'};border-color:${motivation.achievements[b.id]?'rgba(52,211,153,0.25)':'var(--border)'};">
                  <div style="font-size:1.8rem;margin-bottom:0.1rem;">${b.icon}</div>
                  <div class="stat-label">${b.label}</div>
                  <div style="font-size:0.62rem;margin-top:0.2rem;color:${motivation.achievements[b.id]?'var(--success)':'var(--text-muted)'};">${motivation.achievements[b.id]?'✅ 已解锁':'🔒 未解锁'}</div>
                </div>`).join('');
                document.getElementById('penaltyLog').innerHTML = motivation.penaltyLog.length ? motivation
                    .penaltyLog.map(p =>
                        `<div style="padding:0.35rem 0;border-bottom:1px solid var(--border);">📅 ${p.date} — ${p.reason}</div>`
                        ).join('') : '<p style="color:var(--text-muted);padding:0.4rem 0;">🎉 暂无惩罚记录</p>';
            }

            function detectAnomalies(scores, dates) {
                const mean = scores.reduce((a, b) => a + b, 0) / scores.length;
                const variance = scores.reduce((s, v) => s + (v - mean) ** 2, 0) / scores.length;
                const std = Math.sqrt(variance);
                const threshold = 2 * std;
                return scores.reduce((arr, v, i) => {
                    if (Math.abs(v - mean) > threshold) arr.push({ date: dates[i], score: v, deviation: ((v -
                            mean) / std).toFixed(2) });
                    return arr;
                }, []);
            }

            function getRecentAll(recs) {
                const ds = [...new Set(recs.map(r => r.date))].sort();
                return recs.filter(r => ds.includes(r.date)).sort((a, b) => new Date(a.date) - new Date(b.date));
            }

            function renderHistory() {
                const recs = records.filter(r => r.subject === '数学');
                const recent = getRecentAll(recs).slice(-100);
                const cd = document.getElementById('historyChart'),
                    td = document.getElementById('historyTable');
                if (recent.length < 2) { cd.innerHTML =
                    '<p style="text-align:center;padding:3rem;color:var(--text-muted);">数据不足（需≥2条）</p>';
                    td.innerHTML = ''; return; }
                const dates = recent.map(r => r.date),
                    tests = recent.map(r => r.test);
                const anomalies = detectAnomalies(tests, dates);
                window._anomalies = anomalies;
                disposeChart('history');
                const chart = echarts.init(cd);
                chartInstances.history = chart;
                chart.setOption({
                    ...darkTheme,
                    tooltip: { trigger: 'axis' },
                    grid: { left: 45, right: 30, top: 15, bottom: 35 },
                    xAxis: { type: 'category', data: dates, axisLabel: { rotate: 45, fontSize: 9,
                            color: '#64748b', interval: 'auto' }, axisLine: { lineStyle: { color: 'rgba(255,255,255,0.06)' } } },
                    yAxis: { max: 100, axisLabel: { fontSize: 9, color: '#64748b' },
                        splitLine: { lineStyle: { color: 'rgba(255,255,255,0.04)' } } },
                    series: [{
                        name: '测试%',
                        type: 'line',
                        data: tests,
                        smooth: true,
                        color: '#6d8aff',
                        symbolSize: 3,
                        lineStyle: { width: 2 },
                        areaStyle: { color: 'rgba(109,138,255,0.06)' },
                        markPoint: {
                            data: anomalies.map(a => ({
                                name: a.date,
                                coord: [a.date, a.score],
                                value: a.score,
                                symbol: 'pin',
                                symbolSize: 32,
                                itemStyle: { color: '#f87171' },
                                label: { formatter: '{c}', fontSize: 8, color: '#fff' }
                            }))
                        }
                    }]
                });
                let h = '<thead><tr><th>日期</th><th>测试%</th><th>练习%</th><th>复习%</th><th>记忆%</th><th>状态</th></tr></thead><tbody>';
                recent.forEach(r => {
                    const isA = anomalies.some(a => a.date === r.date);
                    h +=
                        `<tr><td>${r.date}</td><td>${r.test}${isA?' <span class="anomaly-badge">异常</span>':''}</td><td>${r.practice}</td><td>${r.review}</td><td>${r.memory}</td><td>${isA?'⚠️':''}</td></tr>`;
                });
                td.innerHTML = h + '</tbody>';
            }

            function renderAnomalies() {
                const table = document.getElementById('anomalyTable');
                const anoms = window._anomalies || [];
                if (!anoms.length) { table.innerHTML =
                    '<thead><tr><th>日期</th><th>测试%</th><th>偏离 σ</th></tr></thead><tbody><tr><td colspan="3">✅ 暂无异常</td></tr></tbody>'; return; }
                let h = '<thead><tr><th>日期</th><th>测试%</th><th>偏离 σ 倍数</th></tr></thead><tbody>';
                anoms.forEach(a => { h +=
                        `<tr><td>${a.date}</td><td>${a.score}%</td><td><span class="anomaly-badge">${a.deviation}σ</span></td></tr>`; });
                table.innerHTML = h + '</tbody>';
            }

            function renderForecast() {
                const recs = records.filter(r => r.subject === '数学');
                const recent = getRecentAll(recs).slice(-30);
                if (recent.length < 3) return;
                const scores = recent.map(r => r.test),
                    n = scores.length;
                const current = scores[n - 1];
                const sumX = scores.reduce((s, _, i) => s + i, 0),
                    sumY = scores.reduce((s, v) => s + v, 0);
                const sumXY = scores.reduce((s, v, i) => s + i * v, 0),
                    sumX2 = scores.reduce((s, _, i) => s + i * i, 0);
                const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
                const intercept = (sumY - slope * sumX) / n;
                const days = [0, 5, 10, 15, 20, 25, 30];
                const baseline = days.map(d => clamp(intercept + slope * (n - 1 + d), 0, 100)).map(v => +v.toFixed(
                    1));
                baseline[0] = +current.toFixed(1);
                const { target, sMid, sK } = forecastSettings;
                const uplift = days.map(d => d === 0 ? current : clamp(current + (target - current) / (1 + Math.exp(
                    -sK * (d - sMid))), 0, 100)).map(v => +v.toFixed(1));
                const { halfLife, minScore } = forecastSettings;
                const decayRate = Math.log(2) / halfLife;
                const downside = days.map(d => d === 0 ? current : Math.max(minScore, minScore + (current -
                    minScore) * Math.exp(-decayRate * d))).map(v => +v.toFixed(1));
                document.getElementById('forecastInfo').innerHTML =
                    `📊 当前测试正确率 <strong>${current.toFixed(1)}%</strong> → 目标 <strong>${target}%</strong>`;
                const labels = ['今天', '+5d', '+10d', '+15d', '+20d', '+25d', '+30d'];
                document.getElementById('forecastTable').innerHTML =
                    '<thead><tr><th>类别</th>' + labels.map(l => `<th>${l}</th>`).join('') + '</tr></thead><tbody>' +
                    `<tr style="background:rgba(109,138,255,0.06);"><td>📈 基准线</td>${baseline.map(v => `<td>${v}%</td>`).join('')}</tr>` +
                    `<tr style="background:rgba(52,211,153,0.06);"><td>🚀 S型提升</td>${uplift.map(v => `<td>${v}%</td>`).join('')}</tr>` +
                    `<tr style="background:rgba(248,113,113,0.06);"><td>📉 遗忘底线</td>${downside.map(v => `<td>${v}%</td>`).join('')}</tr></tbody>`;
                disposeChart('forecast');
                const chart = echarts.init(document.getElementById('forecastChart'));
                chartInstances.forecast = chart;
                chart.setOption({
                    ...darkTheme,
                    legend: { bottom: 0, textStyle: { color: '#94a3b8', fontSize: 10 } },
                    grid: { left: 45, right: 25, top: 15, bottom: 40 },
                    xAxis: { type: 'category', data: labels, axisLabel: { fontSize: 9,
                        color: '#64748b' }, axisLine: { lineStyle: { color: 'rgba(255,255,255,0.06)' } } },
                    yAxis: { max: 100, min: 0, axisLabel: { fontSize: 9,
                        color: '#64748b' }, splitLine: { lineStyle: { color: 'rgba(255,255,255,0.04)' } } },
                    series: [
                        { name: '基准线', type: 'line', data: baseline, smooth: true, color: '#6d8aff',
                            symbolSize: 6, lineStyle: { width: 2.5 } },
                        { name: 'S型提升', type: 'line', data: uplift, smooth: true, lineStyle: { type: 'dashed',
                                color: '#34d399', width: 2 }, symbol: 'diamond', symbolSize: 7 },
                        { name: '遗忘底线', type: 'line', data: downside, smooth: true, lineStyle: { type: 'dotted',
                                color: '#f87171', width: 2 }, symbol: 'triangle', symbolSize: 7 }
                    ]
                });
            }

            document.getElementById('settingsToggle').addEventListener('click', () => {
                const b = document.getElementById('settingsBody');
                const a = document.getElementById('settingsArrow');
                b.style.display = b.style.display === 'none' ? 'block' : 'none';
                a.textContent = b.style.display === 'none' ? '▶' : '▼';
            });
            document.getElementById('savePredictSettingsBtn').addEventListener('click', () => {
                forecastSettings = {
                    target: +document.getElementById('targetScore').value || 90,
                    sMid: +document.getElementById('sMid').value || 15,
                    sK: +document.getElementById('sK').value || 0.3,
                    halfLife: +document.getElementById('halfLife').value || 12,
                    minScore: +document.getElementById('minScore').value || 40
                };
                save(STORAGE_SETTINGS, forecastSettings);
                const hint = document.getElementById('settingsSavedHint');
                hint.style.display = 'inline';
                setTimeout(() => hint.style.display = 'none', 1800);
                renderForecast();
            });

            function loadSettingsUI() {
                document.getElementById('targetScore').value = forecastSettings.target;
                document.getElementById('sMid').value = forecastSettings.sMid;
                document.getElementById('sK').value = forecastSettings.sK;
                document.getElementById('halfLife').value = forecastSettings.halfLife;
                document.getElementById('minScore').value = forecastSettings.minScore;
            }

            document.getElementById('addRecordNoteBtn').addEventListener('click', () => {
                const date = document.getElementById('recordDate').value;
                const text = document.getElementById('recordText').value.trim();
                const imgInput = document.getElementById('recordImage');
                if (!date) return alert('请选择日期');
                if (!text && !imgInput.files[0]) return alert('请至少填写文字或上传图片');
                const done = (img = '') => {
                    studyNotes.unshift({ date, text, imageBase64: img });
                    save(STORAGE_NOTES, studyNotes);
                    renderNotes();
                    document.getElementById('recordDate').valueAsDate = new Date();
                    document.getElementById('recordText').value = '';
                    imgInput.value = '';
                };
                if (imgInput.files[0]) {
                    const r = new FileReader();
                    r.onload = e => done(e.target.result);
                    r.readAsDataURL(imgInput.files[0]);
                } else done();
            });

            function renderNotes() {
                const c = document.getElementById('recordNotesContainer');
                if (!studyNotes.length) { c.innerHTML =
                    '<p style="text-align:center;color:var(--text-muted);padding:2rem;">📝 暂无学习记录</p>'; return; }
                c.innerHTML = studyNotes.map(n => `
                <div class="record-entry">
                  <div style="font-weight:650;color:var(--accent);">${n.date}</div>
                  <div style="white-space:pre-wrap;margin:0.35rem 0;">${n.text || '(无文字)'}</div>
                  ${n.imageBase64 ? `<img src="${n.imageBase64}" alt="照片" loading="lazy">` : ''}
                </div>`).join('');
            }

            function updateAll() {
                renderDataTable();
                updateMotivation();
                const activeView = document.querySelector('.view.active');
                if (activeView) {
                    const id = activeView.id;
                    if (id === 'view-dashboard') renderDashboard();
                    if (id === 'view-predict') { renderHistory();
                        renderForecast();
                        renderAnomalies(); }
                    if (id === 'view-records') renderNotes();
                    if (id === 'view-motivation') renderMotivation();
                }
            }

            window.addEventListener('resize', () => {
                Object.values(chartInstances).forEach(c => { try { c.resize(); } catch (e) {} });
            });

            // 启动
            loadAll();
            if (!records.length) loadMathSample();
            else updateAll();
            loadSettingsUI();
            document.getElementById('recordDate').valueAsDate = new Date();
        })();
    </script>
</body>
</html>
