<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
    <title>数学量化 · Self-Quantified Learning Dashboard</title>
    <meta name="description" content="自学量化平台 — 用数据驱动数学学习，150天模拟数据 + ECharts 仪表盘 + 预测模型">
    <meta name="theme-color" content="#0f172a">

    <!-- ECharts -->
    <script src="https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js">
    </script>
    <!-- SheetJS -->
    <script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js">
    </script>

    <style>
        :root {
            /* 深色主题 */
            --bg-deep: #0b0f1c;
            --bg-card: rgba(18, 23, 41, 0.75);
            --bg-card-solid: #121729;
            --border-subtle: rgba(255, 255, 255, 0.06);
            --border-glow: rgba(99, 130, 255, 0.25);
            --text-primary: #e8ecf4;
            --text-secondary: #8e97b0;
            --text-muted: #5c6480;

            /* 强调色 */
            --accent: #6d8aff;
            --accent-glow: rgba(109, 138, 255, 0.35);
            --success: #34d399;
            --warning: #fbbf24;
            --danger: #f87171;
            --purple: #a78bfa;
            --cyan: #22d3ee;

            --radius-sm: 10px;
            --radius: 16px;
            --radius-lg: 20px;
            --shadow-card: 0 4px 32px rgba(0, 0, 0, 0.35), 0 1px 0 rgba(255, 255, 255, 0.04) inset;
            --shadow-glow: 0 0 40px rgba(109, 138, 255, 0.12);
            --blur-glass: blur(18px) saturate(120%);
            --transition: 0.25s cubic-bezier(0.4, 0, 0.2, 1);
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        html {
            scroll-behavior: smooth;
            font-size: 15px;
        }

        body {
            font-family: 'Inter', 'SF Pro Display', 'PingFang SC', 'Noto Sans SC', system-ui, -apple-system, sans-serif;
            background: var(--bg-deep);
            color: var(--text-primary);
            display: flex;
            height: 100vh;
            height: 100dvh;
            overflow: hidden;
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
            letter-spacing: -0.01em;
        }

        /* ========== 动态背景粒子 ========== */
        .bg-particles {
            position: fixed;
            inset: 0;
            z-index: 0;
            pointer-events: none;
            overflow: hidden;
        }
        .bg-particles .orb {
            position: absolute;
            border-radius: 50%;
            filter: blur(100px);
            opacity: 0.13;
            animation: float 18s ease-in-out infinite alternate;
        }
        .bg-particles .orb:nth-child(1) {
            width: 500px;
            height: 500px;
            background: var(--accent);
            top: -15%;
            left: -10%;
            animation-delay: 0s;
        }
        .bg-particles .orb:nth-child(2) {
            width: 400px;
            height: 400px;
            background: var(--purple);
            bottom: -20%;
            right: -8%;
            animation-delay: -6s;
        }
        .bg-particles .orb:nth-child(3) {
            width: 300px;
            height: 300px;
            background: var(--cyan);
            top: 50%;
            left: 40%;
            animation-delay: -12s;
        }

        @keyframes float {
            0% {
                transform: translate(0, 0) scale(1);
            }
            33% {
                transform: translate(60px, -50px) scale(1.15);
            }
            66% {
                transform: translate(-30px, 40px) scale(0.9);
            }
            100% {
                transform: translate(20px, -20px) scale(1.05);
            }
        }

        /* 网格点阵 */
        .bg-grid {
            position: fixed;
            inset: 0;
            z-index: 0;
            pointer-events: none;
            background-image: radial-gradient(rgba(255, 255, 255, 0.03) 1.2px, transparent 1.2px);
            background-size: 32px 32px;
            mask-image: linear-gradient(to bottom, transparent 0%, #000 15%, #000 85%, transparent 100%);
        }

        /* ========== 侧边栏 ========== */
        .sidebar {
            width: 250px;
            background: rgba(15, 20, 38, 0.7);
            backdrop-filter: var(--blur-glass);
            -webkit-backdrop-filter: var(--blur-glass);
            border-right: 1px solid var(--border-subtle);
            display: flex;
            flex-direction: column;
            padding: 1.8rem 0;
            flex-shrink: 0;
            z-index: 10;
            position: relative;
        }
        .sidebar-logo {
            padding: 0 1.6rem 1.4rem;
            border-bottom: 1px solid var(--border-subtle);
            margin-bottom: 1rem;
            display: flex;
            align-items: center;
            gap: 0.6rem;
        }
        .sidebar-logo .icon {
            width: 38px;
            height: 38px;
            border-radius: 10px;
            background: linear-gradient(135deg, var(--accent), var(--purple));
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.2rem;
            box-shadow: 0 0 20px var(--accent-glow);
        }
        .sidebar-logo .title {
            font-weight: 700;
            font-size: 1.05rem;
            color: var(--text-primary);
            letter-spacing: -0.02em;
        }
        .sidebar-logo .subtitle {
            font-size: 0.7rem;
            color: var(--text-muted);
            font-weight: 500;
        }
        .nav-item {
            padding: 0.65rem 1.6rem;
            margin: 0.1rem 0.6rem;
            cursor: pointer;
            border-radius: 8px;
            transition: var(--transition);
            color: var(--text-secondary);
            display: flex;
            align-items: center;
            gap: 0.55rem;
            font-weight: 500;
            font-size: 0.88rem;
            position: relative;
            user-select: none;
        }
        .nav-item:hover {
            background: rgba(255, 255, 255, 0.04);
            color: var(--text-primary);
        }
        .nav-item.active {
            background: rgba(109, 138, 255, 0.12);
            color: var(--accent);
            font-weight: 600;
            box-shadow: 0 0 0 1px rgba(109, 138, 255, 0.15) inset;
        }
        .nav-item .badge {
            margin-left: auto;
            font-size: 0.65rem;
            background: var(--accent);
            color: #fff;
            padding: 0.15rem 0.5rem;
            border-radius: 20px;
            font-weight: 700;
        }

        /* ========== 主区域 ========== */
        .main {
            flex: 1;
            display: flex;
            flex-direction: column;
            overflow-y: auto;
            overflow-x: hidden;
            padding: 1.5rem 1.8rem;
            z-index: 5;
            position: relative;
            gap: 1.3rem;
        }
        .main::-webkit-scrollbar {
            width: 5px;
        }
        .main::-webkit-scrollbar-thumb {
            background: rgba(255, 255, 255, 0.08);
            border-radius: 10px;
        }

        /* 顶栏 */
        .top-strip {
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 0.7rem;
            padding: 0.6rem 1.2rem;
            background: var(--bg-card);
            backdrop-filter: var(--blur-glass);
            -webkit-backdrop-filter: var(--blur-glass);
            border-radius: var(--radius);
            border: 1px solid var(--border-subtle);
        }
        .top-strip .pill {
            display: inline-flex;
            align-items: center;
            gap: 0.35rem;
            padding: 0.35rem 0.8rem;
            border-radius: 20px;
            font-size: 0.78rem;
            font-weight: 600;
            background: rgba(255, 255, 255, 0.04);
            color: var(--text-secondary);
            border: 1px solid var(--border-subtle);
        }
        .top-strip .pill.accent {
            background: rgba(109, 138, 255, 0.12);
            color: var(--accent);
            border-color: rgba(109, 138, 255, 0.2);
        }
        .top-strip .pill.warn {
            background: rgba(251, 191, 36, 0.1);
            color: var(--warning);
            border-color: rgba(251, 191, 36, 0.2);
        }

        /* 视图 */
        .view {
            display: none;
            flex-direction: column;
            gap: 1.3rem;
            animation: fadeSlideIn 0.35s ease;
        }
        .view.active {
            display: flex;
        }
        @keyframes fadeSlideIn {
            from {
                opacity: 0;
                transform: translateY(12px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }

        /* 卡片 */
        .card {
            background: var(--bg-card);
            backdrop-filter: var(--blur-glass);
            -webkit-backdrop-filter: var(--blur-glass);
            border-radius: var(--radius);
            padding: 1.4rem;
            border: 1px solid var(--border-subtle);
            box-shadow: var(--shadow-card);
            transition: var(--transition);
        }
        .card:hover {
            border-color: var(--border-glow);
            box-shadow: var(--shadow-card), var(--shadow-glow);
        }
        .card-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 1rem;
            flex-wrap: wrap;
            gap: 0.5rem;
        }
        .card-header h3 {
            font-size: 0.95rem;
            font-weight: 650;
            color: var(--text-primary);
            letter-spacing: -0.01em;
        }
        .card-header .chip {
            font-size: 0.7rem;
            padding: 0.2rem 0.7rem;
            border-radius: 20px;
            font-weight: 600;
        }

        /* 统计卡片网格 */
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(155px, 1fr));
            gap: 0.8rem;
        }
        .stat-card {
            background: var(--bg-card);
            backdrop-filter: var(--blur-glass);
            -webkit-backdrop-filter: var(--blur-glass);
            border-radius: var(--radius);
            padding: 1rem 1.1rem;
            border: 1px solid var(--border-subtle);
            text-align: center;
            transition: var(--transition);
            position: relative;
            overflow: hidden;
        }
        .stat-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 3px;
            background: linear-gradient(90deg, var(--accent), var(--purple));
            opacity: 0;
            transition: var(--transition);
        }
        .stat-card:hover::before {
            opacity: 1;
        }
        .stat-card:hover {
            transform: translateY(-3px);
            border-color: var(--border-glow);
            box-shadow: var(--shadow-glow);
        }
        .stat-value {
            font-size: 1.7rem;
            font-weight: 750;
            background: linear-gradient(135deg, var(--accent), var(--cyan));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            letter-spacing: -0.03em;
            line-height: 1.2;
        }
        .stat-label {
            font-size: 0.72rem;
            color: var(--text-muted);
            margin-top: 0.2rem;
            font-weight: 500;
        }

        /* 双列 */
        .grid-2 {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(380px, 1fr));
            gap: 1.2rem;
        }
        .chart-container {
            width: 100%;
            height: 270px;
        }

        /* 按钮 */
        .btn {
            border: 1px solid var(--border-subtle);
            padding: 0.5rem 1rem;
            border-radius: 8px;
            font-weight: 600;
            cursor: pointer;
            font-size: 0.8rem;
            transition: var(--transition);
            white-space: nowrap;
            display: inline-flex;
            align-items: center;
            gap: 0.3rem;
            background: rgba(255, 255, 255, 0.04);
            color: var(--text-secondary);
            user-select: none;
            -webkit-tap-highlight-color: transparent;
        }
        .btn:hover {
            background: rgba(255, 255, 255, 0.08);
            color: var(--text-primary);
            border-color: rgba(255, 255, 255, 0.15);
        }
        .btn:active {
            transform: scale(0.96);
        }
        .btn-primary {
            background: linear-gradient(135deg, var(--accent), #5b7ef5);
            color: #fff;
            border: none;
            box-shadow: 0 4px 16px var(--accent-glow);
        }
        .btn-primary:hover {
            box-shadow: 0 6px 24px rgba(109, 138, 255, 0.45);
            color: #fff;
        }
        .btn-danger {
            background: rgba(248, 113, 113, 0.15);
            color: var(--danger);
            border-color: rgba(248, 113, 113, 0.25);
        }
        .btn-danger:hover {
            background: rgba(248, 113, 113, 0.25);
            color: #fca5a5;
        }
        .btn-sm {
            padding: 0.35rem 0.7rem;
            font-size: 0.7rem;
            border-radius: 6px;
        }

        /* 表单 */
        .form-row {
            display: flex;
            gap: 0.5rem;
            flex-wrap: wrap;
            align-items: flex-end;
        }
        .form-group {
            display: flex;
            flex-direction: column;
            min-width: 85px;
        }
        .form-group label {
            font-size: 0.68rem;
            margin-bottom: 0.2rem;
            color: var(--text-muted);
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.04em;
        }
        .form-group input,
        .form-group textarea {
            padding: 0.5rem 0.65rem;
            border: 1.5px solid var(--border-subtle);
            border-radius: 8px;
            font-size: 0.82rem;
            background: rgba(255, 255, 255, 0.03);
            color: var(--text-primary);
            transition: var(--transition);
            font-family: inherit;
            outline: none;
        }
        .form-group input:focus,
        .form-group textarea:focus {
            border-color: var(--accent);
            box-shadow: 0 0 0 3px rgba(109, 138, 255, 0.1);
        }
        .form-group textarea {
            resize: vertical;
            min-height: 60px;
        }

        /* 上传区 */
        .upload-area {
            border: 2px dashed var(--border-subtle);
            border-radius: var(--radius);
            padding: 1.5rem;
            text-align: center;
            cursor: pointer;
            transition: var(--transition);
            color: var(--text-muted);
            font-size: 0.82rem;
        }
        .upload-area:hover {
            border-color: var(--accent);
            background: rgba(109, 138, 255, 0.04);
            color: var(--accent);
        }

        /* 表格 */
        .table-wrapper {
            max-height: 380px;
            overflow: auto;
            border-radius: var(--radius-sm);
            border: 1px solid var(--border-subtle);
        }
        .table-wrapper::-webkit-scrollbar {
            height: 4px;
            width: 4px;
        }
        .table-wrapper::-webkit-scrollbar-thumb {
            background: rgba(255, 255, 255, 0.06);
            border-radius: 10px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.75rem;
            white-space: nowrap;
        }
        th,
        td {
            padding: 0.55rem 0.7rem;
            border-bottom: 1px solid var(--border-subtle);
            text-align: center;
        }
        th {
            background: rgba(255, 255, 255, 0.03);
            position: sticky;
            top: 0;
            z-index: 2;
            font-weight: 650;
            color: var(--text-secondary);
            font-size: 0.7rem;
            text-transform: uppercase;
            letter-spacing: 0.03em;
        }
        tr:hover td {
            background: rgba(255, 255, 255, 0.02);
        }

        /* 徽章 */
        .alert-badge {
            padding: 0.25rem 0.7rem;
            border-radius: 20px;
            font-size: 0.7rem;
            font-weight: 600;
            background: rgba(248, 113, 113, 0.15);
            color: var(--danger);
            display: inline-block;
            margin: 0.1rem;
        }
        .anomaly-badge {
            background: rgba(248, 113, 113, 0.2);
            color: #fca5a5;
            padding: 0.15rem 0.5rem;
            border-radius: 20px;
            font-weight: 700;
            font-size: 0.65rem;
            white-space: nowrap;
        }
        .info-row {
            padding: 0.5rem 0.8rem;
            background: rgba(109, 138, 255, 0.06);
            border-radius: 8px;
            font-size: 0.78rem;
            color: var(--text-secondary);
            margin-bottom: 0.6rem;
            border: 1px solid rgba(109, 138, 255, 0.1);
        }
        .info-row strong {
            color: var(--accent);
        }

        /* 折叠 */
        .collapse-header {
            cursor: pointer;
            padding: 0.4rem 0;
            display: flex;
            justify-content: space-between;
            align-items: center;
            user-select: none;
            color: var(--text-secondary);
            font-size: 0.85rem;
            transition: var(--transition);
        }
        .collapse-header:hover {
            color: var(--text-primary);
        }

        /* 记录卡片 */
        .record-entry {
            background: rgba(255, 255, 255, 0.02);
            border: 1px solid var(--border-subtle);
            border-radius: var(--radius);
            padding: 0.8rem;
            margin-bottom: 0.6rem;
            font-size: 0.8rem;
            transition: var(--transition);
        }
        .record-entry:hover {
            border-color: var(--border-glow);
        }
        .record-entry img {
            max-width: 100%;
            max-height: 200px;
            margin: 0.4rem 0;
            border-radius: 8px;
            object-fit: contain;
        }

        /* 预测表格 */
        .prediction-table th {
            background: rgba(109, 138, 255, 0.2);
            color: var(--accent);
        }
        .prediction-table td {
            font-weight: 600;
            font-size: 0.8rem;
        }

        /* ========== 响应式 ========== */
        @media (max-width: 768px) {
            body {
                flex-direction: column;
            }
            .sidebar {
                width: 100%;
                flex-direction: row;
                overflow-x: auto;
                padding: 0.5rem;
                flex-shrink: 0;
                border-right: none;
                border-bottom: 1px solid var(--border-subtle);
                gap: 0.2rem;
            }
            .sidebar-logo {
                display: none;
            }
            .nav-item {
                padding: 0.45rem 0.7rem;
                margin: 0;
                border-radius: 6px;
                font-size: 0.75rem;
                white-space: nowrap;
                flex-shrink: 0;
            }
            .main {
                padding: 0.8rem;
                gap: 0.8rem;
            }
            .stats-grid {
                grid-template-columns: repeat(2, 1fr);
                gap: 0.5rem;
            }
            .stat-value {
                font-size: 1.3rem;
            }
            .grid-2 {
                grid-template-columns: 1fr;
                gap: 0.8rem;
            }
            .chart-container {
                height: 220px;
            }
            .card {
                padding: 1rem;
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
                height: 320px;
            }
        }
    </style>

    <!-- Inter 字体 -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
</head>
<body>

    <!-- 动态背景 -->
    <div class="bg-particles">
        <div class="orb"></div>
        <div class="orb"></div>
        <div class="orb"></div>
    </div>
    <div class="bg-grid"></div>

    <!-- 侧边栏 -->
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

    <!-- 主区域 -->
    <main class="main">
        <!-- 顶栏 -->
        <div class="top-strip">
            <span class="pill accent">🔥 连续 <strong id="streakDays">0</strong> 天</span>
            <span class="pill">⭐ 经验 <strong id="xpValue">0</strong></span>
            <span class="pill warn" id="latestBadge"></span>
        </div>

        <!-- 数据管理 -->
        <section class="view active" id="view-data">
            <div class="card-header">
                <h3>📁 数据管理</h3>
                <div style="display:flex;gap:0.4rem;">
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
                <div class="upload-area" id="uploadArea" style="margin-top:0.6rem;">
                    📂 点击上传 Excel 文件（.xlsx / .xls / .csv）
                    <input type="file" id="fileInput" accept=".xlsx,.xls,.csv" hidden>
                </div>
                <div class="table-wrapper"><table id="dataTable"><thead></thead><tbody></tbody></table></div>
            </div>
        </section>

        <!-- 量化仪表盘 -->
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
                <div class="chart-container" id="forecastChart" style="margin-top:0.6rem;"></div>
            </div>
        </section>

        <!-- 学习记录 -->
        <section class="view" id="view-records">
            <div class="card-header"><h3>📓 学习情况记录</h3></div>
            <div class="card">
                <div class="card-header"><h3>📝 添加新记录</h3></div>
                <div class="form-row">
                    <div class="form-group"><label>日期</label><input type="date" id="recordDate"></div>
                    <div class="form-group" style="flex:1;min-width:200px;"><label>文字记录</label><textarea id="recordText" rows="2" placeholder="今天学了什么…"></textarea></div>
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
                <div id="penaltyLog" style="font-size:0.8rem;color:var(--text-secondary);"></div>
            </div>
        </section>
    </main>

    <!-- GitHub 角标 -->
    <a href="https://github.com" target="_blank" rel="noopener" style="position:fixed;bottom:16px;right:16px;z-index:100;opacity:0.4;transition:opacity 0.3s;"
    onmouseover="this.style.opacity='1'" onmouseout="this.style.opacity='0.4'">
    <svg width="28" height="28" viewBox="0 0 24 24" fill="white"><path d="M12 0C5.37 0 0 5.37 0 12c0 5.31 3.435 9.795 8.205 11.385.6.105.825-.255.825-.57 0-.285-.015-1.23-.015-2.235-3.015.555-3.795-.735-4.035-1.41-.135-.345-.72-1.41-1.23-1.695-.42-.225-1.02-.78-.015-.795.945-.015 1.62.87 1.845 1.23 1.08 1.815 2.805 1.305 3.495.99.105-.78.42-1.305.765-1.605-2.67-.3-5.46-1.335-5.46-5.925 0-1.305.465-2.385 1.23-3.225-.12-.3-.54-1.53.12-3.18 0 0 1.005-.315 3.3 1.23.96-.27 1.98-.405 3-.405s2.04.135 3 .405c2.295-1.56 3.3-1.23 3.3-1.23.66 1.65.24 2.88.12 3.18.765.84 1.23 1.905 1.23 3.225 0 4.605-2.805 5.625-5.475 5.925.435.375.81 1.095.81 2.22 0 1.605-.015 2.895-.015 3.3 0 .315.225.69.825.57A12.02 12.02 0 0024 12c0-6.63-5.37-12-12-12z"/></svg>
</a>

<script>
    (function() {
        /* ============================================
           自学量化平台 · 数学专属 — 核心逻辑
           ============================================ */
        const STORAGE_DATA = 'sq_math_data_v4';
        const STORAGE_MOTIV = 'sq_math_motiv_v4';
        const STORAGE_SETTINGS = 'sq_math_settings_v4';
        const STORAGE_NOTES = 'sq_math_notes_v4';

        let records = [];
        let motivation = { xp: 0, streak: 0, achievements: {}, penaltyLog: [], lastCalcDate: '',
            lastPenaltyDate: '' };
        let forecastSettings = { target: 90, sMid: 15, sK: 0.3, halfLife: 12, minScore: 40 };
        let studyNotes = [];
        const chartInstances = {};

        // ECharts 深色主题通用配置
        const darkTheme = {
            textStyle: { color: '#8e97b0' },
            tooltip: { backgroundColor: 'rgba(18,23,41,0.95)', borderColor: 'rgba(109,138,255,0.3)',
                textStyle: { color: '#e8ecf4' } },
        };

        function disposeChart(key) { if (chartInstances[key]) { chartInstances[key].dispose(); delete chartInstances[
                key]; } }

        // 持久化
        function save(k, v) { localStorage.setItem(k, JSON.stringify(v)); }

        function load(k) { const s = localStorage.getItem(k); return s ? JSON.parse(s) : null; }

        function loadAll() {
            records = load(STORAGE_DATA) || [];
            motivation = load(STORAGE_MOTIV) || { xp: 0, streak: 0, achievements: {}, penaltyLog: [],
                lastCalcDate: '', lastPenaltyDate: '' };
            forecastSettings = load(STORAGE_SETTINGS) || { target: 90, sMid: 15, sK: 0.3, halfLife: 12, minScore: 40 };
            studyNotes = load(STORAGE_NOTES) || [];
        }

        // ============ 导航 ============
        document.querySelectorAll('.nav-item').forEach(item => {
            item.addEventListener('click', () => {
                const vn = item.dataset.view;
                document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
                item.classList.add('active');
                document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
                const viewEl = document.getElementById(`view-${vn}`);
                if (viewEl) viewEl.classList.add('active');
                // 按需渲染
                if (vn === 'data') renderDataTable();
                if (vn === 'dashboard') renderDashboard();
                if (vn === 'predict') { renderHistory();
                    renderForecast();
                    renderAnomalies(); }
                if (vn === 'records') renderNotes();
                if (vn === 'motivation') renderMotivation();
            });
        });

        // ============ 数据录入 ============
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

        function clamp(v, min, max) { return Math.max(min, Math.min(max, v || 0)); }

        window.deleteRecord = (idx) => { if (confirm('确认删除这条记录？')) { records.splice(idx, 1);
                save(STORAGE_DATA, records);
                updateAll(); } };

        document.getElementById('deleteAllBtn').addEventListener('click', () => {
            if (confirm('⚠️ 确定清空全部数据？此操作不可恢复。')) { records = [];
                save(STORAGE_DATA, records);
                updateAll(); }
        });

        // ============ Excel 导入 ============
        const uploadArea = document.getElementById('uploadArea'),
            fileInput = document.getElementById('fileInput');
        uploadArea.addEventListener('click', () => fileInput.click());
        uploadArea.addEventListener('dragover', e => { e.preventDefault();
            uploadArea.style.borderColor = 'var(--accent)'; });
        uploadArea.addEventListener('dragleave', () => { uploadArea.style.borderColor = 'var(--border-subtle)'; });
        uploadArea.addEventListener('drop', e => {
            e.preventDefault();
            uploadArea.style.borderColor = 'var(--border-subtle)';
            if (e.dataTransfer.files[0]) processExcel(e.dataTransfer.files[0]);
        });
        fileInput.addEventListener('change', e => { if (e.target.files[0]) processExcel(e.target.files[0]); });

        function processExcel(file) {
            const reader = new FileReader();
            reader.onload = e => {
                try {
                    const wb = XLSX.read(e.target.result, { type: 'binary' });
                    const ws = wb.Sheets[wb.SheetNames[0]];
                    const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
                    if (json.length < 2) return alert('需要标题行');
                    const headers = json[0];
                    const map = { '学科': 'subject', '日期': 'date', '学习时长(分钟)': 'duration', '练习正确率%': 'practice',
                        '测试正确率%': 'test', '复习完成率%': 'review', '知识点记忆%': 'memory', '权重': 'weight' };
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
                    alert(`✅ 成功导入 ${parsed.length} 条记录`);
                } catch (ex) { alert('❌ 解析失败，请检查文件格式'); }
            };
            reader.readAsBinaryString(file);
        }

        // ============ 150天示例数据 ============
        function loadMathSample() {
            if (records.length && !confirm('将替换为150天数学模拟数据，继续？')) return;
            const start = new Date();
            start.setDate(start.getDate() - 149);
            const sample = [];
            for (let i = 0; i < 150; i++) {
                const d = new Date(start);
                d.setDate(d.getDate() + i);
                const progress = i / 149;
                const baseScore = 45 + progress * 40 + Math.sin(progress * Math.PI * 10) * 15;
                const noise = (Math.random() - 0.5) * 25;
                const testScore = clamp(Math.round(baseScore + noise), 0, 100);
                sample.push({
                    subject: '数学',
                    date: d.toISOString().split('T')[0],
                    duration: 45 + Math.round(Math.random() * 40),
                    practice: clamp(testScore + Math.round((Math.random() - 0.5) * 12), 0, 100),
                    test: testScore,
                    review: clamp(testScore - 5 + Math.round((Math.random() - 0.5) * 15), 0, 100),
                    memory: clamp(testScore - 3 + Math.round((Math.random() - 0.5) * 10), 0, 100),
                    weight: 20
                });
            }
            records = sample;
            save(STORAGE_DATA, records);
            updateAll();
        }
        document.getElementById('saveSampleBtn').addEventListener('click', loadMathSample);

        // ============ 数据表格 ============
        function renderDataTable() {
            const t = document.getElementById('dataTable');
            if (!records.length) { t.innerHTML =
                '<tr><td colspan="9" style="padding:2rem;color:var(--text-muted);">📭 暂无数据，请加载示例或手动录入</td></tr>'; return; }
            const cols = ['subject', 'date', 'duration', 'practice', 'test', 'review', 'memory', 'weight'];
            const hds = ['学科', '日期', '时长', '练习', '测试', '复习', '记忆', '权重'];
            let html = '<thead><tr>' + hds.map(h => `<th>${h}</th>`).join('') + '<th>操作</th></tr></thead><tbody>';
            records.forEach((r, i) => {
                html +=
                    `<tr>${cols.map(c => `<td>${r[c]}</td>`).join('')}<td><button class="btn btn-danger btn-sm" onclick="deleteRecord(${i})">删</button></td></tr>`;
            });
            t.innerHTML = html + '</tbody>';
        }

        // ============ 仪表盘 ============
        function renderDashboard() {
            const math = records.filter(r => r.subject === '数学');
            if (!math.length) {
                document.getElementById('statsCards').innerHTML =
                    '<p style="grid-column:1/-1;text-align:center;color:var(--text-muted);padding:2rem;">📭 暂无数学数据</p>';
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
            const avgDuration = Math.round(totalMin / math.length);
            const overall = ((latest.practice + latest.test + latest.review + latest.memory) / 4).toFixed(1);

            document.getElementById('statsCards').innerHTML = [
                { v: overall, l: '综合评分' }, { v: latest.test + '%', l: '最新测试' },
                { v: Math.floor(totalMin / 60) + 'h ' + totalMin % 60 + 'm', l: '总时长' },
                { v: days + ' 天', l: '学习天数' }, { v: avgTest + '%', l: '平均测试' },
                { v: maxTest + '% / ' + minTest + '%', l: '最高 / 最低' },
                { v: avgDuration + ' 分钟', l: '日均时长' },
            ].map(s => `<div class="stat-card"><div class="stat-value">${s.v}</div><div class="stat-label">${s.l}</div></div>`)
                .join('');

            // 图表
            const chrono = [...math].sort((a, b) => new Date(a.date) - new Date(b.date)).slice(-100);
            const dates = chrono.map(r => r.date);
            const testVals = chrono.map(r => r.test);
            const durVals = chrono.map(r => r.duration);

            disposeChart('trend');
            const trendChart = echarts.init(document.getElementById('trendChart'));
            chartInstances.trend = trendChart;
            trendChart.setOption({
                ...darkTheme,
                tooltip: { trigger: 'axis' },
                legend: { bottom: 0, textStyle: { color: '#8e97b0', fontSize: 10 } },
                grid: { left: 48, right: 48, top: 15, bottom: 40 },
                xAxis: { type: 'category', data: dates, axisLabel: { rotate: 35, fontSize: 9, color: '#5c6480',
                        interval: 'auto' }, axisLine: { lineStyle: { color: 'rgba(255,255,255,0.06)' } } },
                yAxis: [
                    { type: 'value', name: '%', max: 100, axisLabel: { fontSize: 9, color: '#5c6480' },
                        splitLine: { lineStyle: { color: 'rgba(255,255,255,0.04)' } } },
                    { type: 'value', name: '分', axisLabel: { fontSize: 9, color: '#5c6480' },
                        splitLine: { show: false } }
                ],
                series: [
                    { name: '测试%', type: 'line', data: testVals, smooth: true, color: '#6d8aff',
                        symbolSize: 3, lineStyle: { width: 2 } },
                    { name: '时长', type: 'bar', yAxisIndex: 1, data: durVals, color: 'rgba(167,139,250,0.5)',
                        barWidth: '55%' }
                ]
            });

            const bins = [0, 40, 50, 60, 70, 80, 90, 101];
            const counts = bins.slice(0, -1).map((low, i) => math.filter(r => r.test >= low && r.test < bins[i + 1])
                .length);
            disposeChart('dist');
            const distChart = echarts.init(document.getElementById('distChart'));
            chartInstances.dist = distChart;
            distChart.setOption({
                ...darkTheme,
                tooltip: { trigger: 'axis' },
                grid: { left: 40, right: 20, top: 10, bottom: 30 },
                xAxis: { type: 'category', data: ['<40', '40-50', '50-60', '60-70', '70-80', '80-90', '90+'],
                    axisLabel: { fontSize: 9, color: '#5c6480' }, axisLine: { lineStyle: { color: 'rgba(255,255,255,0.06)' } } },
                yAxis: { type: 'value', name: '天', axisLabel: { fontSize: 9, color: '#5c6480' },
                    splitLine: { lineStyle: { color: 'rgba(255,255,255,0.04)' } } },
                series: [{ type: 'bar', data: counts, color: '#34d399', barWidth: '55%', itemStyle: { borderRadius: [6,
                        6, 0, 0
                    ] } }]
            });

            disposeChart('radar');
            const radarChart = echarts.init(document.getElementById('radarChart'));
            chartInstances.radar = radarChart;
            radarChart.setOption({
                ...darkTheme,
                radar: {
                    indicator: [{ name: '练习', max: 100 }, { name: '测试', max: 100 }, { name: '复习', max: 100 },
                    { name: '记忆', max: 100 }],
                    radius: '62%',
                    axisName: { color: '#8e97b0', fontSize: 10 },
                    splitArea: { areaStyle: { color: ['rgba(109,138,255,0.02)', 'rgba(109,138,255,0.04)'] } }
                },
                series: [{ type: 'radar', data: [{ name: '数学', value: [latest.practice, latest.test, latest.review,
                        latest.memory
                    ], areaStyle: { color: 'rgba(109,138,255,0.15)' }, lineStyle: { color: '#6d8aff',
                        width: 2 } }], symbolSize: 4 }]
            });

            disposeChart('scatter');
            const scatterChart = echarts.init(document.getElementById('scatterChart'));
            chartInstances.scatter = scatterChart;
            scatterChart.setOption({
                ...darkTheme,
                grid: { left: 50, right: 25, top: 15, bottom: 30 },
                xAxis: { type: 'value', name: '练习%', max: 100, axisLabel: { fontSize: 9, color: '#5c6480' },
                    splitLine: { lineStyle: { color: 'rgba(255,255,255,0.04)' } } },
                yAxis: { type: 'value', name: '测试%', max: 100, axisLabel: { fontSize: 9, color: '#5c6480' },
                    splitLine: { lineStyle: { color: 'rgba(255,255,255,0.04)' } } },
                series: [{ type: 'scatter', data: math.map(r => [r.practice, r.test]), symbolSize: 5, color: '#fbbf24',
                    itemStyle: { borderColor: 'rgba(0,0,0,0.3)', borderWidth: 1 } }]
            });

            const alerts = [];
            if (latest.test < 60) alerts.push(`测试 ${latest.test}%`);
            if (latest.review < 50) alerts.push(`复习 ${latest.review}%`);
            if (latest.memory < 50) alerts.push(`记忆 ${latest.memory}%`);
            document.getElementById('alertContainer').innerHTML =
                `<div class="card-header"><h3>⚠️ 预警</h3></div>${alerts.length ? alerts.map(a => `<span class="alert-badge">${a}</span>`).join(' ') : '<span style="color:var(--success);font-size:0.8rem;">✅ 一切正常，继续保持！</span>'}`;
        }

        // ============ 激励系统 ============
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
            const todayRecs = records.filter(r => r.date === today);
            if (todayRecs.length && today !== motivation.lastCalcDate) {
                motivation.xp += Math.round(todayRecs.reduce((s, r) => s + (r.test / 100) * 10 + r.duration * 0.5, 0));
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
            <div class="stat-card" style="background:${motivation.achievements[b.id]?'rgba(52,211,153,0.12)':'rgba(255,255,255,0.02)'};border-color:${motivation.achievements[b.id]?'rgba(52,211,153,0.3)':'var(--border-subtle)'};">
              <div style="font-size:2rem;margin-bottom:0.2rem;">${b.icon}</div>
              <div class="stat-label">${b.label}</div>
              <div style="font-size:0.65rem;margin-top:0.3rem;color:${motivation.achievements[b.id]?'var(--success)':'var(--text-muted)'};">${motivation.achievements[b.id]?'✅ 已解锁':'🔒 未解锁'}</div>
            </div>`).join('');
            document.getElementById('penaltyLog').innerHTML = motivation.penaltyLog.length ? motivation.penaltyLog.map(
                p =>
                `<div style="padding:0.4rem 0;border-bottom:1px solid var(--border-subtle);">📅 ${p.date} — ${p.reason}</div>`
                ).join('') : '<p style="color:var(--text-muted);padding:0.5rem 0;">🎉 暂无惩罚记录，继续保持！</p>';
        }

        // ============ 预测中心 ============
        function detectAnomalies(scores, dates) {
            const mean = scores.reduce((a, b) => a + b, 0) / scores.length;
            const variance = scores.reduce((s, v) => s + (v - mean) ** 2, 0) / scores.length;
            const std = Math.sqrt(variance);
            const threshold = 2 * std;
            return scores.reduce((arr, v, i) => {
                if (Math.abs(v - mean) > threshold) arr.push({ date: dates[i], score: v, deviation: ((v - mean) /
                    std).toFixed(2) });
                return arr;
            }, []);
        }

        function getRecentAll(recs) {
            const dateSet = [...new Set(recs.map(r => r.date))].sort();
            return recs.filter(r => dateSet.includes(r.date)).sort((a, b) => new Date(a.date) - new Date(b.date));
        }

        function renderHistory() {
            const recs = records.filter(r => r.subject === '数学');
            const recent = getRecentAll(recs).slice(-100);
            const chartDom = document.getElementById('historyChart'),
                tableDom = document.getElementById('historyTable');
            if (recent.length < 2) { chartDom.innerHTML =
                '<p style="text-align:center;padding:3rem;color:var(--text-muted);">数据不足（需≥2条）</p>';
                tableDom.innerHTML = ''; return; }
            const dates = recent.map(r => r.date),
                tests = recent.map(r => r.test);
            const anomalies = detectAnomalies(tests, dates);
            window._lastAnomalies = anomalies;
            disposeChart('history');
            const chart = echarts.init(chartDom);
            chartInstances.history = chart;
            chart.setOption({
                ...darkTheme,
                tooltip: { trigger: 'axis' },
                grid: { left: 45, right: 30, top: 15, bottom: 35 },
                xAxis: { type: 'category', data: dates, axisLabel: { rotate: 45, fontSize: 9, color: '#5c6480',
                        interval: 'auto' }, axisLine: { lineStyle: { color: 'rgba(255,255,255,0.06)' } } },
                yAxis: { max: 100, axisLabel: { fontSize: 9, color: '#5c6480' },
                    splitLine: { lineStyle: { color: 'rgba(255,255,255,0.04)' } } },
                series: [{
                    name: '测试%',
                    type: 'line',
                    data: tests,
                    smooth: true,
                    color: '#6d8aff',
                    symbolSize: 3,
                    lineStyle: { width: 2 },
                    areaStyle: { color: 'rgba(109,138,255,0.08)' },
                    markPoint: {
                        data: anomalies.map(a => ({
                            name: a.date,
                            coord: [a.date, a.score],
                            value: a.score,
                            symbol: 'pin',
                            symbolSize: 35,
                            itemStyle: { color: '#f87171' },
                            label: { formatter: '{c}', fontSize: 8, color: '#fff' }
                        }))
                    }
                }]
            });
            let html =
                '<thead><tr><th>日期</th><th>测试%</th><th>练习%</th><th>复习%</th><th>记忆%</th><th>状态</th></tr></thead><tbody>';
            recent.forEach(r => {
                const isAnomaly = anomalies.some(a => a.date === r.date);
                html +=
                    `<tr><td>${r.date}</td><td>${r.test}${isAnomaly?' <span class="anomaly-badge">异常</span>':''}</td><td>${r.practice}</td><td>${r.review}</td><td>${r.memory}</td><td>${isAnomaly?'⚠️':''}</td></tr>`;
            });
            tableDom.innerHTML = html + '</tbody>';
        }

        function renderAnomalies() {
            const table = document.getElementById('anomalyTable');
            const anomalies = window._lastAnomalies || [];
            if (!anomalies.length) {
                table.innerHTML =
                    '<thead><tr><th>日期</th><th>测试%</th><th>偏离 σ</th></tr></thead><tbody><tr><td colspan="3">✅ 暂无异常</td></tr></tbody>';
                return;
            }
            let html = '<thead><tr><th>日期</th><th>测试%</th><th>偏离 σ 倍数</th></tr></thead><tbody>';
            anomalies.forEach(a => { html +=
                    `<tr><td>${a.date}</td><td>${a.score}%</td><td><span class="anomaly-badge">${a.deviation}σ</span></td></tr>`; });
            table.innerHTML = html + '</tbody>';
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
            const baseline = days.map(d => clamp(intercept + slope * (n - 1 + d), 0, 100)).map(v => +v.toFixed(1));
            baseline[0] = +current.toFixed(1);
            const { target, sMid, sK } = forecastSettings;
            const uplift = days.map(d => d === 0 ? current : clamp(current + (target - current) / (1 + Math.exp(-sK * (
                d - sMid))), 0, 100)).map(v => +v.toFixed(1));
            const { halfLife, minScore } = forecastSettings;
            const decayRate = Math.log(2) / halfLife;
            const downside = days.map(d => d === 0 ? current : Math.max(minScore, minScore + (current - minScore) * Math
                .exp(-decayRate * d))).map(v => +v.toFixed(1));
            document.getElementById('forecastInfo').innerHTML =
                `📊 当前测试正确率 <strong>${current.toFixed(1)}%</strong> → 目标 <strong>${target}%</strong>`;
            const labels = ['今天', '+5d', '+10d', '+15d', '+20d', '+25d', '+30d'];
            document.getElementById('forecastTable').innerHTML =
                '<thead><tr><th>类别</th>' + labels.map(l => `<th>${l}</th>`).join('') + '</tr></thead><tbody>' +
                `<tr style="background:rgba(109,138,255,0.08);"><td>📈 基准线</td>${baseline.map(v => `<td>${v}%</td>`).join('')}</tr>` +
                `<tr style="background:rgba(52,211,153,0.08);"><td>🚀 S型提升</td>${uplift.map(v => `<td>${v}%</td>`).join('')}</tr>` +
                `<tr style="background:rgba(248,113,113,0.08);"><td>📉 遗忘底线</td>${downside.map(v => `<td>${v}%</td>`).join('')}</tr></tbody>`;
            disposeChart('forecast');
            const chart = echarts.init(document.getElementById('forecastChart'));
            chartInstances.forecast = chart;
            chart.setOption({
                ...darkTheme,
                legend: { bottom: 0, textStyle: { color: '#8e97b0', fontSize: 10 } },
                grid: { left: 45, right: 25, top: 15, bottom: 40 },
                xAxis: { type: 'category', data: labels, axisLabel: { fontSize: 9, color: '#5c6480' },
                    axisLine: { lineStyle: { color: 'rgba(255,255,255,0.06)' } } },
                yAxis: { max: 100, min: 0, axisLabel: { fontSize: 9, color: '#5c6480' },
                    splitLine: { lineStyle: { color: 'rgba(255,255,255,0.04)' } } },
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

        // ============ 预测参数折叠 ============
        document.getElementById('settingsToggle').addEventListener('click', () => {
            const body = document.getElementById('settingsBody');
            const arrow = document.getElementById('settingsArrow');
            if (body.style.display === 'none') { body.style.display = 'block';
                arrow.textContent = '▼'; } else { body.style.display = 'none';
                arrow.textContent = '▶'; }
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

        // ============ 学习记录 ============
        document.getElementById('addRecordNoteBtn').addEventListener('click', () => {
            const date = document.getElementById('recordDate').value;
            const text = document.getElementById('recordText').value.trim();
            const imageInput = document.getElementById('recordImage');
            if (!date) return alert('请选择日期');
            if (!text && !imageInput.files[0]) return alert('请至少填写文字或上传图片');
            const done = (img = '') => {
                studyNotes.unshift({ date, text, imageBase64: img });
                save(STORAGE_NOTES, studyNotes);
                renderNotes();
                document.getElementById('recordDate').valueAsDate = new Date();
                document.getElementById('recordText').value = '';
                imageInput.value = '';
            };
            if (imageInput.files[0]) {
                const r = new FileReader();
                r.onload = e => done(e.target.result);
                r.readAsDataURL(imageInput.files[0]);
            } else done();
        });

        function renderNotes() {
            const c = document.getElementById('recordNotesContainer');
            if (!studyNotes.length) { c.innerHTML =
                '<p style="text-align:center;color:var(--text-muted);padding:2rem;">📝 暂无学习记录，点击上方添加第一条吧</p>'; return; }
            c.innerHTML = studyNotes.map(n => `
            <div class="record-entry">
              <div style="font-weight:650;color:var(--accent);">${n.date}</div>
              <div style="white-space:pre-wrap;margin:0.4rem 0;">${n.text || '(无文字)'}</div>
              ${n.imageBase64 ? `<img src="${n.imageBase64}" alt="学习照片" loading="lazy">` : ''}
            </div>`).join('');
        }

        // ============ 全局更新 ============
        function updateAll() {
            renderDataTable();
            updateMotivation();
            if (document.getElementById('view-dashboard').classList.contains('active')) renderDashboard();
            if (document.getElementById('view-predict').classList.contains('active')) { renderHistory();
                renderForecast();
                renderAnomalies(); }
            if (document.getElementById('view-records').classList.contains('active')) renderNotes();
            if (document.getElementById('view-motivation').classList.contains('active')) renderMotivation();
        }

        // ============ 窗口大小变化时重绘图表 ============
        window.addEventListener('resize', () => {
            Object.values(chartInstances).forEach(c => { try { c.resize(); } catch (e) {} });
        });

        // ============ 启动 ============
        loadAll();
        if (!records.length) loadMathSample();
        else updateAll();
        loadSettingsUI();
        document.getElementById('recordDate').valueAsDate = new Date();
    })();
</script>
</body>
</html>
