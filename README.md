[自学量化平台-数学专属.html](https://github.com/user-attachments/files/30872868/-.html)
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover">
  <title>自学量化平台 · 数学专属</title>
  <script src="https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js">
  </script>
  <script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js">
  </script>
  <style>
    :root {
      --primary: #4361ee;
      --primary-light: #4895ef;
      --success: #10b981;
      --warning: #f59e0b;
      --danger: #ef4444;
      --bg: #f1f5f9;
      --card: #ffffff;
      --text: #1e293b;
      --border: #e2e8f0;
      --radius: 12px;
      --shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
      --safe-bottom: env(safe-area-inset-bottom, 0px);
    }
    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
      font-family: 'Segoe UI', Roboto, 'PingFang SC', 'Noto Sans SC', sans-serif;
    }
    body {
      background: var(--bg);
      color: var(--text);
      display: flex;
      flex-direction: column;
      height: 100vh;
      height: 100dvh;
      overflow: hidden;
      -webkit-tap-highlight-color: transparent;
      -webkit-font-smoothing: antialiased;
    }

    /* ========== 顶部状态栏 ========== */
    .top-bar {
      background: var(--card);
      padding: 0.6rem 1rem;
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-size: 0.8rem;
      box-shadow: 0 1px 4px rgba(0, 0, 0, 0.04);
      flex-shrink: 0;
      gap: 0.4rem;
      flex-wrap: wrap;
    }
    .top-bar span {
      white-space: nowrap;
    }
    .top-bar strong {
      color: var(--primary);
      font-size: 0.9rem;
    }

    /* ========== 主内容区 ========== */
    .main {
      flex: 1;
      overflow-y: auto;
      overflow-x: hidden;
      padding: 0.8rem;
      padding-bottom: calc(4.5rem + var(--safe-bottom));
      -webkit-overflow-scrolling: touch;
    }

    .view {
      display: none;
      flex-direction: column;
      gap: 0.8rem;
    }
    .view.active {
      display: flex;
    }

    /* ========== 底部导航 ========== */
    .bottom-nav {
      position: fixed;
      bottom: 0;
      left: 0;
      right: 0;
      background: var(--card);
      display: flex;
      border-top: 1px solid var(--border);
      padding-bottom: var(--safe-bottom);
      z-index: 100;
      box-shadow: 0 -2px 10px rgba(0, 0, 0, 0.06);
    }
    .nav-item {
      flex: 1;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: 0.4rem 0.2rem;
      cursor: pointer;
      color: #888;
      font-size: 0.65rem;
      transition: 0.15s;
      gap: 0.15rem;
      border: none;
      background: none;
      -webkit-tap-highlight-color: transparent;
      user-select: none;
      position: relative;
    }
    .nav-item .nav-icon {
      font-size: 1.2rem;
      transition: 0.15s;
    }
    .nav-item.active {
      color: var(--primary);
      font-weight: 700;
    }
    .nav-item.active::before {
      content: '';
      position: absolute;
      top: 0;
      left: 20%;
      right: 20%;
      height: 3px;
      background: var(--primary);
      border-radius: 0 0 3px 3px;
    }

    /* ========== 通用组件 ========== */
    .card {
      background: var(--card);
      border-radius: var(--radius);
      padding: 0.9rem;
      box-shadow: var(--shadow);
    }
    .card h3,
    .card h4 {
      font-size: 0.9rem;
      margin-bottom: 0.6rem;
      color: #333;
    }
    .header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 0.4rem;
    }
    .header h2 {
      font-size: 1rem;
    }

    .btn {
      border: none;
      padding: 0.5rem 0.9rem;
      border-radius: 8px;
      font-weight: 600;
      cursor: pointer;
      font-size: 0.78rem;
      transition: 0.15s;
      white-space: nowrap;
      -webkit-tap-highlight-color: transparent;
      user-select: none;
      display: inline-flex;
      align-items: center;
      gap: 0.25rem;
    }
    .btn:active {
      transform: scale(0.96);
      opacity: 0.85;
    }
    .btn-primary {
      background: var(--primary);
      color: #fff;
    }
    .btn-danger {
      background: var(--danger);
      color: #fff;
      font-size: 0.7rem;
      padding: 0.3rem 0.6rem;
    }
    .btn-sm {
      font-size: 0.65rem;
      padding: 0.25rem 0.5rem;
    }

    /* ========== 统计卡片网格 ========== */
    .stats-grid {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 0.5rem;
    }
    .stat-card {
      background: var(--card);
      border-radius: var(--radius);
      padding: 0.7rem;
      box-shadow: var(--shadow);
      text-align: center;
    }
    .stat-value {
      font-size: 1.3rem;
      font-weight: 700;
      color: var(--primary);
      word-break: break-all;
    }
    .stat-label {
      font-size: 0.68rem;
      color: #777;
      margin-top: 0.15rem;
    }

    /* ========== 图表容器 ========== */
    .chart-container {
      width: 100%;
      height: 220px;
    }

    /* ========== 双列网格 ========== */
    .grid-2 {
      display: flex;
      flex-direction: column;
      gap: 0.8rem;
    }

    /* ========== 表单 ========== */
    .form-row {
      display: flex;
      flex-direction: column;
      gap: 0.5rem;
    }
    .form-group {
      display: flex;
      flex-direction: column;
    }
    .form-group label {
      font-size: 0.7rem;
      margin-bottom: 0.15rem;
      color: #666;
    }
    .form-group input,
    .form-group textarea {
      padding: 0.55rem 0.6rem;
      border: 1.5px solid var(--border);
      border-radius: 8px;
      font-size: 0.85rem;
      width: 100%;
      background: #fafbfc;
      transition: 0.15s;
      -webkit-appearance: none;
    }
    .form-group input:focus,
    .form-group textarea:focus {
      border-color: var(--primary);
      outline: none;
      background: #fff;
    }
    .form-group textarea {
      resize: vertical;
      min-height: 60px;
    }

    /* ========== 上传区 ========== */
    .upload-area {
      border: 2px dashed var(--border);
      border-radius: var(--radius);
      padding: 1.2rem;
      text-align: center;
      cursor: pointer;
      font-size: 0.8rem;
      color: #888;
      transition: 0.15s;
    }
    .upload-area:active {
      border-color: var(--primary);
      background: #fafbff;
    }

    /* ========== 表格 ========== */
    .table-wrapper {
      max-height: 300px;
      overflow: auto;
      border: 1px solid var(--border);
      border-radius: var(--radius);
      -webkit-overflow-scrolling: touch;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 0.7rem;
      white-space: nowrap;
    }
    th,
    td {
      padding: 0.4rem 0.5rem;
      border-bottom: 1px solid var(--border);
      text-align: center;
    }
    th {
      background: #f1f5f9;
      position: sticky;
      top: 0;
      z-index: 1;
      font-size: 0.68rem;
    }

    /* ========== 标签/徽章 ========== */
    .alert-badge {
      padding: 0.2rem 0.6rem;
      border-radius: 20px;
      font-size: 0.7rem;
      font-weight: 600;
      background: #fee2e2;
      color: var(--danger);
      display: inline-block;
      margin: 0.15rem;
    }
    .anomaly-badge {
      background: #ffe5e5;
      color: #b91c1c;
      padding: 0.15rem 0.45rem;
      border-radius: 20px;
      font-weight: 700;
      font-size: 0.65rem;
      white-space: nowrap;
    }
    .info-row {
      padding: 0.5rem 0.7rem;
      background: #f8f9ff;
      border-radius: 8px;
      font-size: 0.75rem;
      margin-bottom: 0.5rem;
    }

    /* ========== 学习记录卡片 ========== */
    .record-entry {
      background: #f8fafd;
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 0.7rem;
      margin-bottom: 0.5rem;
      font-size: 0.8rem;
    }
    .record-entry img {
      max-width: 100%;
      max-height: 180px;
      margin: 0.4rem 0;
      border-radius: 8px;
      object-fit: contain;
    }

    /* ========== 折叠面板 ========== */
    .collapse-header {
      cursor: pointer;
      padding: 0.5rem 0;
      font-size: 0.85rem;
      display: flex;
      justify-content: space-between;
      align-items: center;
      user-select: none;
      -webkit-tap-highlight-color: transparent;
    }

    /* ========== 平板及以上 ========== */
    @media (min-width: 768px) {
      body {
        flex-direction: row;
      }
      .top-bar {
        display: none;
      }
      .bottom-nav {
        position: static;
        flex-direction: column;
        width: 200px;
        border-top: none;
        border-right: 1px solid var(--border);
        padding-bottom: 0;
        box-shadow: none;
        flex-shrink: 0;
        gap: 0;
      }
      .nav-item {
        flex-direction: row;
        padding: 0.7rem 1.2rem;
        font-size: 0.85rem;
        gap: 0.5rem;
        justify-content: flex-start;
        border-left: 3px solid transparent;
      }
      .nav-item.active::before {
        display: none;
      }
      .nav-item.active {
        background: #f0f4ff;
        border-left-color: var(--primary);
      }
      .nav-item .nav-icon {
        font-size: 1rem;
      }
      .main {
        padding: 1.2rem;
        padding-bottom: 1.2rem;
      }
      .grid-2 {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(340px, 1fr));
        gap: 1rem;
      }
      .stats-grid {
        grid-template-columns: repeat(auto-fill, minmax(140px, 1fr));
        gap: 0.7rem;
      }
      .chart-container {
        height: 260px;
      }
      .form-row {
        flex-direction: row;
        flex-wrap: wrap;
        align-items: flex-end;
      }
      .form-group {
        min-width: 90px;
      }
      .card {
        padding: 1.2rem;
      }
      .stat-value {
        font-size: 1.6rem;
      }
      table {
        font-size: 0.8rem;
        white-space: normal;
      }
    }

    @media (min-width: 1024px) {
      .stats-grid {
        grid-template-columns: repeat(7, 1fr);
      }
      .chart-container {
        height: 300px;
      }
    }
  </style>
</head>
<body>

  <!-- 顶部状态栏（仅手机可见） -->
  <div class="top-bar">
    <span>🔥 连续 <strong id="streakDays">0</strong> 天</span>
    <span>⭐ 经验 <strong id="xpValue">0</strong></span>
    <span id="latestBadge" style="color:var(--warning);font-weight:600;"></span>
  </div>

  <!-- 主内容区 -->
  <main class="main">
    <!-- 数据管理 -->
    <section class="view active" id="view-data">
      <div class="header">
        <h2>📁 数据管理</h2>
        <div style="display:flex;gap:0.3rem;flex-wrap:wrap;">
          <button class="btn btn-primary btn-sm" id="saveSampleBtn">📥 加载150天示例</button>
          <button class="btn btn-danger btn-sm" id="deleteAllBtn">🗑️ 清空</button>
        </div>
      </div>
      <div class="card">
        <h3>📝 实时录入</h3>
        <div class="form-row">
          <div class="form-group"><label>学科</label><input type="text" id="inputSubject" value="数学" readonly></div>
          <div class="form-group"><label>日期</label><input type="date" id="inputDate"></div>
          <div class="form-group"><label>时长(分)</label><input type="number" id="inputDuration" value="60" min="0"></div>
          <div class="form-group"><label>练习%</label><input type="number" id="inputPractice" value="75" min="0" max="100"></div>
          <div class="form-group"><label>测试%</label><input type="number" id="inputTest" value="70" min="0" max="100"></div>
          <div class="form-group"><label>复习%</label><input type="number" id="inputReview" value="65" min="0" max="100"></div>
          <div class="form-group"><label>记忆%</label><input type="number" id="inputMemory" value="68" min="0" max="100"></div>
          <div class="form-group"><label>权重</label><input type="number" id="inputWeight" value="20" min="0"></div>
          <button class="btn btn-primary" id="addRecordBtn">✅ 添加记录</button>
        </div>
        <div class="upload-area" id="uploadArea" style="margin-top:0.6rem;">
          <p>📂 点击上传 Excel 文件</p>
          <input type="file" id="fileInput" accept=".xlsx,.xls,.csv" hidden>
        </div>
        <div class="table-wrapper"><table id="dataTable"><thead></thead><tbody></tbody></table></div>
      </div>
    </section>

    <!-- 量化仪表盘 -->
    <section class="view" id="view-dashboard">
      <div class="header"><h2>📈 学习量化仪表盘</h2></div>
      <div class="stats-grid" id="statsCards"></div>
      <div class="grid-2">
        <div class="card"><h4>📊 学习时长 &amp; 测试趋势</h4><div class="chart-container" id="trendChart"></div></div>
        <div class="card"><h4>📈 测试分数分布</h4><div class="chart-container" id="distChart"></div></div>
      </div>
      <div class="grid-2">
        <div class="card"><h4>🎯 各维度雷达图</h4><div class="chart-container" id="radarChart"></div></div>
        <div class="card"><h4>🔍 练习 vs 测试相关性</h4><div class="chart-container" id="scatterChart"></div></div>
      </div>
      <div class="card" id="alertContainer"></div>
    </section>

    <!-- 预测中心 -->
    <section class="view" id="view-predict">
      <div class="header"><h2>🔮 预测中心 · 数学</h2></div>
      <div class="card">
        <div class="collapse-header" id="settingsToggle"><strong>⚙️ 预测参数</strong><span id="settingsArrow">▼</span></div>
        <div id="settingsBody">
          <div class="form-row">
            <div class="form-group"><label>目标正确率%</label><input type="number" id="targetScore" value="90" min="30" max="100"></div>
            <div class="form-group"><label>S型中点(天)</label><input type="number" id="sMid" value="15" min="1"></div>
            <div class="form-group"><label>陡峭度 k</label><input type="number" id="sK" value="0.3" step="0.05"></div>
            <div class="form-group"><label>遗忘半衰期</label><input type="number" id="halfLife" value="12" min="1"></div>
            <div class="form-group"><label>底线分%</label><input type="number" id="minScore" value="40" min="0" max="80"></div>
            <button class="btn btn-primary btn-sm" id="savePredictSettingsBtn">💾 保存</button>
            <span id="settingsSavedHint" style="color:var(--success);display:none;font-size:0.8rem;">✅</span>
          </div>
        </div>
      </div>
      <div class="card">
        <h4>📜 测试走势 <span class="anomaly-badge">红标=异常</span></h4>
        <div class="chart-container" id="historyChart"></div>
        <div class="table-wrapper" style="margin-top:0.5rem;"><table id="historyTable"></table></div>
      </div>
      <div class="card">
        <h4>⚠️ 异常点记录</h4>
        <div class="table-wrapper"><table id="anomalyTable"><thead><tr><th>日期</th><th>测试%</th><th>偏离σ</th></tr></thead><tbody><tr><td colspan="3">暂无异常</td></tr></tbody></table></div>
      </div>
      <div class="card">
        <h4>🔮 未来30天预测</h4>
        <div class="info-row" id="forecastInfo"></div>
        <div class="table-wrapper"><table class="prediction-table" id="forecastTable"></table></div>
        <div class="chart-container" id="forecastChart" style="margin-top:0.5rem;"></div>
      </div>
    </section>

    <!-- 学习记录 -->
    <section class="view" id="view-records">
      <div class="header"><h2>📓 学习情况记录</h2></div>
      <div class="card">
        <h4>📝 添加新记录</h4>
        <div class="form-row">
          <div class="form-group"><label>日期</label><input type="date" id="recordDate"></div>
          <div class="form-group"><label>文字记录</label><textarea id="recordText" rows="2" placeholder="今天学了什么…"></textarea></div>
          <div class="form-group"><label>上传图片</label><input type="file" id="recordImage" accept="image/*"></div>
          <button class="btn btn-primary" id="addRecordNoteBtn">➕ 添加</button>
        </div>
      </div>
      <div class="card">
        <h4>📋 历史记录</h4>
        <div id="recordNotesContainer"></div>
      </div>
    </section>

    <!-- 激励中心 -->
    <section class="view" id="view-motivation">
      <div class="card">
        <h4>🏅 成就</h4>
        <div id="achievementList" class="stats-grid"></div>
      </div>
      <div class="card">
        <h4>📜 惩罚记录</h4>
        <div id="penaltyLog" style="font-size:0.78rem;"></div>
      </div>
    </section>
  </main>

  <!-- 底部导航（手机端） -->
  <nav class="bottom-nav">
    <div class="nav-item active" data-view="data">
      <span class="nav-icon">📁</span>数据
    </div>
    <div class="nav-item" data-view="dashboard">
      <span class="nav-icon">📈</span>仪表盘
    </div>
    <div class="nav-item" data-view="predict">
      <span class="nav-icon">🔮</span>预测
    </div>
    <div class="nav-item" data-view="records">
      <span class="nav-icon">📓</span>记录
    </div>
    <div class="nav-item" data-view="motivation">
      <span class="nav-icon">🏆</span>激励
    </div>
  </nav>

  <script>
    (function() {
      const STORAGE_DATA = 'sq_math_data_v3';
      const STORAGE_MOTIV = 'sq_math_motiv_v3';
      const STORAGE_SETTINGS = 'sq_math_settings_v3';
      const STORAGE_NOTES = 'sq_math_notes_v3';

      let records = [];
      let motivation = { xp: 0, streak: 0, achievements: {}, penaltyLog: [], lastCalcDate: '' };
      let forecastSettings = { target: 90, sMid: 15, sK: 0.3, halfLife: 12, minScore: 40 };
      let studyNotes = [];
      const chartInstances = {};

      function disposeChart(key) {
        if (chartInstances[key]) { chartInstances[key].dispose();
          delete chartInstances[key]; }
      }

      function saveData() { localStorage.setItem(STORAGE_DATA, JSON.stringify(records)); }

      function loadData() { const s = localStorage.getItem(STORAGE_DATA); if (s) try { records = JSON.parse(
          s); } catch (e) {} }

      function saveMot() { localStorage.setItem(STORAGE_MOTIV, JSON.stringify(motivation)); }

      function loadMot() { const s = localStorage.getItem(STORAGE_MOTIV); if (s) try { motivation = JSON.parse(
          s); } catch (e) {} }

      function saveSettings() { localStorage.setItem(STORAGE_SETTINGS, JSON.stringify(forecastSettings)); }

      function loadSettings() { const s = localStorage.getItem(STORAGE_SETTINGS); if (s) try { forecastSettings =
            JSON.parse(s); } catch (e) {} }

      function saveNotes() { localStorage.setItem(STORAGE_NOTES, JSON.stringify(studyNotes)); }

      function loadNotes() { const s = localStorage.getItem(STORAGE_NOTES); if (s) try { studyNotes = JSON.parse(
          s); } catch (e) {} }

      // 导航
      document.querySelectorAll('.nav-item').forEach(item => {
        item.addEventListener('click', () => {
          const vn = item.dataset.view;
          document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
          item.classList.add('active');
          document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
          document.getElementById(`view-${vn}`).classList.add('active');
          if (vn === 'data') renderDataTable();
          if (vn === 'dashboard') renderDashboard();
          if (vn === 'predict') { renderHistory();
            renderForecast();
            renderAnomalies(); }
          if (vn === 'records') renderNotes();
          if (vn === 'motivation') renderMotivation();
        });
      });

      document.getElementById('inputDate').valueAsDate = new Date();
      document.getElementById('addRecordBtn').addEventListener('click', () => {
        records.unshift({
          subject: '数学',
          date: document.getElementById('inputDate').value,
          duration: +document.getElementById('inputDuration').value,
          practice: +document.getElementById('inputPractice').value,
          test: +document.getElementById('inputTest').value,
          review: +document.getElementById('inputReview').value,
          memory: +document.getElementById('inputMemory').value,
          weight: +document.getElementById('inputWeight').value
        });
        saveData();
        updateAll();
      });
      window.deleteRecord = (idx) => { if (confirm('删除？')) { records.splice(idx, 1);
          saveData();
          updateAll(); } };
      document.getElementById('deleteAllBtn').addEventListener('click', () => { if (confirm('清空全部？')) { records = [];
          saveData();
          updateAll(); } });
      const uploadArea = document.getElementById('uploadArea'),
        fileInput = document.getElementById('fileInput');
      uploadArea.addEventListener('click', () => fileInput.click());
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
              obj.practice = +obj.practice || 0;
              obj.test = +obj.test || 0;
              obj.review = +obj.review || 0;
              obj.memory = +obj.memory || 0;
              obj.duration = +obj.duration || 0;
              obj.weight = +obj.weight || 20;
              obj.subject = obj.subject || '数学';
              return obj;
            });
            records = [...records, ...parsed];
            saveData();
            updateAll();
            alert(`导入${parsed.length}条`);
          } catch (ex) { alert('解析失败'); }
        };
        reader.readAsBinaryString(file);
      }

      function loadMathSample() {
        if (records.length && !confirm('将替换为150天数学数据，继续？')) return;
        const start = new Date();
        start.setDate(start.getDate() - 149);
        const sample = [];
        for (let i = 0; i < 150; i++) {
          const d = new Date(start);
          d.setDate(d.getDate() + i);
          const progress = i / 149;
          const baseScore = 45 + progress * 40 + Math.sin(progress * Math.PI * 10) * 15;
          const noise = (Math.random() - 0.5) * 25;
          let testScore = Math.round(Math.min(100, Math.max(0, baseScore + noise)));
          const practice = Math.min(100, testScore + Math.round((Math.random() - 0.5) * 12));
          const review = Math.min(100, testScore - 5 + Math.round((Math.random() - 0.5) * 15));
          const memory = Math.min(100, testScore - 3 + Math.round((Math.random() - 0.5) * 10));
          sample.push({
            subject: '数学',
            date: d.toISOString().split('T')[0],
            duration: 45 + Math.round(Math.random() * 40),
            practice: Math.max(0, practice),
            test: testScore,
            review: Math.max(0, review),
            memory: Math.max(0, memory),
            weight: 20
          });
        }
        records = sample;
        saveData();
        updateAll();
      }
      document.getElementById('saveSampleBtn').addEventListener('click', loadMathSample);

      function renderDataTable() {
        const t = document.getElementById('dataTable');
        if (!records.length) { t.innerHTML = '<tr><td colspan="9">暂无数据</td></tr>'; return; }
        const cols = ['subject', 'date', 'duration', 'practice', 'test', 'review', 'memory', 'weight'];
        const hds = ['学科', '日期', '时长', '练习', '测试', '复习', '记忆', '权重'];
        t.innerHTML = '<thead><tr>' + hds.map(h => `<th>${h}</th>`).join('') +
          '<th>操作</th></tr></thead><tbody>';
        records.forEach((r, i) => {
          t.innerHTML += `<tr>${cols.map(c => `<td>${r[c]}</td>`).join('')}<td><button class="btn btn-danger btn-sm" onclick="deleteRecord(${i})">删</button></td></tr>`;
        });
        t.innerHTML += '</tbody>';
      }

      function renderDashboard() {
        const math = records.filter(r => r.subject === '数学');
        if (!math.length) {
          document.getElementById('statsCards').innerHTML = '<p style="grid-column:1/-1;text-align:center;">暂无数据</p>';
          document.getElementById('alertContainer').innerHTML = '';
          return;
        }
        const latest = math.sort((a, b) => new Date(b.date) - new Date(a.date))[0];
        const totalMin = math.reduce((s, r) => s + r.duration, 0);
        const days = new Set(math.map(r => r.date)).size;
        const avgTest = (math.reduce((s, r) => s + r.test, 0) / math.length).toFixed(1);
        const maxTest = Math.max(...math.map(r => r.test));
        const minTest = Math.min(...math.map(r => r.test));
        const avgDuration = Math.round(totalMin / math.length);
        const overall = ((latest.practice + latest.test + latest.review + latest.memory) / 4).toFixed(1);

        document.getElementById('statsCards').innerHTML = `
          <div class="stat-card"><div class="stat-value">${overall}</div><div class="stat-label">综合评分</div></div>
          <div class="stat-card"><div class="stat-value">${latest.test}%</div><div class="stat-label">最新测试</div></div>
          <div class="stat-card"><div class="stat-value">${Math.floor(totalMin/60)}h${totalMin%60}m</div><div class="stat-label">总时长</div></div>
          <div class="stat-card"><div class="stat-value">${days}天</div><div class="stat-label">学习天数</div></div>
          <div class="stat-card"><div class="stat-value">${avgTest}%</div><div class="stat-label">平均测试</div></div>
          <div class="stat-card"><div class="stat-value">${maxTest}/${minTest}</div><div class="stat-label">最高/最低</div></div>
          <div class="stat-card"><div class="stat-value">${avgDuration}分</div><div class="stat-label">日均时长</div></div>
        `;

        const sorted = [...math].sort((a, b) => new Date(a.date) - new Date(b.date));
        const recent100 = sorted.slice(-100);
        const dates = recent100.map(r => r.date);
        const testVals = recent100.map(r => r.test);
        const durVals = recent100.map(r => r.duration);

        disposeChart('trend');
        const trendChart = echarts.init(document.getElementById('trendChart'));
        chartInstances.trend = trendChart;
        trendChart.setOption({
          tooltip: { trigger: 'axis' },
          legend: { bottom: 0, textStyle: { fontSize: 10 } },
          grid: { left: 40, right: 40, top: 10, bottom: 35 },
          xAxis: { type: 'category', data: dates, axisLabel: { rotate: 30, fontSize: 9, interval: 'auto' } },
          yAxis: [{ type: 'value', name: '%', max: 100, axisLabel: { fontSize: 9 } }, { type: 'value',
            name: '分',
            axisLabel: { fontSize: 9 } }],
          series: [
            { name: '测试%', type: 'line', data: testVals, smooth: true, color: '#4361ee',
            symbolSize: 4 },
            { name: '时长', type: 'bar', yAxisIndex: 1, data: durVals, color: '#a0c4ff',
            barWidth: '60%' }
          ]
        });

        const bins = [0, 40, 50, 60, 70, 80, 90, 101];
        const counts = bins.slice(0, -1).map((low, i) => math.filter(r => r.test >= low && r.test < bins[i + 1])
          .length);
        const binLabels = ['<40', '40-50', '50-60', '60-70', '70-80', '80-90', '90+'];
        disposeChart('dist');
        const distChart = echarts.init(document.getElementById('distChart'));
        chartInstances.dist = distChart;
        distChart.setOption({
          tooltip: { trigger: 'axis' },
          grid: { left: 40, right: 20, top: 10, bottom: 25 },
          xAxis: { type: 'category', data: binLabels, axisLabel: { fontSize: 9 } },
          yAxis: { type: 'value', name: '天', axisLabel: { fontSize: 9 } },
          series: [{ type: 'bar', data: counts, color: '#10b981', barWidth: '60%' }]
        });

        disposeChart('radar');
        const radarChart = echarts.init(document.getElementById('radarChart'));
        chartInstances.radar = radarChart;
        radarChart.setOption({
          radar: { indicator: [{ name: '练习', max: 100 }, { name: '测试', max: 100 }, { name: '复习',
              max: 100 }, { name: '记忆', max: 100 }], radius: '60%' },
          series: [{ type: 'radar', data: [{ name: '数学', value: [latest.practice, latest.test, latest
                .review, latest.memory
              ] }] }]
        });

        disposeChart('scatter');
        const scatterChart = echarts.init(document.getElementById('scatterChart'));
        chartInstances.scatter = scatterChart;
        const scatterData = math.map(r => [r.practice, r.test]);
        scatterChart.setOption({
          grid: { left: 45, right: 20, top: 10, bottom: 25 },
          xAxis: { type: 'value', name: '练习%', max: 100, axisLabel: { fontSize: 9 } },
          yAxis: { type: 'value', name: '测试%', max: 100, axisLabel: { fontSize: 9 } },
          series: [{ type: 'scatter', data: scatterData, symbolSize: 5, color: '#f59e0b' }]
        });

        const alerts = [];
        if (latest.test < 60) alerts.push(`测试${latest.test}%`);
        if (latest.review < 50) alerts.push(`复习${latest.review}%`);
        if (latest.memory < 50) alerts.push(`记忆${latest.memory}%`);
        document.getElementById('alertContainer').innerHTML =
          `<h4>⚠️ 预警</h4>${alerts.length ? alerts.map(a => `<span class="alert-badge">${a}</span>`).join(' ') : '<span style="color:var(--success);font-size:0.8rem;">✅ 一切正常</span>'}`;
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
        const todayRecs = records.filter(r => r.date === today);
        if (todayRecs.length && today !== motivation.lastCalcDate) {
          const xp = Math.round(todayRecs.reduce((s, r) => s + (r.test / 100) * 10 + r.duration * 0.5, 0));
          motivation.xp += xp;
          motivation.lastCalcDate = today;
        }
        if (streak >= 7) motivation.achievements.streak7 = true;
        if (records.reduce((s, r) => s + r.duration, 0) > 3000) motivation.achievements.total3k = true;
        const lastDate = dates.length ? new Date(dates[dates.length - 1]) : new Date(0);
        const diff = Math.floor((new Date() - lastDate) / 86400000);
        if (diff >= 3 && (!motivation.lastPenaltyDate || motivation.lastPenaltyDate !== today)) {
          motivation.xp = Math.max(0, motivation.xp - 10);
          motivation.penaltyLog.unshift({ date: today, reason: `连续${diff}天未学习，扣除10经验` });
          motivation.lastPenaltyDate = today;
        }
        saveMot();
        document.getElementById('streakDays').textContent = streak;
        document.getElementById('xpValue').textContent = motivation.xp;
        document.getElementById('latestBadge').textContent = motivation.achievements.streak7 ? '🔥连续7天' : '';
      }

      function renderMotivation() {
        const badges = [{ id: 'streak7', icon: '🔥', label: '连续7天' }, { id: 'total3k', icon: '⏱️',
          label: '学习超3000分钟' }];
        document.getElementById('achievementList').innerHTML = badges.map(b =>
          `<div class="stat-card" style="background:${motivation.achievements[b.id]?'#d1fae5':'#f3f4f6'}"><div style="font-size:1.5rem;">${b.icon}</div><div class="stat-label">${b.label}</div></div>`
        ).join('');
        document.getElementById('penaltyLog').innerHTML = motivation.penaltyLog.length ? motivation.penaltyLog.map(
          p => `<div style="padding:0.3rem 0;border-bottom:1px solid var(--border);">📅 ${p.date} - ${p.reason}</div>`
        ).join('') : '🎉 暂无惩罚';
      }

      function detectAnomalies(scores, dates) {
        const mean = scores.reduce((a, b) => a + b, 0) / scores.length;
        const variance = scores.reduce((s, v) => s + (v - mean) ** 2, 0) / scores.length;
        const std = Math.sqrt(variance);
        const threshold = 2 * std;
        const anomalies = [];
        scores.forEach((v, i) => {
          if (Math.abs(v - mean) > threshold) {
            anomalies.push({ date: dates[i], score: v, deviation: ((v - mean) / std).toFixed(2) });
          }
        });
        return anomalies;
      }

      function getRecentAll(recs) {
        const dates = [...new Set(recs.map(r => r.date))].sort();
        return recs.filter(r => dates.includes(r.date)).sort((a, b) => new Date(a.date) - new Date(b.date));
      }

      function renderHistory() {
        const recs = records.filter(r => r.subject === '数学');
        const recent = getRecentAll(recs).slice(-100);
        const chartDom = document.getElementById('historyChart'),
          tableDom = document.getElementById('historyTable');
        if (recent.length < 2) { chartDom.innerHTML = '<p style="text-align:center;padding:2rem;">数据不足</p>';
          tableDom.innerHTML = ''; return; }
        const dates = recent.map(r => r.date),
          tests = recent.map(r => r.test);
        const anomalies = detectAnomalies(tests, dates);
        disposeChart('history');
        const chart = echarts.init(chartDom);
        chartInstances.history = chart;
        chart.setOption({
          tooltip: { trigger: 'axis' },
          grid: { left: 40, right: 30, top: 15, bottom: 30 },
          xAxis: { type: 'category', data: dates, axisLabel: { rotate: 45, fontSize: 9, interval: 'auto' } },
          yAxis: { max: 100, axisLabel: { fontSize: 9 } },
          series: [{
            name: '测试%',
            type: 'line',
            data: tests,
            smooth: true,
            color: '#4361ee',
            symbolSize: 3,
            markPoint: {
              data: anomalies.map(a => ({
                name: a.date,
                coord: [a.date, a.score],
                value: a.score,
                symbol: 'pin',
                symbolSize: 30,
                itemStyle: { color: '#ef4444' },
                label: { formatter: '{c}', fontSize: 8 }
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
        window._lastAnomalies = anomalies;
      }

      function renderAnomalies() {
        const table = document.getElementById('anomalyTable');
        const anomalies = window._lastAnomalies || [];
        if (!anomalies.length) {
          table.innerHTML =
            '<thead><tr><th>日期</th><th>测试%</th><th>偏离σ</th></tr></thead><tbody><tr><td colspan="3">暂无异常</td></tr></tbody>';
          return;
        }
        let html = '<thead><tr><th>日期</th><th>测试%</th><th>偏离σ倍数</th></tr></thead><tbody>';
        anomalies.forEach(a => { html +=
            `<tr><td>${a.date}</td><td>${a.score}%</td><td>${a.deviation}σ</td></tr>`; });
        table.innerHTML = html + '</tbody>';
      }

      function renderForecast() {
        const recs = records.filter(r => r.subject === '数学');
        const recent = getRecentAll(recs).slice(-30);
        if (recent.length < 3) return;
        const scores = recent.map(r => r.test);
        const n = scores.length;
        const current = scores[n - 1];
        const sumX = scores.reduce((s, _, i) => s + i, 0),
          sumY = scores.reduce((s, v) => s + v, 0);
        const sumXY = scores.reduce((s, v, i) => s + i * v, 0),
          sumX2 = scores.reduce((s, _, i) => s + i * i, 0);
        const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
        const intercept = (sumY - slope * sumX) / n;
        const days = [0, 5, 10, 15, 20, 25, 30];
        const baseline = days.map(d => Math.min(100, Math.max(0, intercept + slope * (n - 1 + d)))).map(v => +v.toFixed(
          1));
        baseline[0] = +current.toFixed(1);
        const { target, sMid, sK } = forecastSettings;
        const uplift = days.map(d => {
          if (d === 0) return current;
          const growth = target - current;
          const logistic = 1 / (1 + Math.exp(-sK * (d - sMid)));
          return Math.min(100, current + growth * logistic);
        }).map(v => +v.toFixed(1));
        const { halfLife, minScore } = forecastSettings;
        const decayRate = Math.log(2) / halfLife;
        const downside = days.map(d => {
          if (d === 0) return current;
          const dropped = minScore + (current - minScore) * Math.exp(-decayRate * d);
          return Math.max(minScore, dropped);
        }).map(v => +v.toFixed(1));
        document.getElementById('forecastInfo').innerHTML =
          `📊 当前测试正确率 <strong>${current.toFixed(1)}%</strong> ，目标 <strong>${target}%</strong>`;
        const labels = ['今天', '+5d', '+10d', '+15d', '+20d', '+25d', '+30d'];
        document.getElementById('forecastTable').innerHTML =
          '<thead><tr><th>类别</th>' + labels.map(l => `<th>${l}</th>`).join('') + '</tr></thead><tbody>' +
          `<tr style="background:#eef2ff"><td>基准线</td>${baseline.map(v => `<td>${v}%</td>`).join('')}</tr>` +
          `<tr style="background:#d1fae5"><td>S型提升</td>${uplift.map(v => `<td>${v}%</td>`).join('')}</tr>` +
          `<tr style="background:#fee2e2"><td>遗忘底线</td>${downside.map(v => `<td>${v}%</td>`).join('')}</tr></tbody>`;
        disposeChart('forecast');
        const chart = echarts.init(document.getElementById('forecastChart'));
        chartInstances.forecast = chart;
        chart.setOption({
          legend: { bottom: 0, textStyle: { fontSize: 10 } },
          grid: { left: 40, right: 20, top: 10, bottom: 35 },
          xAxis: { type: 'category', data: labels, axisLabel: { fontSize: 9 } },
          yAxis: { max: 100, min: 0, axisLabel: { fontSize: 9 } },
          series: [
            { name: '基准线', type: 'line', data: baseline, smooth: true, color: '#4361ee',
              symbolSize: 5 },
            { name: 'S型提升', type: 'line', data: uplift, smooth: true, lineStyle: { type: 'dashed',
                color: '#10b981' }, symbol: 'diamond', symbolSize: 5 },
            { name: '遗忘底线', type: 'line', data: downside, smooth: true, lineStyle: { type: 'dotted',
                color: '#ef4444' }, symbol: 'triangle', symbolSize: 5 }
          ]
        });
      }

      document.getElementById('settingsToggle').addEventListener('click', () => {
        const body = document.getElementById('settingsBody');
        const arrow = document.getElementById('settingsArrow');
        if (body.style.display === 'none') { body.style.display = 'block';
          arrow.textContent = '▼'; } else { body.style.display = 'none';
          arrow.textContent = '▶'; }
      });
      document.getElementById('savePredictSettingsBtn').addEventListener('click', () => {
        forecastSettings.target = +document.getElementById('targetScore').value || 90;
        forecastSettings.sMid = +document.getElementById('sMid').value || 15;
        forecastSettings.sK = +document.getElementById('sK').value || 0.3;
        forecastSettings.halfLife = +document.getElementById('halfLife').value || 12;
        forecastSettings.minScore = +document.getElementById('minScore').value || 40;
        saveSettings();
        document.getElementById('settingsSavedHint').style.display = 'inline';
        setTimeout(() => document.getElementById('settingsSavedHint').style.display = 'none', 2000);
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
        const imageInput = document.getElementById('recordImage');
        if (!date) return alert('请选择日期');
        if (!text && !imageInput.files[0]) return alert('至少填写文字或上传图片');
        const processAndSave = (imageBase64 = '') => {
          studyNotes.unshift({ date, text, imageBase64 });
          saveNotes();
          renderNotes();
          document.getElementById('recordDate').valueAsDate = new Date();
          document.getElementById('recordText').value = '';
          imageInput.value = '';
        };
        if (imageInput.files[0]) {
          const reader = new FileReader();
          reader.onload = e => processAndSave(e.target.result);
          reader.readAsDataURL(imageInput.files[0]);
        } else processAndSave();
      });

      function renderNotes() {
        const container = document.getElementById('recordNotesContainer');
        if (!studyNotes.length) { container.innerHTML =
          '<p style="text-align:center;color:#999;padding:1rem;">暂无学习记录，请添加。</p>'; return; }
        container.innerHTML = studyNotes.map(note => `
          <div class="record-entry">
            <div><strong>${note.date}</strong></div>
            <div style="white-space:pre-wrap;margin:0.4rem 0;">${note.text || '(无文字)'}</div>
            ${note.imageBase64 ? `<img src="${note.imageBase64}" alt="学习照片" loading="lazy">` : ''}
          </div>
        `).join('');
      }

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

      loadData();
      loadMot();
      loadSettings();
      loadNotes();
      if (!records.length) loadMathSample();
      else updateAll();
      loadSettingsUI();
      document.getElementById('recordDate').valueAsDate = new Date();
    })();
  </script>
</body>
</html>
