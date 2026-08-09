[自学量化平台-数学专属-丰富仪表盘.html](https://github.com/user-attachments/files/30872608/-.-.html)
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes">
  <title>自学量化平台 · 数学专属 (丰富仪表盘)</title>
  <script src="https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js"></script>
  <script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
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
      --radius: 14px;
      --shadow: 0 4px 6px rgba(0,0,0,0.04);
    }
    * { margin:0; padding:0; box-sizing:border-box; font-family:'Segoe UI', Roboto, 'PingFang SC', sans-serif; }
    body { background:var(--bg); color:var(--text); display:flex; height:100vh; overflow:hidden; }
    .sidebar { width:240px; background:var(--card); box-shadow:var(--shadow); display:flex; flex-direction:column; padding:1.5rem 0; flex-shrink:0; }
    .logo { font-size:1.3rem; font-weight:700; padding:0 1.5rem 1.2rem; color:var(--primary); border-bottom:1px solid var(--border); margin-bottom:0.8rem; }
    .nav-item { padding:0.7rem 1.5rem; margin:0.15rem 0; cursor:pointer; border-left:3px solid transparent; transition:0.2s; color:#555; display:flex; align-items:center; gap:0.4rem; }
    .nav-item.active { background:#f0f4ff; border-left-color:var(--primary); color:var(--primary); font-weight:600; }
    .nav-item:hover { background:#f8f9ff; }
    .main { flex:1; display:flex; flex-direction:column; overflow-y:auto; padding:1.5rem; }
    .view { display:none; flex-direction:column; gap:1.5rem; }
    .view.active { display:flex; }
    .header { display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:0.5rem; }
    .btn { border:none; padding:0.55rem 1.1rem; border-radius:8px; font-weight:600; cursor:pointer; display:inline-flex; align-items:center; gap:0.3rem; font-size:0.9rem; transition:0.2s; }
    .btn-primary { background:var(--primary); color:#fff; } .btn-primary:hover { background:var(--primary-light); }
    .btn-danger { background:var(--danger); color:#fff; } .btn-danger:hover { background:#dc2626; }
    .card { background:var(--card); border-radius:var(--radius); padding:1.5rem; box-shadow:var(--shadow); }
    .stats-grid { display:grid; grid-template-columns: repeat(auto-fill, minmax(180px,1fr)); gap:1rem; margin-bottom:1rem; }
    .stat-card { background:var(--card); border-radius:var(--radius); padding:1rem; box-shadow:var(--shadow); text-align:center; }
    .stat-value { font-size:1.8rem; font-weight:700; color:var(--primary); }
    .stat-label { font-size:0.85rem; color:#555; margin-top:0.3rem; }
    .grid-2 { display:grid; grid-template-columns:repeat(auto-fit, minmax(350px,1fr)); gap:1.5rem; }
    .chart-container { width:100%; height:280px; }
    .upload-area { border:2px dashed var(--border); border-radius:var(--radius); padding:2rem; text-align:center; cursor:pointer; margin-bottom:1rem; }
    .upload-area:hover { border-color:var(--primary); background:#fafbff; }
    .table-wrapper { max-height:400px; overflow:auto; border:1px solid var(--border); border-radius:var(--radius); }
    table { width:100%; border-collapse:collapse; font-size:0.85rem; }
    th,td { padding:0.5rem 0.7rem; border-bottom:1px solid var(--border); text-align:center; }
    th { background:#f1f5f9; position:sticky; top:0; }
    .alert-badge { padding:0.3rem 0.8rem; border-radius:20px; font-size:0.8rem; font-weight:600; background:#fee2e2; color:var(--danger); }
    .info-row { padding:0.7rem 1rem; background:#f8f9ff; border-radius:8px; font-size:0.85rem; margin-bottom:0.8rem; }
    .prediction-table th { background:var(--primary); color:#fff; }
    .prediction-table td { background:#f8f9ff; font-weight:600; }
    .form-row { display:flex; gap:0.6rem; flex-wrap:wrap; align-items:flex-end; }
    .form-group { display:flex; flex-direction:column; min-width:100px; }
    .form-group label { font-size:0.78rem; margin-bottom:0.2rem; color:#555; }
    .record-entry { background:#f8fafd; border:1px solid var(--border); border-radius:var(--radius); padding:1rem; margin-bottom:1rem; }
    .record-entry img { max-width:200px; max-height:150px; margin:0.5rem 0; border-radius:8px; }
    .anomaly-badge { background:#ffe5e5; color:#b91c1c; padding:0.2rem 0.6rem; border-radius:20px; font-weight:600; font-size:0.8rem; }
    @media (max-width:768px) {
      body { flex-direction:column; }
      .sidebar { width:100%; flex-direction:row; overflow-x:auto; padding:0.5rem; }
      .logo { display:none; }
      .nav-item { padding:0.5rem 0.8rem; border-left:none; border-bottom:3px solid transparent; }
      .nav-item.active { border-left:none; border-bottom-color:var(--primary); }
    }
  </style>
</head>
<body>
<nav class="sidebar">
  <div class="logo">📊 数学量化</div>
  <div class="nav-item active" data-view="data">📁 数据管理</div>
  <div class="nav-item" data-view="dashboard">📈 量化仪表盘</div>
  <div class="nav-item" data-view="predict">🔮 预测中心</div>
  <div class="nav-item" data-view="records">📓 学习记录</div>
  <div class="nav-item" data-view="motivation">🏆 激励中心</div>
</nav>

<main class="main">
  <div style="background:var(--card);border-radius:8px;padding:0.5rem 1rem;display:flex;justify-content:space-between;box-shadow:var(--shadow);margin-bottom:0.5rem;">
    <span>🔥 连续 <strong id="streakDays">0</strong> 天</span>
    <span>⭐ 经验 <strong id="xpValue">0</strong></span>
    <span id="latestBadge" style="color:var(--warning);font-weight:600;"></span>
  </div>

  <!-- 数据管理 -->
  <section class="view active" id="view-data">
    <div class="header"><h2>📁 数据管理</h2><div><button class="btn btn-primary" id="saveSampleBtn">📥 加载150天数学示例</button><button class="btn btn-danger" id="deleteAllBtn">🗑️ 清空</button></div></div>
    <div class="card">
      <h3>📝 实时录入</h3>
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
      <div class="upload-area" id="uploadArea"><p>📂 拖拽 Excel 文件上传</p><input type="file" id="fileInput" accept=".xlsx,.xls,.csv" hidden></div>
      <div class="table-wrapper"><table id="dataTable"><thead></thead><tbody></tbody></table></div>
    </div>
  </section>

  <!-- 增强版仪表盘 -->
  <section class="view" id="view-dashboard">
    <div class="header"><h2>📈 学习量化仪表盘</h2></div>
    <div class="stats-grid" id="statsCards"></div>
    <div class="grid-2">
      <div class="card"><h4>📊 学习时长与测试趋势</h4><div class="chart-container" id="trendChart"></div></div>
      <div class="card"><h4>📈 测试分数分布</h4><div class="chart-container" id="distChart"></div></div>
    </div>
    <div class="grid-2">
      <div class="card"><h4>🎯 各维度平均雷达图</h4><div class="chart-container" id="radarChart"></div></div>
      <div class="card"><h4>🔍 练习 vs 测试相关性</h4><div class="chart-container" id="scatterChart"></div></div>
    </div>
    <div class="card" id="alertContainer"></div>
  </section>

  <!-- 预测中心 -->
  <section class="view" id="view-predict">
    <div class="header"><h2>🔮 预测中心 · 数学</h2></div>
    <div class="card" id="predictSettingsCard">
      <div class="collapse-header" id="settingsToggle"><strong>⚙️ 预测参数设置 (S型提升 / 遗忘底线)</strong><span id="settingsArrow">▼</span></div>
      <div id="settingsBody">
        <div class="form-row">
          <div class="form-group"><label>目标测试正确率%</label><input type="number" id="targetScore" value="90" min="30" max="100"></div>
          <div class="form-group"><label>S型中点(天)</label><input type="number" id="sMid" value="15" min="1"></div>
          <div class="form-group"><label>S型陡峭度 k</label><input type="number" id="sK" value="0.3" step="0.05"></div>
          <div class="form-group"><label>遗忘半衰期(天)</label><input type="number" id="halfLife" value="12" min="1"></div>
          <div class="form-group"><label>底线最低分%</label><input type="number" id="minScore" value="40" min="0" max="80"></div>
          <button class="btn btn-primary" id="savePredictSettingsBtn">💾 保存</button>
          <span id="settingsSavedHint" style="color:var(--success);display:none;">✅</span>
        </div>
      </div>
    </div>
    <div class="card" id="historyCard">
      <h3>📜 历史测试正确率走势（异常点 <span class="anomaly-badge">红色标注</span>）</h3>
      <div class="chart-container" id="historyChart"></div>
      <div class="table-wrapper" style="max-height:250px;margin-top:1rem;"><table id="historyTable"></table></div>
    </div>
    <div class="card" id="anomalyCard">
      <h3>⚠️ 检测到的异常点记录</h3>
      <div class="table-wrapper"><table id="anomalyTable"><thead><tr><th>日期</th><th>测试%</th><th>偏离倍数</th></tr></thead><tbody><tr><td colspan="3">暂无异常</td></tr></tbody></table></div>
    </div>
    <div class="card" id="forecastCard">
      <h3>🔮 未来30天预测（基准·S型提升·遗忘底线）</h3>
      <div class="info-row" id="forecastInfo"></div>
      <table class="prediction-table" id="forecastTable"></table>
      <div class="chart-container" id="forecastChart"></div>
    </div>
  </section>

  <!-- 学习记录 -->
  <section class="view" id="view-records">
    <div class="header"><h2>📓 学习情况记录</h2></div>
    <div class="card">
      <h3>📝 添加新记录</h3>
      <div class="form-row">
        <div class="form-group"><label>日期</label><input type="date" id="recordDate"></div>
        <div class="form-group"><label>文字记录</label><textarea id="recordText" rows="2" style="width:300px;"></textarea></div>
        <div class="form-group"><label>上传图片</label><input type="file" id="recordImage" accept="image/*"></div>
        <button class="btn btn-primary" id="addRecordNoteBtn">➕ 添加记录</button>
      </div>
    </div>
    <div class="card">
      <h3>📋 历史记录</h3>
      <div id="recordNotesContainer"></div>
    </div>
  </section>

  <!-- 激励中心 -->
  <section class="view" id="view-motivation">
    <div class="card"><h3>🏅 成就</h3><div id="achievementList" class="grid-2"></div></div>
    <div class="card"><h3>📜 惩罚记录</h3><div id="penaltyLog"></div></div>
  </section>
</main>

<script>
  (function() {
    // 存储键
    const STORAGE_DATA = 'sq_math_data_v3';
    const STORAGE_MOTIV = 'sq_math_motiv_v3';
    const STORAGE_SETTINGS = 'sq_math_settings_v3';
    const STORAGE_NOTES = 'sq_math_notes_v3';

    let records = [];
    let motivation = { xp:0, streak:0, achievements:{}, penaltyLog:[], lastCalcDate:'' };
    let forecastSettings = { target: 90, sMid: 15, sK: 0.3, halfLife: 12, minScore: 40 };
    let studyNotes = [];
    // 图表实例管理
    const chartInstances = {};

    function disposeChart(key) {
      if (chartInstances[key]) {
        chartInstances[key].dispose();
        delete chartInstances[key];
      }
    }

    // 持久化
    function saveData(){ localStorage.setItem(STORAGE_DATA, JSON.stringify(records)); }
    function loadData(){ const s=localStorage.getItem(STORAGE_DATA); if(s) try{records=JSON.parse(s)}catch(e){} }
    function saveMot(){ localStorage.setItem(STORAGE_MOTIV, JSON.stringify(motivation)); }
    function loadMot(){ const s=localStorage.getItem(STORAGE_MOTIV); if(s) try{motivation=JSON.parse(s)}catch(e){} }
    function saveSettings(){ localStorage.setItem(STORAGE_SETTINGS, JSON.stringify(forecastSettings)); }
    function loadSettings(){ const s=localStorage.getItem(STORAGE_SETTINGS); if(s) try{forecastSettings=JSON.parse(s)}catch(e){} }
    function saveNotes(){ localStorage.setItem(STORAGE_NOTES, JSON.stringify(studyNotes)); }
    function loadNotes(){ const s=localStorage.getItem(STORAGE_NOTES); if(s) try{studyNotes=JSON.parse(s)}catch(e){} }

    // 导航
    document.querySelectorAll('.nav-item').forEach(item => {
      item.addEventListener('click', () => {
        const vn = item.dataset.view;
        document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
        item.classList.add('active');
        document.querySelectorAll('.view').forEach(v=>v.classList.remove('active'));
        document.getElementById(`view-${vn}`).classList.add('active');
        if(vn==='data') renderDataTable();
        if(vn==='dashboard') renderDashboard();
        if(vn==='predict') { renderHistory(); renderForecast(); renderAnomalies(); }
        if(vn==='records') renderNotes();
        if(vn==='motivation') renderMotivation();
      });
    });

    document.getElementById('inputDate').valueAsDate = new Date();
    document.getElementById('addRecordBtn').addEventListener('click',()=>{
      records.unshift({
        subject: '数学', date: document.getElementById('inputDate').value,
        duration: +document.getElementById('inputDuration').value,
        practice: +document.getElementById('inputPractice').value,
        test: +document.getElementById('inputTest').value,
        review: +document.getElementById('inputReview').value,
        memory: +document.getElementById('inputMemory').value,
        weight: +document.getElementById('inputWeight').value
      });
      saveData(); updateAll();
    });
    window.deleteRecord = (idx) => { if(confirm('删除？')){records.splice(idx,1); saveData(); updateAll();} };
    document.getElementById('deleteAllBtn').addEventListener('click',()=>{ if(confirm('清空全部？')){records=[]; saveData(); updateAll();} });
    const uploadArea=document.getElementById('uploadArea'), fileInput=document.getElementById('fileInput');
    uploadArea.addEventListener('click',()=>fileInput.click());
    fileInput.addEventListener('change',e=>{ if(e.target.files[0]) processExcel(e.target.files[0]); });
    function processExcel(file){
      const reader=new FileReader();
      reader.onload=e=>{
        try{
          const wb=XLSX.read(e.target.result,{type:'binary'});
          const ws=wb.Sheets[wb.SheetNames[0]];
          const json=XLSX.utils.sheet_to_json(ws,{header:1});
          if(json.length<2) return alert('需要标题行');
          const headers=json[0];
          const map={'学科':'subject','日期':'date','学习时长(分钟)':'duration','练习正确率%':'practice','测试正确率%':'test','复习完成率%':'review','知识点记忆%':'memory','权重':'weight'};
          const parsed=json.slice(1).filter(r=>r.some(c=>c!=null)).map(r=>{
            const obj={}; headers.forEach((h,i)=>{ if(map[h]) obj[map[h]]= r[i]; });
            obj.practice=+obj.practice||0; obj.test=+obj.test||0; obj.review=+obj.review||0; obj.memory=+obj.memory||0; obj.duration=+obj.duration||0; obj.weight=+obj.weight||20;
            obj.subject = obj.subject || '数学';
            return obj;
          });
          records=[...records,...parsed]; saveData(); updateAll(); alert(`导入${parsed.length}条`);
        }catch(ex){alert('解析失败')}
      };
      reader.readAsBinaryString(file);
    }

    function loadMathSample(){
      if(records.length && !confirm('将替换为150天数学数据，继续？')) return;
      const start = new Date(); start.setDate(start.getDate() - 149);
      const sample = [];
      for (let i = 0; i < 150; i++) {
        const d = new Date(start); d.setDate(d.getDate() + i);
        const progress = i / 149;
        const baseScore = 45 + progress * 40 + Math.sin(progress * Math.PI * 10) * 15;
        const noise = (Math.random() - 0.5) * 25;
        let testScore = Math.round(Math.min(100, Math.max(0, baseScore + noise)));
        const practice = Math.min(100, testScore + Math.round((Math.random()-0.5)*12));
        const review = Math.min(100, testScore - 5 + Math.round((Math.random()-0.5)*15));
        const memory = Math.min(100, testScore - 3 + Math.round((Math.random()-0.5)*10));
        sample.push({
          subject: '数学', date: d.toISOString().split('T')[0],
          duration: 45 + Math.round(Math.random() * 40),
          practice: Math.max(0, practice), test: testScore,
          review: Math.max(0, review), memory: Math.max(0, memory), weight: 20
        });
      }
      records = sample; saveData(); updateAll();
    }
    document.getElementById('saveSampleBtn').addEventListener('click',loadMathSample);

    function renderDataTable(){
      const t=document.getElementById('dataTable');
      if(!records.length){ t.innerHTML='<tr><td colspan="9">暂无数据</td></tr>'; return; }
      const cols=['subject','date','duration','practice','test','review','memory','weight'];
      const hds=['学科','日期','时长','练习','测试','复习','记忆','权重'];
      t.innerHTML='<thead><tr>'+hds.map(h=>`<th>${h}</th>`).join('')+'<th>操作</th></tr></thead><tbody>';
      records.forEach((r,i)=>{ t.innerHTML+=`<tr>${cols.map(c=>`<td>${r[c]}</td>`).join('')}<td><button class="btn btn-danger btn-sm" onclick="deleteRecord(${i})">删除</button></td></tr>`; });
      t.innerHTML+='</tbody>';
    }

    // 增强仪表盘渲染
    function renderDashboard(){
      const math = records.filter(r=>r.subject==='数学');
      if(!math.length) {
        document.getElementById('statsCards').innerHTML = '<p>暂无数据</p>';
        document.getElementById('alertContainer').innerHTML = '';
        return;
      }
      // 统计卡片
      const latest = math.sort((a,b)=>new Date(b.date)-new Date(a.date))[0];
      const totalMin = math.reduce((s,r)=>s+r.duration,0);
      const days = new Set(math.map(r=>r.date)).size;
      const avgTest = (math.reduce((s,r)=>s+r.test,0)/math.length).toFixed(1);
      const maxTest = Math.max(...math.map(r=>r.test));
      const minTest = Math.min(...math.map(r=>r.test));
      const avgDuration = Math.round(totalMin / math.length);
      const overall = ((latest.practice+latest.test+latest.review+latest.memory)/4).toFixed(1);

      document.getElementById('statsCards').innerHTML = `
        <div class="stat-card"><div class="stat-value">${overall}</div><div class="stat-label">综合评分</div></div>
        <div class="stat-card"><div class="stat-value">${latest.test}%</div><div class="stat-label">最新测试正确率</div></div>
        <div class="stat-card"><div class="stat-value">${Math.floor(totalMin/60)}h ${totalMin%60}m</div><div class="stat-label">总学习时长</div></div>
        <div class="stat-card"><div class="stat-value">${days} 天</div><div class="stat-label">学习天数</div></div>
        <div class="stat-card"><div class="stat-value">${avgTest}%</div><div class="stat-label">平均测试正确率</div></div>
        <div class="stat-card"><div class="stat-value">${maxTest}% / ${minTest}%</div><div class="stat-label">最高 / 最低</div></div>
        <div class="stat-card"><div class="stat-value">${avgDuration} 分钟</div><div class="stat-label">日均学习时长</div></div>
      `;

      // 图表1：学习时长与测试趋势（双轴）
      const sorted = [...math].sort((a,b)=>new Date(a.date)-new Date(b.date));
      const recent100 = sorted.slice(-100);
      const dates = recent100.map(r=>r.date);
      const testVals = recent100.map(r=>r.test);
      const durVals = recent100.map(r=>r.duration);
      disposeChart('trend');
      const trendChart = echarts.init(document.getElementById('trendChart'));
      chartInstances.trend = trendChart;
      trendChart.setOption({
        tooltip: { trigger: 'axis' },
        legend: { bottom: 0 },
        xAxis: { type: 'category', data: dates, axisLabel: { rotate: 30, fontSize: 10 } },
        yAxis: [
          { type: 'value', name: '正确率%', max: 100 },
          { type: 'value', name: '时长(分钟)' }
        ],
        series: [
          { name: '测试正确率', type: 'line', data: testVals, smooth: true, color: '#4361ee' },
          { name: '学习时长', type: 'bar', yAxisIndex: 1, data: durVals, color: '#a0c4ff' }
        ]
      });

      // 图表2：测试分数分布直方图
      const bins = [0,40,50,60,70,80,90,101];
      const counts = bins.slice(0,-1).map((low,i) => {
        return math.filter(r => r.test >= low && r.test < bins[i+1]).length;
      });
      const binLabels = ['<40','40-50','50-60','60-70','70-80','80-90','90+'];
      disposeChart('dist');
      const distChart = echarts.init(document.getElementById('distChart'));
      chartInstances.dist = distChart;
      distChart.setOption({
        tooltip: { trigger: 'axis' },
        xAxis: { type: 'category', data: binLabels },
        yAxis: { type: 'value', name: '天数' },
        series: [{ type: 'bar', data: counts, color: '#10b981', barWidth: '60%' }]
      });

      // 图表3：雷达图
      disposeChart('radar');
      const radarChart = echarts.init(document.getElementById('radarChart'));
      chartInstances.radar = radarChart;
      radarChart.setOption({
        radar: { indicator: [
          { name: '练习', max:100 }, { name: '测试', max:100 },
          { name: '复习', max:100 }, { name: '记忆', max:100 }
        ]},
        series: [{ type: 'radar', data: [{ name:'数学', value:[latest.practice, latest.test, latest.review, latest.memory] }] }]
      });

      // 图表4：练习 vs 测试散点图
      disposeChart('scatter');
      const scatterChart = echarts.init(document.getElementById('scatterChart'));
      chartInstances.scatter = scatterChart;
      const scatterData = math.map(r => [r.practice, r.test]);
      scatterChart.setOption({
        xAxis: { type: 'value', name: '练习正确率%', max:100 },
        yAxis: { type: 'value', name: '测试正确率%', max:100 },
        series: [{ type: 'scatter', data: scatterData, symbolSize: 6, color: '#f59e0b' }]
      });

      // 预警
      const alerts = [];
      if(latest.test<60) alerts.push(`测试${latest.test}%`);
      if(latest.review<50) alerts.push(`复习${latest.review}%`);
      if(latest.memory<50) alerts.push(`记忆${latest.memory}%`);
      document.getElementById('alertContainer').innerHTML = `<h3>⚠️ 预警</h3>${alerts.length?alerts.map(a=>`<span class="alert-badge">${a}</span>`).join(' '):'<p class="text-success">✅ 一切正常</p>'}`;
    }

    // 激励 (不变)
    function updateMotivation(){
      const dates=[...new Set(records.map(r=>r.date))].sort();
      const today=new Date().toISOString().split('T')[0];
      let streak=0, set=new Set(dates), d=new Date(today);
      if(!set.has(today)) d.setDate(d.getDate()-1);
      while(set.has(d.toISOString().split('T')[0])){ streak++; d.setDate(d.getDate()-1); }
      motivation.streak=streak;
      const todayRecs=records.filter(r=>r.date===today);
      if(todayRecs.length && today!==motivation.lastCalcDate){
        const xp=Math.round(todayRecs.reduce((s,r)=>s+(r.test/100)*10+r.duration*0.5,0));
        motivation.xp+=xp; motivation.lastCalcDate=today;
      }
      if(streak>=7) motivation.achievements.streak7=true;
      if(records.reduce((s,r)=>s+r.duration,0)>3000) motivation.achievements.total3k=true;
      const lastDate=dates.length?new Date(dates[dates.length-1]):new Date(0);
      const diff=Math.floor((new Date()-lastDate)/86400000);
      if(diff>=3 && (!motivation.lastPenaltyDate || motivation.lastPenaltyDate!==today)){
        motivation.xp=Math.max(0,motivation.xp-10);
        motivation.penaltyLog.unshift({date:today,reason:`连续${diff}天未学习，扣除10经验`});
        motivation.lastPenaltyDate=today;
      }
      saveMot();
      document.getElementById('streakDays').textContent=streak;
      document.getElementById('xpValue').textContent=motivation.xp;
      document.getElementById('latestBadge').textContent=motivation.achievements.streak7?'🔥连续7天':'';
    }
    function renderMotivation(){
      const badges=[{id:'streak7',icon:'🔥',label:'连续7天'},{id:'total3k',icon:'⏱️',label:'学习超3000分钟'}];
      document.getElementById('achievementList').innerHTML=badges.map(b=>`<div class="metric-card" style="background:${motivation.achievements[b.id]?'#d1fae5':'#f3f4f6'}"><div style="font-size:1.8rem">${b.icon}</div><div>${b.label}</div></div>`).join('');
      document.getElementById('penaltyLog').innerHTML=motivation.penaltyLog.length?motivation.penaltyLog.map(p=>`<div>📅 ${p.date} - ${p.reason}</div>`).join(''):'🎉 暂无惩罚';
    }

    // 预测相关 (保持)
    function detectAnomalies(scores, dates) {
      const mean = scores.reduce((a,b)=>a+b,0)/scores.length;
      const variance = scores.reduce((s,v)=>s+(v-mean)**2,0)/scores.length;
      const std = Math.sqrt(variance);
      const threshold = 2 * std;
      const anomalies = [];
      scores.forEach((v,i) => {
        if (Math.abs(v-mean) > threshold) {
          anomalies.push({ date: dates[i], score: v, deviation: ((v-mean)/std).toFixed(2) });
        }
      });
      return anomalies;
    }
    function getRecentAll(recs) {
      const dates = [...new Set(recs.map(r=>r.date))].sort();
      return recs.filter(r=>dates.includes(r.date)).sort((a,b)=>new Date(a.date)-new Date(b.date));
    }
    function renderHistory() {
      const recs = records.filter(r=>r.subject==='数学');
      const recent = getRecentAll(recs).slice(-100);
      const chartDom = document.getElementById('historyChart'), tableDom = document.getElementById('historyTable');
      if(recent.length < 2) { chartDom.innerHTML='<p>数据不足</p>'; tableDom.innerHTML=''; return; }
      const dates = recent.map(r=>r.date), tests = recent.map(r=>r.test);
      const anomalies = detectAnomalies(tests, dates);
      disposeChart('history');
      const chart = echarts.init(chartDom);
      chartInstances.history = chart;
      chart.setOption({
        tooltip: { trigger: 'axis' },
        xAxis: { type: 'category', data: dates, axisLabel: { rotate: 45 } },
        yAxis: { max: 100 },
        series: [{
          name: '测试正确率', type: 'line', data: tests, smooth: true, color: '#4361ee',
          markPoint: {
            data: anomalies.map(a => ({
              name: a.date, coord: [a.date, a.score], value: a.score,
              symbol: 'pin', symbolSize: 40, itemStyle: { color: '#ef4444' },
              label: { formatter: '{c}%', fontSize: 10 }
            }))
          }
        }]
      });
      let html = '<thead><tr><th>日期</th><th>测试%</th><th>练习%</th><th>复习%</th><th>记忆%</th><th>状态</th></tr></thead><tbody>';
      recent.forEach(r => {
        const isAnomaly = anomalies.some(a=>a.date===r.date);
        html += `<tr><td>${r.date}</td><td>${r.test} ${isAnomaly?'<span class="anomaly-badge">异常</span>':''}</td><td>${r.practice}</td><td>${r.review}</td><td>${r.memory}</td><td>${isAnomaly?'⚠️':''}</td></tr>`;
      });
      tableDom.innerHTML = html + '</tbody>';
      window._lastAnomalies = anomalies;
    }
    function renderAnomalies() {
      const table = document.getElementById('anomalyTable');
      const anomalies = window._lastAnomalies || [];
      if (!anomalies.length) {
        table.innerHTML = '<thead><tr><th>日期</th><th>测试%</th><th>偏离倍数</th></tr></thead><tbody><tr><td colspan="3">暂无异常</td></tr></tbody>';
        return;
      }
      let html = '<thead><tr><th>日期</th><th>测试%</th><th>偏离标准差倍数</th></tr></thead><tbody>';
      anomalies.forEach(a => { html += `<tr><td>${a.date}</td><td>${a.score}%</td><td>${a.deviation}σ</td></tr>`; });
      table.innerHTML = html + '</tbody>';
    }
    function renderForecast() {
      const recs = records.filter(r=>r.subject==='数学');
      const recent = getRecentAll(recs).slice(-30);
      if (recent.length < 3) return;
      const scores = recent.map(r=>r.test); const n = scores.length;
      const current = scores[n-1];
      const sumX = scores.reduce((s,_,i)=>s+i,0), sumY = scores.reduce((s,v)=>s+v,0);
      const sumXY = scores.reduce((s,v,i)=>s+i*v,0), sumX2 = scores.reduce((s,_,i)=>s+i*i,0);
      const slope = (n*sumXY - sumX*sumY) / (n*sumX2 - sumX*sumX);
      const intercept = (sumY - slope*sumX) / n;
      const days = [0,5,10,15,20,25,30];
      const baseline = days.map(d => Math.min(100, Math.max(0, intercept + slope*(n-1+d)))).map(v=>+v.toFixed(1));
      baseline[0] = +current.toFixed(1);
      const {target, sMid, sK} = forecastSettings;
      const uplift = days.map(d => {
        if (d===0) return current;
        const growth = target - current;
        const logistic = 1 / (1 + Math.exp(-sK * (d - sMid)));
        return Math.min(100, current + growth * logistic);
      }).map(v=>+v.toFixed(1));
      const {halfLife, minScore} = forecastSettings;
      const decayRate = Math.log(2) / halfLife;
      const downside = days.map(d => {
        if (d===0) return current;
        const dropped = minScore + (current - minScore) * Math.exp(-decayRate * d);
        return Math.max(minScore, dropped);
      }).map(v=>+v.toFixed(1));
      document.getElementById('forecastInfo').innerHTML = `📊 当前测试正确率 ${current.toFixed(1)}% ，目标 ${target}%`;
      const labels = ['今天','+5d','+10d','+15d','+20d','+25d','+30d'];
      document.getElementById('forecastTable').innerHTML = '<thead><tr><th>类别</th>'+labels.map(l=>`<th>${l}</th>`).join('')+'</tr></thead><tbody>'+
        `<tr style="background:#eef2ff"><td>基准线</td>${baseline.map(v=>`<td>${v}%</td>`).join('')}</tr>`+
        `<tr style="background:#d1fae5"><td>S型提升</td>${uplift.map(v=>`<td>${v}%</td>`).join('')}</tr>`+
        `<tr style="background:#fee2e2"><td>遗忘底线</td>${downside.map(v=>`<td>${v}%</td>`).join('')}</tr></tbody>`;
      disposeChart('forecast');
      const chart = echarts.init(document.getElementById('forecastChart'));
      chartInstances.forecast = chart;
      chart.setOption({
        legend: { bottom: 0 },
        xAxis: { type: 'category', data: labels },
        yAxis: { max: 100, min: 0 },
        series: [
          { name:'基准线', type:'line', data:baseline, smooth:true, color:'#4361ee', symbol:'circle' },
          { name:'S型提升', type:'line', data:uplift, smooth:true, lineStyle:{type:'dashed',color:'#10b981'}, symbol:'diamond' },
          { name:'遗忘底线', type:'line', data:downside, smooth:true, lineStyle:{type:'dotted',color:'#ef4444'}, symbol:'triangle' }
        ]
      });
    }
    document.getElementById('settingsToggle').addEventListener('click',()=>{
      const body = document.getElementById('settingsBody');
      const arrow = document.getElementById('settingsArrow');
      if (body.style.display === 'none') { body.style.display = 'block'; arrow.textContent = '▼'; }
      else { body.style.display = 'none'; arrow.textContent = '▶'; }
    });
    document.getElementById('savePredictSettingsBtn').addEventListener('click',()=>{
      forecastSettings.target = +document.getElementById('targetScore').value || 90;
      forecastSettings.sMid = +document.getElementById('sMid').value || 15;
      forecastSettings.sK = +document.getElementById('sK').value || 0.3;
      forecastSettings.halfLife = +document.getElementById('halfLife').value || 12;
      forecastSettings.minScore = +document.getElementById('minScore').value || 40;
      saveSettings();
      document.getElementById('settingsSavedHint').style.display='inline';
      setTimeout(()=>document.getElementById('settingsSavedHint').style.display='none',2000);
      renderForecast();
    });
    function loadSettingsUI(){
      document.getElementById('targetScore').value = forecastSettings.target;
      document.getElementById('sMid').value = forecastSettings.sMid;
      document.getElementById('sK').value = forecastSettings.sK;
      document.getElementById('halfLife').value = forecastSettings.halfLife;
      document.getElementById('minScore').value = forecastSettings.minScore;
    }

    // 学习记录
    document.getElementById('addRecordNoteBtn').addEventListener('click', () => {
      const date = document.getElementById('recordDate').value;
      const text = document.getElementById('recordText').value.trim();
      const imageInput = document.getElementById('recordImage');
      if (!date) return alert('请选择日期');
      if (!text && !imageInput.files[0]) return alert('至少填写文字或上传图片');
      const processAndSave = (imageBase64 = '') => {
        studyNotes.unshift({ date, text, imageBase64 });
        saveNotes(); renderNotes();
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
      if (!studyNotes.length) { container.innerHTML = '<p>暂无学习记录，请添加。</p>'; return; }
      container.innerHTML = studyNotes.map(note => `
        <div class="record-entry">
          <div><strong>${note.date}</strong></div>
          <div style="white-space:pre-wrap; margin:0.5rem 0;">${note.text || '(无文字)'}</div>
          ${note.imageBase64 ? `<img src="${note.imageBase64}" alt="学习照片">` : ''}
        </div>
      `).join('');
    }

    function updateAll() {
      renderDataTable(); updateMotivation();
      if (document.getElementById('view-dashboard').classList.contains('active')) renderDashboard();
      if (document.getElementById('view-predict').classList.contains('active')) { renderHistory(); renderForecast(); renderAnomalies(); }
      if (document.getElementById('view-records').classList.contains('active')) renderNotes();
      if (document.getElementById('view-motivation').classList.contains('active')) renderMotivation();
    }

    loadData(); loadMot(); loadSettings(); loadNotes();
    if (!records.length) loadMathSample();
    else updateAll();
    loadSettingsUI();
    document.getElementById('recordDate').valueAsDate = new Date();
    document.querySelector('.nav-item[data-view="data"]').click();
  })();
</script>
</body>
</html>
