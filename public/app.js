let currentData = null;
let radarChartInstance = null;

// ============================================================
// Start Analysis
// ============================================================
async function startAnalysis() {
  const url = document.getElementById('urlInput').value.trim();
  if (!url) {
    document.getElementById('errorMsg').textContent = 'URLを入力してください';
    return;
  }

  try {
    new URL(url);
  } catch {
    document.getElementById('errorMsg').textContent = '有効なURLを入力してください';
    return;
  }

  document.getElementById('errorMsg').textContent = '';
  document.getElementById('analyzeBtn').disabled = true;
  document.getElementById('loadingOverlay').style.display = 'flex';

  animateSteps();

  try {
    const response = await fetch('/api/analyze', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ url })
    });

    if (!response.ok) {
      const err = await response.json();
      throw new Error(err.error || '分析に失敗しました');
    }

    currentData = await response.json();
    renderReport(currentData);

    document.getElementById('inputScreen').style.display = 'none';
    document.getElementById('reportScreen').style.display = 'block';
  } catch (error) {
    document.getElementById('errorMsg').textContent = error.message;
  } finally {
    document.getElementById('analyzeBtn').disabled = false;
    document.getElementById('loadingOverlay').style.display = 'none';
  }
}

function animateSteps() {
  const steps = ['step1', 'step2', 'step3', 'step4'];
  const delays = [0, 3000, 8000, 15000];

  steps.forEach((id) => {
    document.getElementById(id).className = 'step';
  });

  steps.forEach((id, i) => {
    setTimeout(() => {
      for (let j = 0; j < i; j++) {
        document.getElementById(steps[j]).className = 'step done';
      }
      document.getElementById(id).className = 'step active';
    }, delays[i]);
  });
}

// ============================================================
// Render Report
// ============================================================
function renderReport(data) {
  const now = new Date();
  const dateStr = `${now.getFullYear()}/${String(now.getMonth() + 1).padStart(2, '0')}`;

  // Slide 1 header
  document.getElementById('rptCompanyName').textContent = data.companyName || '';
  document.getElementById('rptDate').textContent = dateStr;
  document.getElementById('rptTotalScore').textContent = data.totalScore || 0;

  // Slide 2 header
  document.getElementById('rptCompanyName2').textContent = data.companyName || '';
  document.getElementById('rptDate2').textContent = dateStr;

  // Diagnostic type
  const dtText = data.diagnosticType || '';
  document.getElementById('rptDiagType').textContent = `【${dtText}】`;
  document.getElementById('rptDiagDesc').textContent = data.diagnosticTypeDescription || '';

  // Slide 1: Summary
  renderSummaryContent(data);

  // Slide 2: Proposals
  renderProposalsContent(data);

  // One-point advice
  if (data.onePointAdvice) {
    document.getElementById('rptAdvice').innerHTML =
      `<strong>${escapeHtml(data.onePointAdvice.focus)}</strong><br><br>${escapeHtml(data.onePointAdvice.advice)}`;
  }

  // Radar chart
  renderRadarChart(data);

  // Detail tables
  renderDetailTables(data);
}

// ---- Slide 1: Summary ----
function renderSummaryContent(data) {
  const el = document.getElementById('rptSummaryContent');
  let html = '';
  html += `<div style="margin-bottom:14px;">今回は「<span style="color:#2980b9;font-weight:bold;">${escapeHtml(data.url)}</span>」サイトを対象に調査しています。</div>`;
  html += `<div>${escapeHtml(data.summary || '')}</div>`;
  el.innerHTML = html;
}

// ---- Slide 2: Proposals ----
function renderProposalsContent(data) {
  const el = document.getElementById('rptProposalsContent');
  let html = '';

  html += `<div class="section-title">【さらなる向上のための改善提案】</div>`;
  html += `<p style="margin-bottom:16px;font-size:13px;color:#666;">本サイトをより使いやすくし、問合せ削減につなげるための具体的なステップです。</p>`;

  if (data.proposals) {
    data.proposals.forEach((p, i) => {
      const priority = i === 0 ? '  <span style="background:#e74c3c;color:#fff;font-size:11px;padding:2px 8px;border-radius:3px;margin-left:8px;">優先度 高</span>' : '';
      html += `<div class="proposal-block">`;
      html += `<div class="proposal-title">${i + 1}) ${escapeHtml(p.title)}${priority}</div>`;
      html += `<div class="proposal-current"><strong>現状:</strong> ${escapeHtml(p.current)}</div>`;
      html += `<div class="proposal-suggestion"><strong>提案:</strong> → ${escapeHtml(p.suggestion)}</div>`;
      html += `</div>`;
    });
  }

  el.innerHTML = html;
}

// ============================================================
// Radar Chart
// ============================================================
function renderRadarChart(data) {
  const ctx = document.getElementById('radarChart').getContext('2d');

  if (radarChartInstance) {
    radarChartInstance.destroy();
  }

  const scores = data.scores || {};
  const companyScores = [
    scores.induction?.subtotal || 0,
    scores.classification?.subtotal || 0,
    scores.content?.subtotal || 0,
    scores.functionality?.subtotal || 0
  ];

  const avgScores = [
    data.averageScores?.induction || 5,
    data.averageScores?.classification || 5,
    data.averageScores?.content || 7,
    data.averageScores?.functionality || 3
  ];

  radarChartInstance = new Chart(ctx, {
    type: 'radar',
    data: {
      labels: ['誘導', '分類', '表記', '機能'],
      datasets: [
        {
          label: '御社',
          data: companyScores,
          borderColor: '#CC0000',
          backgroundColor: 'rgba(204, 0, 0, 0.1)',
          borderWidth: 2.5,
          pointBackgroundColor: '#CC0000',
          pointRadius: 4
        },
        {
          label: '他社平均',
          data: avgScores,
          borderColor: '#334466',
          backgroundColor: 'rgba(51, 68, 102, 0.05)',
          borderWidth: 2,
          pointBackgroundColor: '#334466',
          pointRadius: 3
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            color: '#333',
            font: { size: 11 },
            padding: 10,
            usePointStyle: true
          }
        }
      },
      scales: {
        r: {
          min: 0,
          max: 10,
          ticks: {
            stepSize: 5,
            color: '#999',
            font: { size: 9 },
            backdropColor: 'transparent'
          },
          grid: { color: 'rgba(0,0,0,0.1)' },
          angleLines: { color: 'rgba(0,0,0,0.1)' },
          pointLabels: {
            color: '#1a3a5c',
            font: { size: 13, weight: 'bold' }
          }
        }
      }
    },
    plugins: [{
      afterDatasetsDraw(chart) {
        const meta = chart.getDatasetMeta(0);
        const ctx2 = chart.ctx;
        ctx2.save();
        ctx2.font = 'bold 11px Meiryo';
        ctx2.fillStyle = '#CC0000';
        ctx2.textAlign = 'center';
        meta.data.forEach((point, i) => {
          const val = companyScores[i].toFixed(1);
          ctx2.fillText(val, point.x, point.y - 10);
        });
        ctx2.restore();
      }
    }]
  });
}

// ============================================================
// Detail Tables
// ============================================================
function renderDetailTables(data) {
  const container = document.getElementById('detailTables');
  const categories = [
    { key: 'induction', name: 'A. 誘導 (Induction)', max: 8 },
    { key: 'classification', name: 'B. 分類 (Classification)', max: 4 },
    { key: 'content', name: 'C. 表記 (Content)', max: 8 },
    { key: 'functionality', name: 'D. 機能 (Functionality)', max: 8 }
  ];

  let html = '';
  categories.forEach(cat => {
    const catData = data.scores?.[cat.key];
    if (!catData) return;

    html += `<table class="detail-table">`;
    html += `<thead><tr><th>${cat.name}</th><th style="width:60px">評価</th><th style="width:60px">点数</th><th>コメント</th></tr></thead>`;
    html += `<tbody>`;

    (catData.items || []).forEach(item => {
      const scoreClass = item.score === 2 ? 'score-good' : item.score === 1 ? 'score-ok' : 'score-improve';
      html += `<tr>`;
      html += `<td>${escapeHtml(item.name)}</td>`;
      html += `<td style="text-align:center"><span class="score-circle ${scoreClass}">${item.symbol}</span></td>`;
      html += `<td style="text-align:center">${item.score}</td>`;
      html += `<td>${escapeHtml(item.comment)}</td>`;
      html += `</tr>`;
    });

    html += `<tr class="subtotal-row"><td>小計</td><td></td><td style="text-align:center">${catData.subtotal} / ${cat.max}</td><td></td></tr>`;
    html += `</tbody></table>`;
  });

  container.innerHTML = html;
}

// ============================================================
// Download PNG (both slides combined)
// ============================================================
async function downloadPNG() {
  try {
    const slide1 = document.getElementById('reportSlide1');
    const slide2 = document.getElementById('reportSlide2');

    const canvas1 = await html2canvas(slide1, { scale: 2, backgroundColor: '#ffffff', useCORS: true, logging: false });
    const canvas2 = await html2canvas(slide2, { scale: 2, backgroundColor: '#ffffff', useCORS: true, logging: false });

    // Combine into one tall image with gap
    const gap = 40;
    const combined = document.createElement('canvas');
    combined.width = Math.max(canvas1.width, canvas2.width);
    combined.height = canvas1.height + gap + canvas2.height;
    const ctx = combined.getContext('2d');
    ctx.fillStyle = '#e8ecef';
    ctx.fillRect(0, 0, combined.width, combined.height);
    ctx.drawImage(canvas1, 0, 0);
    ctx.drawImage(canvas2, 0, canvas1.height + gap);

    const link = document.createElement('a');
    link.download = `FAQ診断レポート_${currentData?.companyName || 'report'}.png`;
    link.href = combined.toDataURL('image/png');
    link.click();
  } catch (error) {
    alert('PNG生成中にエラーが発生しました: ' + error.message);
  }
}

// ============================================================
// Download PPTX
// ============================================================
async function downloadPPTX() {
  if (!currentData?.analysisId) {
    alert('分析結果がありません');
    return;
  }

  try {
    const radarCanvas = document.getElementById('radarChart');
    const radarImage = radarCanvas.toDataURL('image/png');

    const response = await fetch('/api/generate-pptx', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        analysisId: currentData.analysisId,
        radarChartImage: radarImage
      })
    });

    if (!response.ok) {
      const err = await response.json();
      throw new Error(err.error);
    }

    const blob = await response.blob();
    const link = document.createElement('a');
    link.download = `FAQ診断レポート_${currentData?.companyName || 'report'}.pptx`;
    link.href = URL.createObjectURL(blob);
    link.click();
    URL.revokeObjectURL(link.href);
  } catch (error) {
    alert('PPTX生成中にエラーが発生しました: ' + error.message);
  }
}

// ============================================================
// Navigation
// ============================================================
function showInputScreen() {
  document.getElementById('reportScreen').style.display = 'none';
  document.getElementById('inputScreen').style.display = 'block';
}

// ============================================================
// Utility
// ============================================================
function escapeHtml(str) {
  if (!str) return '';
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('urlInput').addEventListener('keypress', (e) => {
    if (e.key === 'Enter') startAnalysis();
  });
});
