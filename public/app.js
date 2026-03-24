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

  // Animate loading steps
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

  steps.forEach((id, i) => {
    document.getElementById(id).className = 'step';
  });

  steps.forEach((id, i) => {
    setTimeout(() => {
      // Mark previous as done
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
  // Header info
  document.getElementById('rptCompanyName').textContent = data.companyName || '';
  const now = new Date();
  document.getElementById('rptDate').textContent =
    `${now.getFullYear()}/${String(now.getMonth() + 1).padStart(2, '0')}`;

  // Score
  document.getElementById('rptTotalScore').textContent = data.totalScore || 0;

  // Diagnostic type
  const dtText = data.diagnosticType || '';
  document.getElementById('rptDiagType').textContent = `【${dtText}】`;
  document.getElementById('rptDiagDesc').textContent = data.diagnosticTypeDescription || '';

  // Main content
  renderMainContent(data);

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

function renderMainContent(data) {
  const el = document.getElementById('rptMainContent');
  let html = '';

  // URL and summary
  html += `<div style="margin-bottom:12px;font-size:13px;">今回は「<span style="color:#2980b9;font-weight:bold;">${escapeHtml(data.url)}</span>」サイトを対象に調査しています。</div>`;
  html += `<div style="margin-bottom:16px;font-size:13px;line-height:1.9;">${escapeHtml(data.summary || '')}</div>`;

  // Proposals
  html += `<div class="section-title">【改善提案】</div>`;
  if (data.proposals) {
    data.proposals.forEach((p, i) => {
      html += `<div style="margin-bottom:14px;">`;
      html += `<span class="proposal-title">${i + 1}) ${escapeHtml(p.title)}</span><br>`;
      html += `<span style="font-size:12px;">${escapeHtml(p.current)}</span><br>`;
      html += `<span class="arrow-text" style="font-size:12px;">→ ${escapeHtml(p.suggestion)}</span>`;
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
          grid: {
            color: 'rgba(0,0,0,0.1)'
          },
          angleLines: {
            color: 'rgba(0,0,0,0.1)'
          },
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
        const ctx = chart.ctx;
        ctx.save();
        ctx.font = 'bold 11px Meiryo';
        ctx.fillStyle = '#CC0000';
        ctx.textAlign = 'center';
        meta.data.forEach((point, i) => {
          const val = companyScores[i].toFixed(1);
          ctx.fillText(val, point.x, point.y - 10);
        });
        ctx.restore();
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
// Download PNG
// ============================================================
async function downloadPNG() {
  const slide = document.getElementById('reportSlide');

  try {
    const canvas = await html2canvas(slide, {
      scale: 2,
      backgroundColor: '#e8e8e8',
      useCORS: true,
      logging: false
    });

    const link = document.createElement('a');
    link.download = `FAQ診断レポート_${currentData?.companyName || 'report'}.png`;
    link.href = canvas.toDataURL('image/png');
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
    // Get radar chart as image
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

// Allow Enter key to trigger analysis
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('urlInput').addEventListener('keypress', (e) => {
    if (e.key === 'Enter') startAnalysis();
  });
});
