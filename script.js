// Dados de exemplo para os gráficos
const months = ['dez/2025 - mar/2026','jan/2026 - abr/2026','fev/2026 - mai/2026','mar/2026 - jun/2026','abr/2026 - jul/2026','mai/2026 - ago/2026','jun/2026 - set/2026'];
const maturityAvg = [3.81,3.40,3.31,3.39,3.35,3.37,3.04];
const projectsTotal = [18,28,39,57,57,63,96];

// Gauge (usando doughnut para simular)
const gaugeCtx = document.getElementById('gaugeChart').getContext('2d');
const gaugeValue = 3.04; // exemplo
new Chart(gaugeCtx, {
  type: 'doughnut',
  data: {
    datasets: [{
      data: [gaugeValue, 5 - gaugeValue],
      backgroundColor: ['#27ae60','#eee'],
      borderWidth: 0
    }]
  },
  options: {
    cutout: '80%',
    responsive:true,
    maintainAspectRatio:false,
    plugins:{legend:{display:false},tooltip:{enabled:false}}
  }
});

// Bar + line combo chart
const barLineCtx = document.getElementById('barLineChart').getContext('2d');
new Chart(barLineCtx, {
  data: {
    labels: months,
    datasets: [
      { type:'bar', label:'Média Geral Por Mês', data:maturityAvg, backgroundColor:'#27ae60', yAxisID:'A' },
      { type:'line', label:'Total de Projetos', data:projectsTotal, borderColor:'#6b2ea0', backgroundColor:'#6b2ea0', yAxisID:'B', tension:0.3, pointRadius:6 }
    ]
  },
  options: {
    responsive:true,
    scales: {
      A: { type:'linear', position:'left', display:false, min:0, max:5 },
      B: { type:'linear', position:'right', grid:{drawOnChartArea:false}, beginAtZero:true }
    },
    plugins:{legend:{position:'top'}}
  }
});

// Combo chart (stacked bars example)
const comboCtx = document.getElementById('comboChart').getContext('2d');
const totalClients = [10,10,10,14,14,14,14];
const totalProjects = [18,28,39,57,57,63,96];
const projectsWithAssessment = [10,10,28,39,51,53,88];
new Chart(comboCtx, {
  type:'bar',
  data:{labels:months,datasets:[
    {label:'Total de Clientes',data:totalClients,backgroundColor:'#2f6be4'},
    {label:'Total de Projetos',data:totalProjects,backgroundColor:'#ff7f50'},
    {label:'Projetos com Assessment',data:projectsWithAssessment,backgroundColor:'#6b2ea0'}
  ]},
  options:{responsive:true,plugins:{legend:{position:'top'}},scales:{y:{beginAtZero:true}}}
});
