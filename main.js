// Office ist bereit, Registrierung des Service Workers
Office.onReady(function(info) {
  if (info.host === Office.HostType.Excel) {
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('/service-worker.js').then(function() {
        console.log('Service Worker registered');
      }).catch(function(error) {
        console.error('Service Worker registration failed:', error);
      });
    }
  }
});

// Funktion zur Aktualisierung der Diagrammsichtbarkeit basierend auf dem Zellinhalt
function updateChartVisibility() {
  var cellAddress = document.getElementById('cellAddress').value;

  if (!cellAddress) {
    console.error('Please enter a cell address.');
    return;
  }

  Excel.run(function(context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var cell = sheet.getRange(cellAddress);
    cell.load("values");

    var charts = sheet.charts;
    charts.load("items/name, items/visible");

    return context.sync().then(function() {
      var targetChartName = cell.values[0][0];

      // Schleife durch alle Diagramme und aktualisiere die Sichtbarkeit
      charts.items.forEach(function(chart) {
        if (chart.name === targetChartName) {
          chart.visible = true;
        } else {
          chart.visible = false;
        }
      });

      return context.sync();
    });
  }).catch(function(error) {
    console.error(error);
  });
}
