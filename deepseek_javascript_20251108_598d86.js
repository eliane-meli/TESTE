function doGet(e) {
  try {
    // ID da sua planilha
    const SPREADSHEET_ID = '1iLftWtEUacg6iYCzZKhChR6b5nqrNQxTz6R3lulBCro';
    const SHEET_NAME = 'Expedido';
    
    // Obter parâmetros da URL
    const startDate = e.parameter.startDate;
    const endDate = e.parameter.endDate;
    const weekdayFilter = e.parameter.weekday;
    
    // Conectar à planilha
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    // Obter dados
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    
    // Processar dados
    const processedData = processData(data, startDate, endDate, weekdayFilter);
    
    // Retornar como JSON
    return ContentService
      .createTextOutput(JSON.stringify(processedData))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeaders({'Access-Control-Allow-Origin': '*'});
      
  } catch (error) {
    // Em caso de erro
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeaders({'Access-Control-Allow-Origin': '*'});
  }
}

function processData(data, startDate, endDate, weekdayFilter) {
  // Pular cabeçalho
  const rows = data.slice(1);
  
  // Converter datas de filtro
  const start = startDate ? new Date(startDate) : null;
  const end = endDate ? new Date(endDate) : null;
  
  // Arrays para armazenar dados
  const records = [];
  const weekdaySums = {};
  const weekdayCounts = {};
  const monthlySums = {};
  const monthlyCounts = {};
  
  // Processar cada linha
  rows.forEach(row => {
    const dateStr = row[0]; // Coluna A - Data
    const quantity = parseFloat(row[1]); // Coluna B - QNTD Expedida
    const weekday = row[3]; // Coluna D - Dias da Semana
    
    // Pular linhas vazias ou com dados inválidos
    if (!dateStr || isNaN(quantity)) return;
    
    // Converter data
    const dateParts = dateStr.split('/');
    const date = new Date(dateParts[2], dateParts[1] - 1, dateParts[0]);
    
    // Aplicar filtros
    if (start && date < start) return;
    if (end && date > end) return;
    if (weekdayFilter && weekdayFilter !== 'all' && weekday !== weekdayFilter) return;
    
    // Formatar mês/ano
    const monthYear = `${date.getMonth() + 1}/${date.getFullYear()}`;
    
    // Adicionar ao registro
    records.push({
      date: dateStr,
      quantity: quantity,
      weekday: weekday,
      month: monthYear
    });
    
    // Calcular totais por dia da semana
    if (!weekdaySums[weekday]) {
      weekdaySums[weekday] = 0;
      weekdayCounts[weekday] = 0;
    }
    weekdaySums[weekday] += quantity;
    weekdayCounts[weekday] += 1;
    
    // Calcular totais mensais
    if (!monthlySums[monthYear]) {
      monthlySums[monthYear] = 0;
      monthlyCounts[monthYear] = 0;
    }
    monthlySums[monthYear] += quantity;
    monthlyCounts[monthYear] += 1;
  });
  
  // Calcular médias por dia da semana
  const weekdayData = {};
  Object.keys(weekdaySums).forEach(weekday => {
    weekdayData[weekday] = weekdaySums[weekday] / weekdayCounts[weekday];
  });
  
  // Calcular totais mensais
  const monthlyData = {};
  Object.keys(monthlySums).forEach(month => {
    monthlyData[month] = monthlySums[month];
  });
  
  // Calcular estatísticas gerais
  const quantities = records.map(r => r.quantity);
  const totalExpedited = quantities.reduce((sum, q) => sum + q, 0);
  const dailyAverage = totalExpedited / quantities.length;
  const maxVolume = Math.max(...quantities);
  const minVolume = Math.min(...quantities);
  
  // Encontrar datas de máximo e mínimo
  const maxRecord = records.find(r => r.quantity === maxVolume);
  const minRecord = records.find(r => r.quantity === minVolume);
  
  // Calcular distribuição de volumes
  const avgVolume = dailyAverage;
  const lowVolume = quantities.filter(q => q < avgVolume * 0.7).length;
  const mediumVolume = quantities.filter(q => q >= avgVolume * 0.7 && q <= avgVolume * 1.3).length;
  const highVolume = quantities.filter(q => q > avgVolume * 1.3).length;
  
  return {
    success: true,
    stats: {
      totalExpedited: totalExpedited,
      dailyAverage: dailyAverage,
      maxVolume: maxVolume,
      maxVolumeDate: maxRecord ? maxRecord.date : '-',
      minVolume: minVolume,
      minVolumeDate: minRecord ? minRecord.date : '-'
    },
    weekdayData: weekdayData,
    monthlyData: monthlyData,
    distributionData: {
      lowVolume: lowVolume,
      mediumVolume: mediumVolume,
      highVolume: highVolume
    },
    records: records.slice(0, 100) // Limitar a 100 registros na tabela
  };
}