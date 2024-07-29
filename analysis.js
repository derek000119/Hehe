

// Step 5: Sentiment Analysis
function analyzeSentiment() {
    var sheet = SpreadsheetApp.openById('1KWkpd4ZV-QzagcsM8JjKzT3eIUAdPq7pGO3KAAwX1xA'); // Replace with actual sheet ID
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    var sentimentSheet = sheet.getSheetByName('Sentiment Analysis') || sheet.insertSheet('Sentiment Analysis');
    sentimentSheet.clear();
    sentimentSheet.appendRow(['Date', 'Satisfaction Score', 'Feedback Sentiment', 'Area']);
    
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var date = row[0];
      var satisfactionScore = row[1];
      var challenges = row[2];
      var suggestions = row[3];
      var area = row[4];
      
      var feedbackText = challenges + ' ' + suggestions;
      var sentiment = getSentiment(feedbackText);
      
      sentimentSheet.appendRow([date, satisfactionScore, sentiment, area]);
    }
  }
  
  function getSentiment(text) {
    var positiveWords = [
      'good', 'great', 'excellent', 'amazing', 'fantastic', 'helpful', 'productive',
      'efficient', 'satisfying', 'inspiring', 'motivating', 'innovative', 'supportive',
      'collaborative', 'engaging', 'positive', 'proactive', 'empowering', 'rewarding',
      'fulfilling', 'valuable', 'beneficial', 'effective', 'successful', 'outstanding',
      'exceptional', 'impressive', 'remarkable', 'commendable', 'praiseworthy',
      'appreciative', 'harmonious', 'thriving', 'flourishing', 'progressive'
    ];
    
    var negativeWords = [
      'bad', 'poor', 'terrible', 'awful', 'unhelpful', 'frustrating', 'difficult',
      'disappointing', 'stressful', 'challenging', 'overwhelming', 'inefficient',
      'unproductive', 'demotivating', 'discouraging', 'unsupportive', 'hostile',
      'toxic', 'negative', 'problematic', 'burdensome', 'tedious', 'exhausting',
      'unfair', 'inadequate', 'subpar', 'unsatisfactory', 'mediocre', 'lacking',
      'insufficient', 'ineffective', 'counterproductive', 'detrimental', 'hindering',
      'obstructive'
    ];
    
    var words = text.toLowerCase().split(/\W+/);
    var positiveCount = words.filter(word => positiveWords.includes(word)).length;
    var negativeCount = words.filter(word => negativeWords.includes(word)).length;
    
    // Calculate sentiment score
    var totalWords = words.length;
    var sentimentScore = (positiveCount - negativeCount) / totalWords;
    
    // Classify sentiment based on score
    if (sentimentScore > 0.05) return 'Positive';
    if (sentimentScore < -0.05) return 'Negative';
    return 'Neutral';
  }
  
  
  
  function createDashboard() {
    var sheet = SpreadsheetApp.openById('1KWkpd4ZV-QzagcsM8JjKzT3eIUAdPq7pGO3KAAwX1xA'); // Replace with actual sheet ID
    var sentimentSheet = sheet.getSheetByName('Sentiment Analysis');
    var dashboardSheet = sheet.getSheetByName('Dashboard') || sheet.insertSheet('Dashboard');
    dashboardSheet.clear();
  
    // Prepare data
    var data = sentimentSheet.getDataRange().getValues();
    var headers = data.shift(); // Remove and store headers
  
    // Overall Sentiment Chart
    var sentimentCounts = {};
    data.forEach(function(row) {
      var sentiment = row[2]; // Assuming sentiment is in column C (index 2)
      sentimentCounts[sentiment] = (sentimentCounts[sentiment] || 0) + 1;
    });
  
    var sentimentChartData = [['Sentiment', 'Count']];
    for (var sentiment in sentimentCounts) {
      sentimentChartData.push([sentiment, sentimentCounts[sentiment]]);
    }
  
    var sentimentChartRange = dashboardSheet.getRange(1, 1, sentimentChartData.length, 2);
    sentimentChartRange.setValues(sentimentChartData);
  
    var sentimentChart = dashboardSheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sentimentChartRange)
      .setPosition(1, 1, 0, 0)
      .setOption('title', 'Overall Sentiment')
      .setOption('width', 400)
      .setOption('height', 300)
      .build();
    
    dashboardSheet.insertChart(sentimentChart);
  
    // Satisfaction Score Trend
    var satisfactionData = data.map(function(row) {
      return [new Date(row[0]), row[1]]; // Assuming date is in column A (index 0) and satisfaction score in column B (index 1)
    });
  
    satisfactionData.sort(function(a, b) {
      return a[0] - b[0];
    });
  
    var satisfactionChartData = [['Date', 'Satisfaction Score']].concat(satisfactionData);
  
    var satisfactionChartRange = dashboardSheet.getRange(1, 4, satisfactionChartData.length, 2);
    satisfactionChartRange.setValues(satisfactionChartData);
  
    var satisfactionChart = dashboardSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(satisfactionChartRange)
      .setPosition(1, 4, 0, 0)
      .setOption('title', 'Satisfaction Score Trend')
      .setOption('hAxis.title', 'Date')
      .setOption('vAxis.title', 'Score')
      .setOption('width', 600)
      .setOption('height', 400)
      .build();
    
    dashboardSheet.insertChart(satisfactionChart);
  
    // Feedback by Area
    var areaCounts = {};
    data.forEach(function(row) {
      var area = row[3]; // Assuming area is in column D (index 3)
      areaCounts[area] = (areaCounts[area] || 0) + 1;
    });
  
    var areaChartData = [['Area', 'Count']];
    for (var area in areaCounts) {
      areaChartData.push([area, areaCounts[area]]);
    }
  
    var areaChartRange = dashboardSheet.getRange(20, 1, areaChartData.length, 2);
    areaChartRange.setValues(areaChartData);
  
    var areaChart = dashboardSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(areaChartRange)
      .setPosition(20, 1, 0, 0)
      .setOption('title', 'Feedback by Area')
      .setOption('hAxis.title', 'Area')
      .setOption('vAxis.title', 'Count')
      .setOption('width', 600)
      .setOption('height', 400)
      .build();
    
    dashboardSheet.insertChart(areaChart);
  
    Logger.log('Dashboard created successfully!');
  }
  
  // Rest of the code remains the same...
  
  // Step 7: Automated Alerts
  function checkForAlerts() {
    var sheet = SpreadsheetApp.openById('1KWkpd4ZV-QzagcsM8JjKzT3eIUAdPq7pGO3KAAwX1xA'); // Replace with actual sheet ID
    var sentimentSheet = sheet.getSheetByName('Sentiment Analysis');
    var dataRange = sentimentSheet.getDataRange();
    var values = dataRange.getValues();
    
    // Find the index of the "Feedback Sentiment" column
    var headers = values[0];
    var sentimentColumnIndex = headers.indexOf('Feedback Sentiment');
    
    if (sentimentColumnIndex === -1) {
      Logger.log('Error: "Feedback Sentiment" column not found');
      return;
    }
    
    // Get the last 100 sentiment values (or less if there are fewer rows)
    var rowCount = values.length;
    var startRow = Math.max(1, rowCount - 100);  // Start from the 100th last row, or the 2nd row if there are less than 100 rows
    var recentSentiments = values.slice(startRow).map(row => row[sentimentColumnIndex]);
    
    var sentimentCounts = {
      Positive: 0,
      Neutral: 0,
      Negative: 0
    };
    
    recentSentiments.forEach(sentiment => {
      if (sentimentCounts.hasOwnProperty(sentiment)) {
        sentimentCounts[sentiment]++;
      }
    });
    
    // Calculate percentages
    var totalCount = recentSentiments.length;
    var percentages = {};
    for (var sentiment in sentimentCounts) {
      percentages[sentiment] = (sentimentCounts[sentiment] / totalCount * 100).toFixed(2);
    }
    
    // Log the sentiment counts and percentages
    Logger.log('Sentiment analysis of last ' + totalCount + ' feedback entries:');
    Logger.log('Positive: ' + sentimentCounts.Positive + ' (' + percentages.Positive + '%)');
    Logger.log('Neutral: ' + sentimentCounts.Neutral + ' (' + percentages.Neutral + '%)');
    Logger.log('Negative: ' + sentimentCounts.Negative + ' (' + percentages.Negative + '%)');
    
    // Check if negative sentiment percentage is 30% or higher
    if (parseFloat(percentages.Negative) >= 10) {
      sendAlert('High percentage of negative feedback detected. Negative feedback: ' + percentages.Negative + '%');
      Logger.log('Alert has been sent');
    }
  }
  
  
  
  function sendAlert(message) {
    var recipients = ['mrtkz0529@gmail.com']; 
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: 'EchoHR Alert',
      body: message
    });
  }
  
  // Step 8: Run all functions
  function runEchoHRSystem() {
    analyzeSentiment();
    createDashboard();
    checkForAlerts();
    Logger.log('EchoHR system updated successfully!');
  }
  
  // Set up a daily trigger for the EchoHR system
  function createDailyTrigger() {
    ScriptApp.newTrigger('runEchoHRSystem')
      .timeBased()
      .everyDays(1)
      .atHour(1)
      .create();
  }
  