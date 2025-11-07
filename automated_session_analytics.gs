/**
 * This script automates the analysis and reporting of chat history data.
 * 1. runChatAnalysis(): Processes the spreadsheet, saves metrics, and generates chart images.
 * 2. generateMonthlyReport(): Uses the saved metrics and charts to build a new Google Doc report.
 */

const SPREADSHEET_NAME_PREFIX = "Chat History ";
const MAIN_REPORT_FOLDER = "Monthly Report";
const PLOTS_SUBFOLDER = "Plots";
const REPORT_TEMPLATE_NAME = "Monthly Report Template";

// The main function to run the entire analysis and report generation process.
function runChatAnalysis() {
  Logger.log("Starting chat analysis process...");

  // Determine the target spreadsheet and sheet name for the previous month.
  const today = new Date();
  const targetDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const targetYear = targetDate.getFullYear();
  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const targetMonthName = monthNames[targetDate.getMonth()];

  Logger.log(`Targeting data for: ${targetMonthName} ${targetYear}`);

  const prevMonthDate = new Date(targetDate.getFullYear(), targetDate.getMonth() - 1, 1);
  const refMonthName = monthNames[prevMonthDate.getMonth()];
  const refYear = prevMonthDate.getFullYear();
  Logger.log(`Reference (previous) month will be: ${refMonthName} ${refYear}`);

  // Get and process the chat data from the Google Sheet.
  const analysisData = getAndProcessChatData(targetYear, targetMonthName);

  if (!analysisData) {
    Logger.log("Could not process data. Halting execution.");
    return;
  }

  const {
    processedData,
    totalQuestions,
    totalSessions
  } = analysisData;
  Logger.log(`Total questions analyzed: ${totalQuestions}`);
  Logger.log(`Total unique sessions: ${totalSessions}`);

  // Store the key metrics using PropertiesService for later retrieval.
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('totalQuestions', totalQuestions);
  scriptProperties.setProperty('totalSessions', totalSessions);
  scriptProperties.setProperty('targetMonthName', targetMonthName); 
  scriptProperties.setProperty('targetYear', targetYear);
  scriptProperties.setProperty('refMonthName', refMonthName);
  scriptProperties.setProperty('refYear', String(refYear));
  Logger.log("Key metrics have been stored.");


  // Get or create the folder in Google Drive for saving chart images.
  const mainFolder = getFolder(MAIN_REPORT_FOLDER);
  if (!mainFolder) {
      Logger.log(`Main folder "${MAIN_REPORT_FOLDER}" not found. Please create it.`);
      return;
  }
  const plotsFolder = getFolder(PLOTS_SUBFOLDER, mainFolder) || mainFolder.createFolder(PLOTS_SUBFOLDER);
  Logger.log(`Reports will be saved to folder: '${plotsFolder.getName()}'`);

  // Generate, save, and log URLs for all charts.
  const hourlyChart = createHourlySessionsChart(processedData);
  saveChartToDrive(hourlyChart, `hourly_sessions_${targetMonthName}.png`, plotsFolder);
  Logger.log(`Hourly Sessions Chart saved.`);

  const weekdayWeekendCharts = createWeekdayWeekendCharts(processedData);
  saveChartToDrive(weekdayWeekendCharts.weekdayChart, `weekday_sessions_${targetMonthName}.png`, plotsFolder);
  saveChartToDrive(weekdayWeekendCharts.weekendChart, `weekend_sessions_${targetMonthName}.png`, plotsFolder);
  Logger.log(`Weekday/Weekend Sessions Charts saved.`);

  const dailyChart = createDailySessionsChart(processedData);
  saveChartToDrive(dailyChart, `daily_sessions_${targetMonthName}.png`, plotsFolder);
  Logger.log(`Daily Sessions Chart saved.`);
  Logger.log("Chat analysis process completed successfully.");
}

function getFolder(name, parent) {
    const search = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
    if (search.hasNext()) {
        return search.next();
    }
    return null;
}

function getFile(name, parent) {
    const search = parent ? parent.getFilesByName(name) : DriveApp.getFilesByName(name);
    if (search.hasNext()) {
        return search.next();
    }
    return null;
}

function getSheetData(year, monthName) {
  try {
    const spreadsheetName = SPREADSHEET_NAME_PREFIX + year;
    const files = DriveApp.getFilesByName(spreadsheetName); 

    if (!files.hasNext()) {
      Logger.log(`Spreadsheet "${spreadsheetName}" not found in your Google Drive.`);
      return { sheet: null, data: null, headers: null };
    }

    const file = files.next();
    const spreadsheet = SpreadsheetApp.openById(file.getId());
    const sheet = spreadsheet.getSheetByName(monthName);

    if (!sheet) {
      Logger.log(`Sheet "${monthName}" not found in "${spreadsheetName}".`);
      return { sheet: null, data: null, headers: null };
    }
    const range = sheet.getDataRange();
    const values = range.getValues();
    const headers = values.shift(); // Remove header
    return { sheet, data: values, headers };

  } catch (e) {
    Logger.log(`Error accessing sheet ${monthName} for year ${year}: ${e.message}`);
    return { sheet: null, data: null, headers: null };
  }
}


function getAndProcessChatData(year, monthName) {
  try {
    const { sheet, data: rawData, headers } = getSheetData(year, monthName);

    if (!sheet || !rawData || !headers) {
      Logger.log(`Could not retrieve data for ${monthName} ${year}.`);
      return null;
    }

    const totalQuestions = rawData.length;

    // Find column indices to make the script robust to column order changes
    const sessionIDIdx = headers.indexOf('SessionID');
    const todayIdx = headers.indexOf('Today');
    const timeIdx = headers.indexOf('Time');

    if (sessionIDIdx === -1 || todayIdx === -1 || timeIdx === -1) {
        Logger.log("Error: One or more required columns (SessionID, Today, Time) are missing.");
        return null;
    }

    // Combine date and time, creating a JS Date object
    let dataWithDatetime = rawData.map(row => {
      const date = new Date(row[todayIdx]);
      const time = new Date(row[timeIdx]);
      
      // Set the time from the 'Time' column onto the date from the 'Today' column
      date.setHours(time.getHours(), time.getMinutes(), time.getSeconds());
      
      return {
        sessionID: row[sessionIDIdx],
        datetime: date,
      };
    });

    // Keep only the first occurrence for each unique SessionID
    // Sorting not needed as the source data is already chronological
    const uniqueSessions = [];
    const seenSessionIDs = {};
    dataWithDatetime.forEach(row => {
      if (!seenSessionIDs[row.sessionID]) {
        const weekdayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
        const dayIndex = row.datetime.getDay();
        const isWeekend = (dayIndex === 0 || dayIndex === 6);

        uniqueSessions.push({
          ...row,
          hour: row.datetime.getHours(),
          weekday: weekdayNames[dayIndex],
          isWeekend: isWeekend ? 'Weekend' : 'Weekday'
        });
        seenSessionIDs[row.sessionID] = true;
      }
    });

    return {
      processedData: uniqueSessions,
      totalQuestions: totalQuestions,
      totalSessions: uniqueSessions.length,
    };

  } catch (e) {
    Logger.log(`An error occurred: ${e.message}`);
    return null;
  }
}

function groupByAndCount(data, key) {
  return data.reduce((acc, obj) => {
    const group = obj[key];
    acc[group] = (acc[group] || 0) + 1;
    return acc;
  }, {});
}

function createHourlySessionsChart(data) {
  const hourlyCounts = groupByAndCount(data, 'hour');
  PropertiesService.getScriptProperties().setProperty('hourlyCounts', JSON.stringify(hourlyCounts));
  
  const dataTable = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, 'Hour').addColumn(Charts.ColumnType.NUMBER, 'Sessions');
  for (let hour = 0; hour < 24; hour++) {
    dataTable.addRow([String(hour), hourlyCounts[hour] || 0]);
  }
  return Charts.newColumnChart()
    .setDataTable(dataTable)
    .setTitle('Number of Sessions per Hour')
    .setXAxisTitle('Hour of Day')
    .setYAxisTitle('Number of Sessions')
    .setOption('height', 400)
    .setOption('width', 800)
    .build();
}

function createWeekdayWeekendCharts(data) {
    const weekdayData = data.filter(d => d.isWeekend === 'Weekday');
    const weekendData = data.filter(d => d.isWeekend === 'Weekend');

    // --- NEW: Save data for the report ---
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('weekdayTotal', weekdayData.length);
    scriptProperties.setProperty('weekendTotal', weekendData.length);

    const weekdayCounts = groupByAndCount(weekdayData, 'hour');
    const weekdayTable = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, 'Hour').addColumn(Charts.ColumnType.NUMBER, 'Sessions');
    for (let hour = 0; hour < 24; hour++) {
        weekdayTable.addRow([String(hour), weekdayCounts[hour] || 0]);
    }
    const weekdayChart = Charts.newColumnChart()
        .setDataTable(weekdayTable)
        .setTitle('Weekday Sessions')
        .setXAxisTitle('Hour of Day')
        .setYAxisTitle('Number of Sessions')
        .build();

    const weekendCounts = groupByAndCount(weekendData, 'hour');
    const weekendTable = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, 'Hour').addColumn(Charts.ColumnType.NUMBER, 'Sessions');
    for (let hour = 0; hour < 24; hour++) {
        weekendTable.addRow([String(hour), weekendCounts[hour] || 0]);
    }
    const weekendChart = Charts.newColumnChart()
        .setDataTable(weekendTable)
        .setTitle('Weekend Sessions')
        .setXAxisTitle('Hour of Day')
        .setYAxisTitle('Number of Sessions')
        .build();
        
    return { weekdayChart, weekendChart };
}

function createDailySessionsChart(data) {
  const dailyCounts = groupByAndCount(data, 'weekday');
  PropertiesService.getScriptProperties().setProperty('dailyCounts', JSON.stringify(dailyCounts));
  
  const dataTable = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, 'Day').addColumn(Charts.ColumnType.NUMBER, 'Sessions');
  const dayOrder = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  dayOrder.forEach(day => {
    dataTable.addRow([day, dailyCounts[day] || 0]);
  });
  return Charts.newColumnChart()
    .setDataTable(dataTable)
    .setTitle('Number of Sessions per Day of Week')
    .setXAxisTitle('Day of Week')
    .setYAxisTitle('Number of Sessions')
    .setOption('height', 400)
    .setOption('width', 800)
    .build();
}

/**
 * Saves a chart object as a PNG file in a specified Google Drive folder.
 * @param {GoogleAppsScript.Charts.Chart} chart The chart to save.
 * @param {string} fileName The desired file name for the image.
 * @param {GoogleAppsScript.Drive.Folder} folder The Drive folder to save the file in.
 */
function saveChartToDrive(chart, fileName, folder) {
  // Clear out old file if it exists
  const oldFiles = folder.getFilesByName(fileName);
  if (oldFiles.hasNext()) {
    oldFiles.next().setTrashed(true);
  }
  
  // Create the new file
  const chartBlob = chart.getAs('image/png').setName(fileName);
  folder.createFile(chartBlob);
}


/**
 * MONTHLY REPORT GENERATION
 * Main function to generate the monthly report.
 * Run this after runChatAnalysis() has completed.
 */
function generateMonthlyReport() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const props = scriptProperties.getProperties();

  // Get all stored data from PropertiesService
  const current = {
    monthName: props.targetMonthName,
    year: parseInt(props.targetYear, 10), // Parse year to integer
    totalSessions: parseInt(props.totalSessions, 10),
    totalQuestions: parseInt(props.totalQuestions, 10),
    weekdayTotal: parseInt(props.weekdayTotal, 10),
    weekendTotal: parseInt(props.weekendTotal, 10),
    dailyCounts: JSON.parse(props.dailyCounts || '{}')
  };
  
  // Check if data is available
  if (!current.monthName) {
    Logger.log("ERROR: No analysis data found. Please run 'runChatAnalysis' first.");
    return;
  }
  
  const currentMonthYear = `${current.monthName} ${current.year}`;
  const currentMonthPadded = String(new Date(currentMonthYear + ' 1, 2000').getMonth() + 1);
  const newReportName = `${currentMonthPadded}/${current.year} Monthly Report`;
  
  Logger.log(`Generating report: ${newReportName}`);

  // Read the reference month details saved by runChatAnalysis()
  const refMonthName = props.refMonthName;
  const refYear = parseInt(props.refYear, 10);
  
  if (!refMonthName || !refYear) {
      Logger.log("ERROR: Reference month (refMonthName, refYear) not found in Properties. Please re-run 'runChatAnalysis'.");
      return;
  }

  Logger.log(`Fetching reference data for: ${refMonthName} ${refYear}`);
  
  const previous = { sessions: 0, questions: 0 };
  
  // Get the raw data for the previous month (e.g., September)
  const prevChatData = getSheetData(refYear, refMonthName);
  
  if (prevChatData.data && prevChatData.headers) {
    
    // get previous questions (total rows, as header is already removed by getSheetData)
    previous.questions = prevChatData.data.length; 
    
    // Calculate previous sessions (unique SessionIDs) from that same data
    const sessionIDIdx = prevChatData.headers.indexOf('SessionID');
    
    if (sessionIDIdx === -1) {
        Logger.log(`Warning: 'SessionID' column not found in ${refMonthName} data. Cannot count sessions.`);
    } else {
        const seenSessionIDs = {};
        let uniqueSessionCount = 0;
        
        // Loop through all rows and count unique IDs
        prevChatData.data.forEach(row => {
            const sessionID = row[sessionIDIdx];
            if (sessionID && !seenSessionIDs[sessionID]) {
                uniqueSessionCount++;
                seenSessionIDs[sessionID] = true;
            }
        });
        previous.sessions = uniqueSessionCount;
    }
  } else {
    Logger.log(`Warning: Could not get previous month's data for ${refMonthName} ${refYear}. Setting stats to 0.`);
  }
  
  Logger.log(`Found reference data: ${previous.sessions} sessions, ${previous.questions} questions.`);
  

  // mainReportFolder for other operations
  const mainReportFolder = getFolder(MAIN_REPORT_FOLDER);
  if (!mainReportFolder) {
      Logger.log(`ERROR: Main report folder "${MAIN_REPORT_FOLDER}" not found.`);
      return;
  }
  // Find Report Template
  const templateFile = getFile(REPORT_TEMPLATE_NAME, mainReportFolder);
  if (!templateFile) {
    Logger.log(`ERROR: Report template "${REPORT_TEMPLATE_NAME}" not found in "${MAIN_REPORT_FOLDER}".`);
    return;
  }
  
  // Create New Report from Template
  // Delete old report if it exists
  const oldReport = getFile(newReportName, mainReportFolder);
  if(oldReport) {
    oldReport.setTrashed(true);
    Logger.log(`Removed old report: "${newReportName}"`);
  }
  
  const newReportFile = templateFile.makeCopy(newReportName, mainReportFolder);
  const newReportDoc = DocumentApp.openById(newReportFile.getId());
  const body = newReportDoc.getBody();
  
  // Generate AI Text
  Logger.log("Calling OpenAI to generate report text...");
  const apiKey = scriptProperties.getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    Logger.log("ERROR: 'OPENAI_API_KEY' not found in Script Properties. Please add it.");
    return;
  }
  
  const aiText = generateAiReportText(current, previous, apiKey);
  if (!aiText) {
    Logger.log("ERROR: Failed to generate AI text.");
    return;
  }
  Logger.log("AI text generated successfully.");
  
  // Populate Document with Text
  body.replaceText('{{current_month_year}}', currentMonthYear);
  body.replaceText('{{session_interaction_count_text}}', aiText.sessionText);
  
  // Find and Insert Plots
  Logger.log("Inserting charts into report...");
  const plotsFolder = getFolder(PLOTS_SUBFOLDER, mainReportFolder);
  if (!plotsFolder) {
    Logger.log(`ERROR: Plots folder "${PLOTS_SUBFOLDER}" not found in "${MAIN_REPORT_FOLDER}".`);
    return;
  }
  
  // Get chart images
  const dailyChartBlob = getFile(`daily_sessions_${current.monthName}.png`, plotsFolder).getBlob();
  const weekdayChartBlob = getFile(`weekday_sessions_${current.monthName}.png`, plotsFolder).getBlob();
  const weekendChartBlob = getFile(`weekend_sessions_${current.monthName}.png`, plotsFolder).getBlob();
  const hourlyChartBlob = getFile(`hourly_sessions_${current.monthName}.png`, plotsFolder).getBlob();
  
  // Replace placeholders with images
  replaceTextWithImage(body, '{{daily_chart}}', dailyChartBlob);
  const weekdayPlaceholder = body.findText('{{weekday_weekend_charts}}'); 
  if (weekdayPlaceholder) {
    const element = weekdayPlaceholder.getElement();
    const parent = element.getParent().asParagraph();
    
    // Clear the placeholder text without clearing the whole paragraph
    element.asText().setText(''); 
    
    parent.insertInlineImage(0, weekendChartBlob).setWidth(500);
    parent.insertText(0, '   '); 
    parent.insertInlineImage(0, weekdayChartBlob).setWidth(500);
    
    // Clean up the unused end tag
    body.replaceText('{{weekday_weekend_charts_end}}', '');
  } else {
      Logger.log("Warning: Could not find placeholder '{{weekday_weekend_charts}}'");
      // Also clean up the end tag in case it's orphaned
      body.replaceText('{{weekday_weekend_charts_end}}', '');
  }

  replaceTextWithImage(body, '{{hourly_chart}}', hourlyChartBlob);
  
  newReportDoc.saveAndClose();
  Logger.log(`SUCCESS: Report generated: ${newReportFile.getUrl()}`);
}


function getDaylightStats(dailyCounts) {
    let peakCount = 0;
    let minCount = Infinity;
    let peakDays = [];
    let minDay = '';
    const dayOrder = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

    dayOrder.forEach(day => {
        const count = dailyCounts[day] || 0;
        
        if (count > peakCount) {
            peakCount = count;
            peakDays = [day];
        } else if (count === peakCount) {
            peakDays.push(day);
        }
        
        if (count < minCount) {
            minCount = count;
            minDay = day;
        }
    });

    return {
        peakDays: peakDays.join(' and '),
        peakCount: peakCount,
        minDay: minDay,
        minCount: minCount
    };
}


/**
 * Calls OpenAI gpt-4o-mini to generate the report's text content.
 * @param {object} current The current month's data.
 * @param {object} previous The previous month's data.
 * @param {string} apiKey The OpenAI API key.
 * @returns {object|null} An object with {sessionText, useCaseText} or null.
 */
function generateAiReportText(current, previous, apiKey) {
  const dayStats = getDaylightStats(current.dailyCounts);
  const currentMonthYear = `${current.monthName} ${current.year}`;

  const prompt = `You are a professional analyst writing a monthly user engagement report for a university chatbot named Kingbot.
The report is for ${currentMonthYear}.
Write two paragraphs for the 'Session & Interaction Count' section. Be insightful and professional. Use ALL of the following data:
- Current Month: ${current.totalSessions} sessions and ${current.totalQuestions} questions.
- Previous Month: ${previous.sessions} sessions and ${previous.questions} questions. (Comment on the change).
- Weekday sessions: ${current.weekdayTotal}. Weekend sessions: ${current.weekendTotal}. (Comment on this trend).
- Peak Day(s): ${dayStats.peakDays} (${dayStats.peakCount} sessions).
- Low Day: ${dayStats.minDay} (${dayStats.minCount} sessions).

IMPORTANT: Respond ONLY with the text for the report, split by a special marker.
Start with SESSION_TEXT:: followed by the first two paragraphs.
Do not include any other text, pleasantries, or markdown.`;

  const url = "https://api.openai.com/v1/chat/completions";
  const payload = {
    model: "gpt-4o-mini",
    messages: [{ role: "user", content: prompt }],
    max_tokens: 1000,
    temperature: 0.7
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + apiKey
    },
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    const text = json.choices[0].message.content;

    const sessionText = text.split("SESSION_TEXT::")[1].trim();
    return { sessionText };
    
  } catch (e) {
    Logger.log(`OpenAI API Error: ${e.message}`);
    return null;
  }
}

function replaceTextWithImage(body, placeholder, imageBlob, width = 600) {
  const searchResult = body.findText(placeholder);
  if (searchResult) {
    const element = searchResult.getElement();
    const parent = element.getParent();
    
    if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
      parent.asParagraph().clear(); // Clear the placeholder text
      const image = parent.asParagraph().insertInlineImage(0, imageBlob);
      
      // Resize image
      const aspectRatio = image.getHeight() / image.getWidth();
      image.setWidth(width);
      image.setHeight(width * aspectRatio);
    }
  } else {
    Logger.log(`Warning: Could not find placeholder "${placeholder}" in the document.`);
  }
}
