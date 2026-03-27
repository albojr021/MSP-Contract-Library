const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'); 

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('MSP Notarized Contract Library')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getContractData() {
  const sheetId = '15UV9elecbiNpJ9DbB1PVKHNfZlzFsvm47wifpUK0fhU';
  const tabName = 'RefSeries';
  const startRow = 749;
  
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(tabName);
    
    if (!sheet) throw new Error("Sheet tab 'RefSeries' was not found in the document.");
    
    const lastRow = sheet.getLastRow();
    if (lastRow < startRow) return []; 
    
    const numRows = lastRow - startRow + 1;
    const dataRange = sheet.getRange(startRow, 1, numRows, 43); 
    const data = dataRange.getValues();
    
    const validContracts = [];
    const urlRegex = /^https?:\/\//i; 
    
    data.forEach((row) => {
      const sfcUrl = String(row[18]).trim(); // Col S
      const mlcUrl = String(row[32]).trim(); // Col AG
      const noaUrl = String(row[42]).trim(); // Col AQ
      
      const isValidSfc = urlRegex.test(sfcUrl);
      const isValidMlc = urlRegex.test(mlcUrl);
      const isValidNoa = urlRegex.test(noaUrl);
      
      if (isValidSfc || isValidMlc || isValidNoa) {
        validContracts.push({
          property: String(row[3]).trim(),          // Col D
          contractGrpId: String(row[4]).trim(),     // Col E 
          status: String(row[5]).trim(),            // Col F
          payor: String(row[6]).trim(),             // Col G
          supplier: String(row[7]).trim(),          // Col H
          headcount: String(row[8]).trim(),         // Col I
          kindOfService: String(row[9]).trim(),     // Col J
          startDate: formatDate(row[10]),           // Col K
          endDate: formatDate(row[11]),             // Col L
          sector: String(row[12]).trim(),           // Col M
          refNum: String(row[14]).trim(),           // Col O
          kindOfSfc: String(row[15]).trim(),        // Col P
          sfcUrl: isValidSfc ? sfcUrl : null,       // Col S
          sfcBallWith: String(row[27]).trim(),      // Col AB
          mlcUrl: isValidMlc ? mlcUrl : null,       // Col AG
          mlcBallWith: String(row[37]).trim(),      // Col AL
          noaUrl: isValidNoa ? noaUrl : null        // Col AQ
        });
      }
    });
    
    return JSON.stringify(validContracts);
    
  } catch (error) {
    return JSON.stringify({ error: "Failed to fetch contract data: " + error.toString() });
  }
}

function formatDate(dateString) {
  if (!dateString) return "";
  const date = new Date(dateString);
  if (isNaN(date.getTime())) return String(dateString); 
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "MMM dd, yyyy");
}

function askGeminiAssistant(userMessage, activeContextString) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  
  const systemPrompt = `You are a helpful AI assistant built into the 'MSP Notarized Contract Library' web application.
  
  You have TWO main responsibilities:

  1. APP NAVIGATION GUIDE:
  Teach the user how to use the app (e.g., search, property filter, status tabs, document buttons). If they ask "Saan ko makikita ang graph?", tell them to click the 'View Analytics Dashboard' button.

  2. DATA REPORTER (STRICT ACCURACY):
  Answer data questions using ONLY the 'ANALYTICS_DATA' JSON object provided below. 
  CRITICAL RULE: Do NOT calculate totals or sums yourself. The system has already done the exact math based on current filters. Do NOT manually count items in 'SAMPLE_DATA'.

  SMART MATCHING INSTRUCTION:
  Users will use partial names, acronyms, or informal terms (e.g., "aim one" for "AIM ONE MANPOWER", or "COG" for "Corporate Operations Group"). You must intelligently match their partial queries to the closest exact keys in the JSON data.

  Look for answers in ANALYTICS_DATA keys based on the question:
  - If asked about "total headcount", "total na tao", "manpower per property", or "Top Properties by Headcount", look in the 'headcountPerProperty' data.
  - If asked about "total headcount", "total na tao", or "manpower per agency", "provider", or "supplier", look in the 'headcountPerProvider' data.
  - If asked about "total headcount", "total na tao", or "manpower per kind of service", look in the 'headcountPerService' data.
  - If asked for the GRAND TOTAL of all manpower/tao, sum up the values in 'headcountPerProperty' ONLY IF a total isn't explicitly provided, or refer to the total in the data if available.
  - If asked about "Status breakdown" or "Live/Expired contracts", look in 'statusBreakdown'.
  - If asked about "Contract Volume per Property", look in 'contractVolumePerProperty'.
  - If asked about headcount for a specific service WITHIN a specific property (e.g. "ilang tao sa COG under administrative service"), look in the 'headcountPerPropertyAndService' nested data. Find the property, then find the service inside it.
  - Headcount for a specific Service inside a specific AGENCY/PROVIDER (e.g., "Aim one under admin service"): look in 'headcountPerProviderAndService'. Match the closest agency name, then the service.

  The data is dynamic. It reflects what the user sees on screen based on their active search and property filters. Always reply politely, conversationally, and directly in the user's language (English or Tagalog). If they ask "Ilan ang total na tao sa [Property/Agency/Service]?", give them the exact number from the JSON.

  3. ADMIN CONTACT:
  If you can't answer the user's Queries give them the contact of the admin for more information: "mcdmarketingstorage@megaworld-lifestyle.com" and "jdmorelos@megaworld-lifestyle.com".

  SYSTEM DATA (Based on current filters):
  ${activeContextString}
  `;

  const payload = {
    contents: [
      {
        role: "user",
        parts: [{ text: systemPrompt + "\n\nUser Question: " + userMessage }]
      }
    ]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (result.error) {
      throw new Error(result.error.message);
    }

    return result.candidates[0].content.parts[0].text;
  } catch (error) {
    return "I am currently unable to analyze the data or guide you. Error: " + error.message;
  }
}
