function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Tech VA Tools")
    .addItem("Generate AI Suggestions (Single)", "generateAllSuggestions")
    .addItem("Generate Global Priorities", "generateGlobalPriorities")
    .addToUi();
}


// ------------------------------
// GLOBAL PRIORITY SYSTEM
// ------------------------------
function generateGlobalPriorities() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const apiKey = "API KEY HERE";

  let tasks = [];

  // Collect tasks
  for (let i = 1; i < data.length; i++) {
    const taskName = data[i][0];
    const category = data[i][1];
    const context = data[i][2];
    const aiCell = data[i][5];

    if (context && !aiCell) {
      tasks.push({
        row: i + 1,
        task: taskName,
        category: category,
        context: context
      });
    }
  }

  if (tasks.length === 0) return;

  const prompt =
`You are a Tech VA for a dental clinic.

Rank all tasks by priority considering urgency and patient impact.

You MUST follow these rules:

HIGH PRIORITY:
- patient follow-ups
- missed appointments
- scheduling issues
- post-treatment care
- booking conflicts

MEDIUM PRIORITY:
- patient inquiries
- admin updates
- insurance processing
- internal coordination

LOW PRIORITY:
- marketing content
- social media posts
- newsletters
- promotions
- branding tasks

Return ONLY valid JSON (no markdown, no code blocks):

[
  {
    "row": number,
    "priority": "High|Medium|Low",
    "suggestion": "one clear, specific, operational instruction written like a real dental clinic assistant (include timing, channel, or method when relevant, but keep it concise)"
  }
]

TASKS:
${JSON.stringify(tasks)}`;

  const url = "https://api.groq.com/openai/v1/chat/completions";

  const payload = {
    model: "llama-3.1-8b-instant",
    messages: [{ role: "user", content: prompt }],
    temperature: 0.2,
    max_tokens: 2000
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    let aiText = result.choices[0].message.content;

    // Clean markdown just in case
    const cleanText = aiText
      .replace(/```json/g, "")
      .replace(/```/g, "")
      .trim();

    let parsed;

try {
  parsed = JSON.parse(cleanText);
} catch (e) {
  Logger.log("RAW AI OUTPUT:");
  Logger.log(cleanText);

  throw new Error("AI returned invalid JSON. Check logs.");
}

    parsed.forEach(item => {

      const sheetRow = item.row;

      // ✅ PROTECT HEADER ROW
      if (!sheetRow || sheetRow < 2) return;

      sheet.getRange(sheetRow, 4).setValue(item.priority);   // Column D
      sheet.getRange(sheetRow, 6).setValue(item.suggestion);  // Column F
    });

  } catch (err) {
    Logger.log("ERROR: " + err.message);
  }
}
