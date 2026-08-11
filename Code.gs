function generatePrompt() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var data = sheet.getDataRange().getValues();

  var lastRow = data[data.length - 1];

  var task = lastRow[1];
  var topic = lastRow[2];
  var requirement = lastRow[3];
  var audience = lastRow[4];
  var tone = lastRow[5];
  var language = lastRow[6];
  var aiModel = lastRow[7];
  var length = lastRow[8];
  var extraInstructions = lastRow[9];

  Logger.log(task);
  Logger.log(topic);
  Logger.log(requirement);
  Logger.log(audience);
  Logger.log(tone);
  Logger.log(language);
  Logger.log(aiModel);
  Logger.log(length);
  Logger.log(extraInstructions);

  var prompt = "";

if (task == "Image") {

  prompt =
    "Create a detailed AI image prompt about " + topic +
    ". Audience: " + audience +
    ". Style: " + tone +
    ". Language: " + language +
    ". Requirements: " + requirement +
    ". Extra Instructions: " + extraInstructions;

}
else if (task == "💻 Code Generation") {

  prompt =
    "Generate code for " + topic +
    ". Requirement: " + requirement +
    ". Target Audience: " + audience +
    ". Programming Language: " + language +
    ". Output Length: " + length +
    ". Extra Instructions: " + extraInstructions;

}
else if (task == "📊 PowerPoint Presentation") {

  prompt =
    "Create a professional PowerPoint presentation about " + topic +
    ". Target Audience: " + audience +
    ". Tone: " + tone +
    ". Language: " + language +
    ". Requirements: " + requirement +
    ". Number of slides or length: " + length +
    ". Extra Instructions: " + extraInstructions;

}
else if (task == "📝 Resume") {

  prompt =
    "Create a professional and ATS-friendly resume for " + topic +
    ". Target Audience: " + audience +
    ". Tone: " + tone +
    ". Language: " + language +
    ". Requirements: " + requirement +
    ". Length: " + length +
    ". Extra Instructions: " + extraInstructions;

}

else if (task == "📧 Email Draft") {

  prompt =
    "Write a professional email about " + topic +
    ". Target Audience: " + audience +
    ". Tone: " + tone +
    ". Language: " + language +
    ". Purpose and Requirements: " + requirement +
    ". Length: " + length +
    ". Extra Instructions: " + extraInstructions;

}

else if (task == "📄 Report Writing") {

  prompt =
    "Write a well-structured report about " + topic +
    ". Target Audience: " + audience +
    ". Tone: " + tone +
    ". Language: " + language +
    ". Requirements: " + requirement +
    ". Length: " + length +
    ". Extra Instructions: " + extraInstructions;

}

else if (task == "🌐 Website") {

  prompt =
    "Design and develop a responsive website about " + topic +
    ". Target Audience: " + audience +
    ". Tone and design style: " + tone +
    ". Language: " + language +
    ". Requirements: " + requirement +
    ". Output Length: " + length +
    ". Extra Instructions: " + extraInstructions;

}

else if (task == "📱 Social Media Post") {

  prompt =
    "Create an engaging social media post about " + topic +
    ". Target Audience: " + audience +
    ". Tone: " + tone +
    ". Language: " + language +
    ". Requirements: " + requirement +
    ". Length: " + length +
    ". Extra Instructions: " + extraInstructions;

}

else if (task == "🎓 Assignment") {

  prompt =
    "Create a well-structured academic assignment about " + topic +
    ". Target Audience: " + audience +
    ". Tone: " + tone +
    ". Language: " + language +
    ". Requirements: " + requirement +
    ". Length: " + length +
    ". Extra Instructions: " + extraInstructions;

}

else if (task == "🧮 SQL Query") {

  prompt =
    "Generate an accurate SQL query for " + topic +
    ". Requirements: " + requirement +
    ". Target Audience: " + audience +
    ". SQL dialect or Language: " + language +
    ". Output Length: " + length +
    ". Extra Instructions: " + extraInstructions;

}

else if (task == "📈 Data Analysis") {

  prompt =
    "Perform a detailed data analysis related to " + topic +
    ". Target Audience: " + audience +
    ". Tone: " + tone +
    ". Language: " + language +
    ". Requirements: " + requirement +
    ". Output Length: " + length +
    ". Extra Instructions: " + extraInstructions;

}

else if (task == "🎤 Speech") {

  prompt =
    "Write an engaging speech about " + topic +
    ". Target Audience: " + audience +
    ". Tone: " + tone +
    ". Language: " + language +
    ". Requirements: " + requirement +
    ". Length: " + length +
    ". Extra Instructions: " + extraInstructions;

}

else if (task == "📖 Story") {

  prompt =
    "Write a creative and engaging story about " + topic +
    ". Target Audience: " + audience +
    ". Tone: " + tone +
    ". Language: " + language +
    ". Requirements: " + requirement +
    ". Length: " + length +
    ". Extra Instructions: " + extraInstructions;

}

else if (task == "🤖 AI Prompt") {

  prompt =
    "Create a clear, detailed and optimized AI prompt for " + topic +
    ". Target Audience: " + audience +
    ". Desired Tone: " + tone +
    ". Language: " + language +
    ". Requirements: " + requirement +
    ". Output Length: " + length +
    ". Extra Instructions: " + extraInstructions;

}

else {

  prompt =
    "Create a high-quality response about " + topic +
    ". Target Audience: " + audience +
    ". Tone: " + tone +
    ". Language: " + language +
    ". Requirements: " + requirement +
    ". Output Length: " + length +
    ". Extra Instructions: " + extraInstructions;

}
Logger.log(prompt);
sheet.getRange(data.length, 11).setValue(prompt);

}
