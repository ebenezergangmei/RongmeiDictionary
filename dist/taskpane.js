Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    console.log("Rongmei Dictionary Add-in is ready!");
    setupSelectionHandler();
  }
});

const dictionaryUrl = "https://ebenezergangmei.github.io/RongmeiDictionary/dist/dictionary.json";
let dictionaryData = [];

// Load the dictionary from GitHub Pages
async function loadDictionary() {
  try {
    const response = await fetch(dictionaryUrl);
    if (!response.ok) throw new Error("Failed to load dictionary");
    dictionaryData = await response.json();
    console.log("Dictionary loaded successfully");
  } catch (err) {
    console.error("Error loading dictionary:", err);
    document.getElementById("output").innerHTML = "❌ Could not load dictionary data.";
  }
}

// When user selects text
function setupSelectionHandler() {
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, async () => {
    await showMeaningOfSelection();
  });
}

// Show meaning for the selected text
async function showMeaningOfSelection() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const selectedText = selection.text.trim();
    const outputDiv = document.getElementById("output");

    if (!selectedText) return;

    // Clear previous results
    outputDiv.innerHTML = "";

    const words = selectedText.split(/\s+/);
    let results = [];

    for (let word of words) {
      const match = dictionaryData.find(entry => entry.English.toLowerCase() === word.toLowerCase());
      if (match) {
        results.push(match);
      }
    }

    if (results.length > 0) {
      outputDiv.innerHTML = results.map((entry, i) =>
        `<div><b>${i + 1}. ${entry.English}</b>: ${entry.Rongmei}</div>`
      ).join("<hr>");
    } else {
      outputDiv.innerHTML = `No meaning found for: <b>${selectedText}</b>`;
    }

    await context.sync();
  });
}

// Load dictionary on startup
loadDictionary();
