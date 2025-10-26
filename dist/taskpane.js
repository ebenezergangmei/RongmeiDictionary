let dictionary = [];
let allResults = [];

const githubJson = "https://ebenezergangmei.github.io/RongmeiDictionary/dist/dictionary.json";

async function loadDictionary() {
  try {
    const response = await fetch(githubJson);
    dictionary = await response.json();
    document.getElementById("output").innerHTML = "Dictionary loaded from GitHub. Select a word or phrase.";
  } catch (err) {
    console.error("Failed to load GitHub dictionary:", err);
    document.getElementById("output").innerHTML = "⚠️ Could not load GitHub dictionary.";
  }
}

// Call this on page load
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    loadDictionary();
  }
});

// Show meanings for selected words
async function showMeanings(words) {
  words.forEach(word => {
    const entry = dictionary.find(e => e.English.toLowerCase() === word.toLowerCase());
    if (entry) allResults.push(entry);
  });

  document.getElementById("output").innerHTML = allResults.map((e,i) =>
    `<b>${i+1}. ${e.English}</b>: ${e.Rongmei}`
  ).join("<br><br>");
}

// Get selected text in Word and show meanings
async function getSelection() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();
    const words = selection.text.split(/\s+/);
    showMeanings(words);
  });
}
