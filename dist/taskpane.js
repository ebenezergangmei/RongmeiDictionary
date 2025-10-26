let dictionary = {};
let lastSelection = "";

async function loadDictionary() {
  const localJson = "dictionary.json"; // fallback
  const githubJson = "https://ebenezergangmei.github.io/RongmeiDictionary/dist/dictionary.json";

  try {
    // Try fetching from GitHub
    const response = await fetch(githubJson);
    dictionary = await response.json();
    document.getElementById("output").innerHTML =
      "<b>Dictionary loaded from GitHub.</b><br>Select a word or phrase to begin.";
    console.log("✅ Dictionary loaded from GitHub with", Object.keys(dictionary).length, "entries");
  } catch (err) {
    console.warn("⚠️ GitHub fetch failed, loading local JSON", err);
    // Fallback to local copy
    try {
      const responseLocal = await fetch(localJson);
      dictionary = await responseLocal.json();
      document.getElementById("output").innerHTML =
        "<b>Dictionary loaded from local file.</b><br>Select a word or phrase to begin.";
      console.log("✅ Dictionary loaded locally with", Object.keys(dictionary).length, "entries");
    } catch (err2) {
      console.error("❌ Could not load dictionary:", err2);
      document.getElementById("output").textContent = "Failed to load dictionary file.";
    }
  }
}

Office.onReady(() => {
  loadDictionary();

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    onSelectionChange
  );
});

async function onSelectionChange() {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const selectedText = selection.text.trim();
      const output = document.getElementById("output");

      if (!selectedText) return;
      if (selectedText === lastSelection) return; // same selection, skip

      lastSelection = selectedText;

      const words = selectedText.match(/\b[a-zA-Z\.]+\b/g);
      if (!words) {
        output.innerHTML = "No valid words selected.";
        return;
      }

      let html = `<div style="margin-bottom:10px;"><b>Selected:</b> ${selectedText}<hr>`;
      let count = 1;

      for (const w of words) {
        const meaning = dictionary[w.toLowerCase()];
        html += `<div>${count}. <b>${w}</b>: ${meaning ? meaning : "❌ Not found"}</div>`;
        count++;
      }

      html += "</div>";

      output.innerHTML = html;

    });
  } catch (err) {
    console.error("Error:", err);
  }
}
