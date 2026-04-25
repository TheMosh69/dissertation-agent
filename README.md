# Dissertation Agent — Word Add-in

AI-powered dissertation writing assistant. Runs as a sidebar in Microsoft Word.
Writes introductions with live literature search, and extends/interprets results & discussion sections from your text and reaction scheme images.

---

## What you need

- Microsoft Word (Microsoft 365 subscription, Version 2205+ on Windows or 16.61+ on Mac)
- An Anthropic API key: https://console.anthropic.com/settings/keys
- A free GitHub account (for hosting the taskpane — takes 5 minutes)

---

## Step 1 — Host the taskpane on GitHub Pages

GitHub Pages gives you free HTTPS hosting, which Word requires for add-ins.

1. Go to https://github.com and create a new **public** repository called `dissertation-agent`
2. Upload `taskpane.html` to that repository
3. Go to **Settings → Pages** and set Source to **Deploy from branch: main / root**
4. After ~60 seconds, your taskpane will be live at:
   `https://YOUR-GITHUB-USERNAME.github.io/dissertation-agent/taskpane.html`

---

## Step 2 — Update the manifest

Open `manifest.xml` and replace every occurrence of:

    YOUR-GITHUB-USERNAME

with your actual GitHub username (4 replacements total).

Example — if your username is `dado-chem`:

    https://dado-chem.github.io/dissertation-agent/taskpane.html

---

## Step 3 — Sideload the add-in in Word

### On Mac
1. Open Word
2. Go to **Tools → Add-ins → My Add-ins → Upload My Add-in**
3. Select `manifest.xml`

### On Windows
1. Open Word
2. Go to **Insert → Get Add-ins → My Add-ins → Upload My Add-in**
3. Select `manifest.xml`

### On Word for the Web
1. Go to **Insert → Add-ins → Upload My Add-in**
2. Select `manifest.xml`

Once loaded, a **Dissertation Agent** button appears in the **Home** ribbon.

---

## Step 4 — Use it

1. Click **Dissertation Agent** in the ribbon to open the sidebar
2. Enter your Anthropic API key (stored locally in your browser only — never sent anywhere except Anthropic's API)
3. Use the **Introduction** tab to generate a literature-backed introduction
4. Use the **Results & Discussion** tab to extend your existing text and interpret your reaction schemes
5. Hit **Insert into Word** to insert the generated text at your cursor position

### Word integration features
- **Read selection** — select any text in your document and click "Read selection" to pull it into the input field
- **Insert into Word** — inserts generated text directly at your current cursor position
- **Refine** — ask follow-up changes in the same session ("translate to German", "make it more concise")

---

## Troubleshooting

**"Insert into Word" is greyed out** — make sure you opened the add-in from within Word (not a browser tab)

**API errors** — check your API key at https://console.anthropic.com and ensure you have credits

**Add-in doesn't appear after sideloading** — clear Word's add-in cache:
  - Mac: delete `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`
  - Windows: delete `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`

**GitHub Pages not live yet** — wait 1-2 minutes after enabling Pages, then hard-refresh
