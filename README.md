# PowerPoint Style Standardizer Add-in

A professional Office.js Task Pane add-in designed to scan PowerPoint presentations and apply standardized formatting (font family and sizes) based on customizable thresholds for Headlines, Sub-Headlines, and Body text.

## 🚀 Deployment (GitHub Pages)

1. **Host Files:** Ensure `taskpane.html`, `taskpane.js`, and `manifest.xml` are in the root of this repository.
2. **Enable Hosting:**
* Go to **Settings > Pages**.
* Set Source to **Deploy from a branch** (usually `main`).


3. **Update Manifest:** Ensure the `<SourceLocation>` in your `manifest.xml` matches your GitHub Pages URL:
`https://<your-username>.github.io/my-ppt-formatter/taskpane.html`

## 🍎 Installation (macOS Sideloading)

To install the add-in on an Apple Silicon Mac without using the Office Store:

1. Open **Finder** and press `Cmd + Shift + G`.
2. Paste the following path:
`~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`
*(If the `wef` folder does not exist, create it).*
3. Copy the `manifest.xml` file into this folder.
4. **Restart PowerPoint**.
5. Go to **Insert > My Add-ins** and look for **Style Standardizer**.

## 🛠 Features

* **Range-Based Logic:** Identify text blocks by font size (e.g., Headlines  32pt).
* **Batch Updates:** Processes all slides in a single execution to maintain a clean **Undo** stack.
* **Fluent UI:** Built with Office UI Fabric for a native Look & Feel.

## ⚠️ Security Note

Since this is hosted on GitHub Pages, the first time you run it, you may need to open the URL in **Safari** and "Trust" the site if your organization has strict security certificates.
