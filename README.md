---

# Lokalise XLSX Converter

A lightweight, multi-threaded desktop application built with Python and PyQt6. This tool batch-converts translation data exported from Lokalise (or similar localization platforms) into a structured target XLSX format. 

## ✨ Features

* **User-Friendly GUI:** Built with PyQt6, featuring safe folder-selection dialogs and a real-time progress bar.
* **Batch Processing:** Recursively scans an input directory for `.xlsx` files and processes them all at once.
* **Folder Mirroring:** Automatically recreates the exact subfolder structure of your input directory in your output directory.
* **Smart Formatting:** Auto-fits Excel column widths for immediate readability using `openpyxl`.
* **Non-Blocking UI:** Uses `QThread` to handle heavy Pandas processing in the background, keeping the application responsive.

## 🛠️ Prerequisites

To run the source code, you will need Python 3 installed on your machine along with the following libraries:

```bash
pip install pandas openpyxl PyQt6 pyinstaller
```

## 🚀 Running from Source

1. Clone this repository to your local machine.
2. Ensure your required libraries are installed.
3. Run the main application script:
   ```bash
   python main.py
   ```

## 📦 Building the Executable (.exe)

You can package this tool into a standalone executable so it can be shared with users who do not have Python installed. The repository includes a `.spec` file pre-configured to bundle the application, embed the custom icon, and minimize file size by excluding unused Qt modules.

1. Open your terminal in the project root directory.
2. Run the PyInstaller build command:
   ```bash
   pyinstaller LokaliseConverter.spec
   ```
3. Once the build completes, your standalone executable will be located in the newly created `dist/` folder.

## 📖 How to Use

1. Launch the application.
2. Click **Browse...** to select your **Input Folder** (containing your source `.xlsx` files).
3. Click **Browse...** to select your **Output Folder** (where the converted files will be saved).
4. Click **Run**.
5. Wait for the progress bar to complete. A popup will notify you when all files are successfully converted!

---

### 📝 License

This project is licensed under the Apache 2.0 License. © 2026 Vistatec, Ltd.

---