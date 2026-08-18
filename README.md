# Raffle Spinner App 🎰

An interactive, web-based lottery and raffle spinner application built with HTML, CSS, JavaScript, and powered by Python data-processing scripts. 

This project was originally developed for a community raffle event and has been fully anonymized and open-sourced to serve as a robust template for any organization looking to host an engaging, live raffle drawing.

## ✨ Features

- **Interactive Spinning Wheel:** A visually engaging roulette-style wheel with animations, sound effects, and confetti for drawing winners.
- **Prize Catalog Integration:** Displays the current prize being drawn along with its image and description.
- **State Persistence (Crash-Proof):** Winners are automatically saved to the browser's `localStorage`. If the browser crashes or is accidentally refreshed, the state is preserved and drawn prizes are not lost.
- **Fair Randomization:** Secure client-side logic to ensure fair picking among purchased tickets, accounting for users who bought multiple tickets for the same prize.
- **Python Data Pipeline:** Includes Python scripts to process donor data from Excel/CSV files, calculate statistics, and build the final JSON database used by the web app.

## 🛠️ Tech Stack

- **Frontend:** HTML5, CSS3, Vanilla JavaScript (No heavy frameworks required)
- **Backend/Data Processing:** Python 3 (Pandas)
- **Fonts:** Google Fonts (Heebo, Rubik)

## 📁 Project Structure

* `raffle_spinner_template.html` - The base HTML template containing the UI and JavaScript logic.
* `build_web_app.py` - Python script that injects `raffle_data.json` into the HTML template to generate the final `raffle_spinner.html`.
* `raffle_data.json` - The main database file (anonymized) containing prizes, tickets, and participant allocation.
* `raffle_spinner.html` - The final generated single-page application (Open this in your browser to run the raffle).
* `analyze.py`, `parse_data.py` - Utility scripts to analyze ticket distributions and generate statistics.
* `catalog_pages/`, `prize_images/` - Directories for storing images of the prizes being raffled.

## 🚀 How to Use

1. **Prepare your data:** Organize your participants and tickets into a structured format (see `raffle_data.json` for the required schema).
2. **Build the App:** 
   Run the following command to inject your data into the HTML:
   ```bash
   python build_web_app.py
   ```
3. **Run the Raffle:** 
   Simply double-click `raffle_spinner.html` to open it in any modern web browser. No local web server is strictly required, though using one (like Live Server) is recommended.
4. **Draw Winners:** Select a prize category, click "Spin", and celebrate!

## 🔒 Privacy & Security

This repository contains **anonymized mock data**. All Personally Identifiable Information (PII) from the original event (such as real names, phone numbers, and donation amounts) has been completely scrubbed and replaced with generic placeholder data to comply with privacy standards.

## 📄 License

This project is open-source and available under the MIT License. Feel free to fork, modify, and use it for your own community events!
