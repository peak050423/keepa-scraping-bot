# Automated Product Tracking with reCAPTCHA Solver

## Description:

This script automates the process of tracking products on Keepa.com. It uses Selenium to interact with the website, handles reCAPTCHA challenges using audio recognition, and tracks products based on the ASINs (Amazon Standard Identification Numbers) provided in an Excel or OpenOffice file.

### Key Features:

- Upload an Excel or OpenOffice file containing ASINs.
- Automatically solves reCAPTCHA using audio recognition.
- Tracks products by specifying price thresholds.
- Supports multiple locales for Amazon product tracking.

---

## Install Dependencies:

Before running the script, you'll need to install the required Python dependencies. Open a terminal or command prompt and run the following command:

```bash
pip install selenium pandas webdriver-manager tk odfpy pydub SpeechRecognition
```

Also, you need to install ffmpeg. You can download it from

For Windows:

- Download the latest version of ffmpeg from [here](https://ffmpeg.org/download.html).
- Extract the zip file and add the bin folder to your system's PATH environment variable.

For macOS:

- Install ffmpeg via Homebrew:

```bash
brew install ffmpeg
```

For Linux:

- Use your package manager to install ffmpeg. For example, on Ubuntu

```bash
sudo apt-get install ffmpeg
```

Usage:

1. **Run the Script**

   - The script uses Selenium to open Chrome and interact with Keepa.com.
   - When prompted, select an Excel or OpenOffice file that contains the ASINs for the products you wish to track.
     Handling reCAPTCHA:

2. **Handling reCAPTCHA**

   - If a reCAPTCHA challenge is detected, the script will attempt to solve it automatically using audio recognition. The script will download the audio challenge, process it, and submit the solution.

3. **Tracking Products**

   - The script will search for each ASIN, attempt to track the product, and set the price thresholds as specified in the Excel file.

4. **Completion**

   - After processing all ASINs, the script will save the updated file with tracking statuses.

5. **Manual Login**

   - You will be prompted to log in manually to Keepa.com. Once logged in, press Enter to allow the script to continue.

## Run the Script

```bash
python newSkriptKeepa.py
```
