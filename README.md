# AI Cat Assistant: Contract PDF → Excel

This is a web application that extracts rate information, blackout dates, and minimum night stay requirements from hotel/cruise contract PDFs and exports them directly into an Excel (`.xlsx`) format.

## Features
- **Modern UI:** Clean, glassmorphism-based user interface.
- **Auto-Save:** Form inputs automatically save to your browser's local storage.
- **AI Extraction:** Uses the powerful Google Gemini 2.0 API to intelligently parse complex documents.
- **Excel Export:** Instant client-side Excel generation using SheetJS.

## Setup & Deployment (GitHub Pages)
This app is entirely front-end based (Vanilla HTML, CSS, JavaScript). You can host it for free on GitHub Pages!

1. Upload all files (including `index.html`) to a new GitHub repository.
2. Go to your repository's **Settings**.
3. Select **Pages** on the left sidebar.
4. Under **Build and deployment**, select **Deploy from a branch**.
5. Choose the `main` (or `master`) branch and click **Save**.
6. Wait a few minutes, and your site will be live!

## Usage
1. Open the hosted website.
2. Enter your Gemini API Key in the top input box (it will be saved securely in your browser).
3. Upload a Contract PDF.
4. Fill out the property type, Hotel ID, Supplier, and Room details.
5. Click **Let the AI Cat Extract Data!** and wait for the Excel file to generate.
