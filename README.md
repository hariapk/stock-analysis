Streamlit Stock Data Processor and Analysis Tool

This application is a Python-based web tool built using Streamlit, Pandas, and Openpyxl. It processes an uploaded Excel file, calculates industry averages, applies custom flagging logic, sorts the data, and formats the output with alternating row shading.

⚙️ Project Structure

The key files for this application are:

File Name

Purpose

stock-analysis.py

Contains the core Python logic and Streamlit UI.

requirements.txt

Lists Python dependencies (streamlit, pandas, openpyxl). MANDATORY for deployment.

README.md

Deployment and setup instructions.

🚀 Deployment with GitHub and Streamlit Community Cloud (Recommended)

The easiest way to deploy a Streamlit app using GitHub is by leveraging the Streamlit Community Cloud platform.

Step 1: Push to GitHub

Repository Setup:

Create a new GitHub repository (e.g., streamlit-stock-processor).

Add all three files (stock-analysis.py, requirements.txt, and README.md) to the root of this repository.

Commit and push the files to your main branch (main or master).

Step 2: Deploy on Streamlit Community Cloud

Sign Up/Log In: Go to the Streamlit Community Cloud website and log in using your GitHub account.

New App: Click the + New app button.

Link Repository:

Select your GitHub repository (streamlit-stock-processor).

Enter the main file path: stock-analysis.py (The name of your main file).

Click Deploy!.

Streamlit will automatically detect the requirements.txt file, install all dependencies, and build your application. Once built, your app will be live and accessible via a public URL.

💻 Local Testing

You can run this application locally on your machine for testing before deployment.

Clone Repository:

git clone <your-repo-url>
cd <your-repo-name>


Install Dependencies:

pip install -r requirements.txt


Run Streamlit:

streamlit run stock-analysis.py


The application will automatically open in your web browser, usually at http://localhost:8501.