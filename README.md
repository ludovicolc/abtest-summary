# abtest-summary

A Python utility for automating the creation and formatting of A/B test summaries in Google Sheets.

## Installation

You can install the library directly from the GitHub repository using `pip`:

pip install git+[https://github.com/ludovicolc/abtest-summary.git](https://github.com/ludovicolc/abtest-summary.git)


## Setup Instructions

Follow the steps below to set up the service account and enable the Google Sheets API:

1. **Create a Service Account on Google Cloud Console**  
   Go to [Google Cloud Console](https://console.cloud.google.com) and create a service account for your project.

2. **Enable the Google Sheets API**  
   In the **API & Services** section, enable the **Google Sheets API** for your project.

3. **Download the Credential JSON File**  
   Once the service account is created, download the `.json` file containing the credentials.

4. **Add the JSON File to Your Project Folder**  
   Place the downloaded `.json` file in the root directory of your project.

5. **Specify the Path in Your Class**  
   In your code, set the path to the `.json` file inside your class to authenticate with the Google Sheets API.
