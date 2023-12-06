Project Title: Comment Sentiment Analyzer


----------
Table of Contents: 
1. Project Description
2. Requirements
3. Installation
4. How to Run the Program
5. Troubleshooting
6. FAQ
7. Credits


----------
1. Project Description

This Python script intakes a YouTube video link, and outputs all of the video's YouTube comments in an excel sheet, sorted by sentiment (i.e. Positive, Neutral, or Negative). The sentiment analysis done by the 'Hugging Face' cardiffnlp roBERTA-base language model, which was trained on ~58 million tweets.

----------
2. Requirements

Libraries: 
- Python 3.9
- googleapiclient
- openpyxl
- os
- dotenv
- transformers
- tensorflow
- huggingface-hub
- langdetect
Other:
- Google account
- YouTube API key


----------
3. Installation:

- Download the repository from GitHub
  - Click the green 'Code' button
  - Click 'Download ZIP'
  - Unzip the file (find the folder in your computer and right-click 'Extract all')
- Install Python 3.9 to your computer
  - If it does not work, make sure it is added to your system's PATH
  - Note: It must be version 3.9 for Tensorflow to work
- Install Rust
  - Download Visual Studio C++ Build Tools (pre-req for Rust)
    - https://visualstudio.microsoft.com/downloads/
    - Under "All downloads" > "Tools for Visual Studio 2019", find "Build Tools for Visual Studio 2019" and download it. (You can also choose a newer version if available.)
    - Run the installer
      - In the installer, select the "C++ build tools" workload
      - Make sure to include the latest versions of MSVCv142 - VS 2019 C++ x64/x86 build tools and Windows 10 SDK
      - Restart the computer
In the installer, select the "C++ build tools" workload.
    - https://visualstudio.microsoft.com/visual-cpp-build-tools/
  - https://www.rust-lang.org/learn/get-started
  - If it does not work, make sure it is added to your system's PATH
- Create a Virtual environment
  - https://www.youtube.com/watch?v=APOPm01BVrk&ab_channel=CoreySchafer
  - Open a new terminal and navigate to folder you want to save your virtual env to
  - Make sure pip is installed (to check - run: python -m pip --version)
    - If pip is missing, download pip from the official source: https://bootstrap.pypa.io/get-pip.py 
  - Run: python -m pip install virtualenv
  - Run: python -m virtualenv Comment-Sentiment-Analyzer-Virtual-Env
- Install required libraries in virtual env
  - Open a terminal
  - Activate the virtual environment - run: C:\Users\Michelle\.virtualenvs\Comment-Sentiment-Analyzer-Virtual-Env\Scripts\Activate.ps1
  - Run: 
    pip install --upgrade google-api-python-client
    pip install openpyxl
    pip install python-dotenv
    pip install transformers
    pip install tensorflow
    pip install huggingface-hub
    pip install langdetect
- Open the 'env.txt' file in VSCode (or a text/code editor)
  - Fill in all the 'youtube_api_key' line
    - How to Get a YouTube API Key:
      - Log in to Google Developers Console
      - Create a new project
      - On the new project dashboard, click Explore & Enable APIs
      - In the library, navigate to YouTube Data API v3 under YouTube APIs
      - Enable the API
      - Create a credential
      - A screen will appear with the API key
  - Rename the 'env.txt' file to '.env' and save it
- Update the file paths in 'run_script_1.bat' and 'run_script_2.bat' 
  - Note: there are 2 paths per file which need to be updated


----------
4. How to Run the Program: 

- Double click the 'run_script_1.bat' file
  - Open the newly created 'Comment-Analyzer.xlsx' file 
  - Enter the YouTube video link in the cell B2
- Save and close the file
- Double click the 'run_script_2.bat' file
  - Once it is ready, it will be renamed: "YouTube-Comment-Analyzer-Complete.xlsx"
  * Note: on average this script takes 3 seconds per comment, so you may need to wait
  * Note: replies to comments will not be recorded
  * Note: only english comments will be recorded


----------
5. Troubleshooting: 
- The excel files cannot be open while running the script


----------
6. FAQ
  - On average this script takes 3 seconds per comment, so you may need to wait
  - Replies to comments will not be recorded
  - Only english comments will be recorded


----------
7. Credits

Michelle Flandin

