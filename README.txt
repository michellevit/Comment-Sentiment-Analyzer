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

This Python script intakes a Lex Fridman podcast episode, and outputs all of the episode's YouTube and Reddit comments in an excel sheet, sorted by sentiment (i.e. Positive, Neutral, or Negative). It uses the 'Hugging Face' cardiffnlp roBERTA-base language model, which was trained on ~58 million tweets.


----------
2. Requirements

Libraries: 
- Python 3.9 or higher
- googleapiclient
- praw
- os
- dotenv
- transformers-cli
- huggingface-cli
Other:
- Google account
- YouTube API key
- Reddit account
- Reddit API key


----------
3. Installation:

- Download the repository from GitHub
  - Click the green 'Code' button
  - Click 'Download ZIP'
  - Unzip the file (find the folder in your computer and right-click 'Extract all')
-Install Python 3.9.13 to your computer
  - If it does not work, make sure it is added to your system's PATH
-Create a Virtual environment
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
    pip install praw
    pip install python-dotenv
    pip install transformers
    pip install huggingface-hub
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
 - Enter the desired podcast episode number in the cell B2
- Save and close the file
- Double click the 'run_script_2.bat' file
 - Open the newly created file 'Comment Analyzer - Complete.xlsx' (should appear in the same directory)



----------
5. Troubleshooting: 
* The excel files cannot be open while running the script


----------
6. FAQ

-This script will only provide comments for podcast episodes on the Lex Fridman channel, and comments from the Lex Fridman subreddit
-If no podcast episode is entered into original 'Comment Analyzer.xlsx' file (created upon running 'run_script_1.bat'), then this script will automatically get the comments from the latest podcast episode added to the channel


----------
7. Credits

Michelle Flandin

