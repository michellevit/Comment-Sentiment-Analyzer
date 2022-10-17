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
-Python 3.9 or higher
-googleapiclient
-praw
-os
-dotenv
-transformers-cli
-huggingface-cli
Other:
-a Google account
-a YouTube API key
-a Reddit account
-a Reddit API key


----------
3. Installation:

-Install Python 3.9.13 or higher
-Create Virtual Environment (https://www.youtube.com/watch?v=APOPm01BVrk&t=0s): 
--open terminal
--cd into folder
--create Python environment: python -m venv project_yt_api
--activate the virtual environment: project_yt_api\Scripts\activate.bat
--install required libraries


----------
4. How to Run the Program: 

Double click the 'run_script_1.bat' file
-enter the desired podcast episode number in the cell B2
-save and close the file
Double click the 'run_script_2.bat' file
-wait for the new file 'Comment Analyzer - Complete.xlsx' to appear in the same directory, and information will appear in the terminal


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

