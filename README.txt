Create a virtual environment: 

1) CD into the directory of your project
2) python3 -m venv [name of environment folder]
3) Select the python interpreter in VS Code: 
--View -> Command Palette -> Python:Select Interpreter -> [name of virtual environment folder]\Scripts\python.exe 


Virtual Environment Installations: 
pip install --upgrade google-api-python-client
pip install python-dotenv
pip install openpyxl


Create a .env file with the necessary Credentials (line 7 of the program)


Troubleshooting: 
-make sure .xlsx file is closed before running the program