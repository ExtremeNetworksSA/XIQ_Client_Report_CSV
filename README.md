# HistoricalClient_csvtoexcel.py


## Requirements

This script was written to work with python version 2.7 that comes installed default with mac os as well as the new version of python 3.9 if installed manually. To preform the needed tasks to read the csv and and export to excel a couple modules will need to be installed.  
These modules are needed in order to run this script. Please see the Setup Section at the bottom for instructions on installing the needed module. 

When running the script, two arguments need to be give

1. The name of the site the report will be generated for. The name of the site should be added within quotes. ```"Retail Store 1234"```
2. The name of the csv file that is located in the same directory as the HistoricalClient_csvtoexcel.py script.


## How to run the script

### MAC OS
To run the script use terminal and go to the directory the script is in. Make sure the downloaded csv file is in the same directory, then run the script with the needed arguements.
```
python HistoricalClient_csvtoexcel.py "<site name>" <csv file>
```
### Windows 10
To run the script use PowerShell and go to the directory the script is in. Make sure the downloaded csv file is in the same directory, then run the script with the needed arguements.
```
python.exe .\HistoricalClient_csvtoexcel.py "<site name>" .\<csv file>
```
## Setup

### MAC OS
pip can be used to install modules but that will need to be installed on the mac if it currently is not. 

To check if pip is currently installed, in Terminal run ```pip --version```. If pip is installed a version number will be in the response. If pip is not installed it can be installed by running ```curl https://bootstrap.pypa.io/pip/2.7/get-pip.py -o get-pip.py``` followed by ```sudo python get-pip.py```

> credit [blog](https://ahmadawais.com/install-pip-macos-os-x-python/) - can view for more detailed instructions

Once pip is installed, you can install the needed modules located in the requirements.txt file. Navigate in the termial to the folder containing the script and requirements.txt file. To install these modules from the requirements.txt file enter:
```
pip install -r requirements.txt
``` 

### Windows 10
Install python from the Microsoft Store (tested with python 3.9). This will also install pip
Open Windows PowerShell and navigate to the to the folder containing the script and requirements.txt file. To install these modules from the requirements.txt file enter:
```
pip install -r .\requirements.txt
``` 