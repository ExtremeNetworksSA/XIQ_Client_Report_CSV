# HistoricalClient_csvtoexcel.py


## Requirements

This script was written to use python version 2.7 that comes installed default with mac os. To preform the needed tasks to read the csv and and export to excel a couple modules will need to be installed.  
These modules are needed in order to run this script. Please see the Setup Section at the bottom for instructions on installing the needed module. 

When running the script, two arguments need to be give

1. The name of the site the report will be generated for. The name of the site should be added within quotes. ```"Retail Store 1234"```
2. The name of the csv file that is located in the same directory as the HistoricalClient_csvtoexcel.py script.


## How to run the script

To run the script use terminal and go to the directory the script is in. Then run the script with the needed arguements.
```
python HistoricalClient_csvtoexcel.py "<site name>" <csv file>
```

## Setup

pip can be used to install modules but that will need to be installed on the mac if it currently is not. 

To check if pip is currently installed, in Terminal run ```pip --version```. If pip is installed a version number will be in the response. If pip is not installed it can be installed by running ```curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py``` followed by ```sudo python get-pip.py```

> credit [blog](https://ahmadawais.com/install-pip-macos-os-x-python/) - can view for more detailed instructions

Once pip is installed, you can install the needed modules located in the requirements.txt file. To install these modules from the requirements.txt file enter:
```
pip install -r requirements.txt
``` 