# ma3seq2xlsx
Convert exported ma3 sequences to excel format  
Created in **python3**  
It works with grandMA3 software version **2.0.2.0**
# Installation
on MacOs/linux
- Install latest python from: https://www.python.org
- Or if you have brew:
  - `brew install python@3.12`
- run in a terminal with:
  - `source .venv/bin/activate`
- then 
  - `python main.py`
 
If you not have _venv_  or can't use it,   
then need to install openpyxl  
system wide, with:

  - `$ pip install openpyxl`
- run with:
  - `python main.py`

# Usage
The `main.py` shows a menu of your exported  
sequences, from 1-to exported sequence files.  
>------------------------------  
>MENU  
>------------------------------  
>1 -- FirstExportedSequence.xml  
>2 -- SecondExportedSequence.xml  
>3 -- ThirdExportedSequence.xml  
>4 -- Exit  
>Choose xml file to convert to xlsxs.  

Last menu number is always exit the program.  
Choose the file you want to convert to xlsx.  
Then program will execute, and writes the  
choosen xml file to xlsx/[samename].xlsx file  
:
>Chosen XML file:  ..FirstExportedSequence.xml  
>xls directory not found, try to create.  
>Writing file done.  
> Restarting  

The finished files are in `./xlsx` dir under main directory.  
Have fun! :-)

