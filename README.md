### **Requirements**
- Python 3 (tested on versions 3.9 and 3.10 on mac)
- Required packages:
1. [pandas](https://pandas.pydata.org/docs/getting_started/install.html) (command: `pip install pandas`)
2. [selenium](https://selenium-python.readthedocs.io/installation.html) (command: `pip install selenium`)
3. [pillow](https://pillow.readthedocs.io/en/stable/installation.html) (command: please refer to the library website for instruction per os)
4. [openpyxl](https://openpyxl.readthedocs.io/en/stable/) (command: `pip install openpyxl`)
5. [msoffcypto](https://github.com/nolze/msoffcrypto-tool) (command: `pip install msoffcrypto-tool`)
6. [numpy](https://numpy.org/install/) (command: `pip install numpy`)
---
This script has one main function: take the Epic exported antibiogram file and performs data transformation and ouputs with firefox engine for a finalized png file.

*Steps*:
> export excel file from Epic
1. Enter exported file path from Epic <br>
    example: /Users/'User Name'/Downloads/'file name'.xlsx *(Mac)* <br>
2. Enter *Year* (e.g. 2021) <br>
3. Select *Type* (e.g. Blood) <br>
4. Select *Facility* (e.g. Hamilton General Hospital) <br>
5. Enter export file path or file name directly to export on the same folder: <br>
    NOTE: html extension will be automatically added <br>
    example: 'output file name' *(to export on the same folder of the script)* <br>
    example: /Users/'User Name'/Downloads/'output file name' *(to export to an exact folder)* <br>
