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

This script has one main function: take the Epic exported antibiogram file and performs data transformation, finally ouputs with firefox engine for a png file within the same directory.

---

*Steps*:
> to extract from Epic: run antibiogram report for desired location and then export excel file from Epic<br>
> to run the script, either 
> 1. enter `python3 antibiogram_edit.py` in terminal (mac) or cmd (windows)
> 2. use right click on the antibiogram_edit.py file and launch with **python launcher** 
---
> <img width="711" alt="image" src="https://user-images.githubusercontent.com/28236780/152648615-d4b03d32-a5f9-4a80-b003-5672771eefa8.png">
---
* Enter exported file path from Epic, example: /Users/'User Name'/Downloads/'file name'.xlsx *(on Mac)*
* Enter *Year* (e.g. 2021) 
* Select *Type* (e.g. Blood)
* Select *Facility* (e.g. Hamilton General Hospital)
* Enter export file path or file name directly to export on the same folder: 
    * NOTE: html extension will be automatically added <br>
    * example: 'output file name' *(to export on the same folder of the script)*
    * example: /Users/'User Name'/Downloads/'output file name' *(to export to an exact folder)* <br>
