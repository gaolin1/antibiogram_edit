### **Requirements**
- Python 3 (tested on versions 3.9 and 3.10 on mac)
- Required packages:
1. [pandas](https://pandas.pydata.org/docs/getting_started/install.html) (command: `pip install pandas`)
2. [selenium](https://selenium-python.readthedocs.io/installation.html) (command: `pip install selenium`)
3. [pillow](https://pillow.readthedocs.io/en/stable/installation.html) (command: please refer to the library website for instruction per os)
4. [openpyxl](https://openpyxl.readthedocs.io/en/stable/) (command: `pip install openpyxl`)
5. [msoffcypto](https://github.com/nolze/msoffcrypto-tool) (command: `pip install msoffcrypto-tool`)
6. [numpy](https://numpy.org/install/) (command: `pip install numpy`)
- Additional required program: 
  - browser for rendering: [Firefox](https://www.mozilla.org/en-CA/firefox/products/)
  - webdriver: install [geckodriver](https://github.com/mozilla/geckodriver/releases) for mac (copy to /usr/local/bin/) or windows (copy to Python/Scripts)
---

This script has one main function: take the Epic exported antibiogram file and performs data transformation, finally ouputs with firefox engine for a png file within the same directory.

---

*Steps*:
> First, extract from Epic: run antibiogram report for desired location and then export excel file from Epic<br>
> Then, to run the script, either 
> 1. enter `python3 antibiogram_edit.py` in terminal (mac) or cmd (windows)
> 2. use right click on the antibiogram_edit.py file and launch with **python launcher** 
---
> <img width="417" alt="image" src="https://user-images.githubusercontent.com/28236780/152649476-023b2235-0a78-42a5-a91e-52b09c0c6b58.png">
---
* Enter exported file path from Epic, example: /Users/'User Name'/Downloads/'file name'.xlsx *(on Mac)*
   * enter password to the Epic extract file. 
* Enter year of the antibiogram (e.g. 2021) 
* Select Type of the antibiogram (e.g. Blood)
* Select Facility of the antibiorgam (e.g. Hamilton General Hospital)
* Then the script will perform masking of the antibiotic and organism combination and proceeds to render the file in html format
   * note: it's helpful to keep the mouse towards to bottom right portion of the screen to avoid the risk of it been taken part of the screenshot as the script will render in Firefox and take screenshot and crop image for excess white space. 
* the outputed png will be automatically saved on the same directory 
