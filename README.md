### *Requirements*
* Python 3 (tested on versions 3.9 and 3.10 on mac)
* Required packages:
1. [pandas](https://pandas.pydata.org/docs/getting_started/install.html) (install command: `pip install pandas`)
2. [selenium](https://selenium-python.readthedocs.io/installation.html) (install command: `pip install selenium`)
3. [pillow](https://pillow.readthedocs.io/en/stable/installation.html) (install command: please refer to the library website for instruction per os)
4. [openpyxl](https://openpyxl.readthedocs.io/en/stable/) (install command: `pip install openpyxl`)
5. [msoffcypto](https://github.com/nolze/msoffcrypto-tool) (install command: `pip install msoffcrypto-tool`)
6. [numpy](https://numpy.org/install/) (install command: `pip install numpy`)
7. [tkinter](https://docs.python.org/3/library/tkinter.html) (install command: `pip install tk`)
- Additional required program: 
  - browser for rendering: [Firefox](https://www.mozilla.org/en-CA/firefox/products/) (only browser to support sideway rotation)
    - recommend rendering running script on mac if available to match rendering on the sample output file 
  - webdriver: install [geckodriver](https://github.com/mozilla/geckodriver/releases) for mac (copy to /usr/local/bin/) or windows (copy to Python/Scripts)
---
### *Intended Use*
This script has one main function: take the Epic exported antibiogram file and performs data transformation (e.g. puts data in table form, add footer and header etc.), finally ouputs with firefox engine for a png file within the same directory.

---
### *Running the script*

#### Within Epic:
> First, extract from Epic: run antibiogram report for desired location (and/or service area for Joseph Brant Hospital) and then export excel file from Epic<br>
#### Then, On local computer:
> To start the script, either 
> 1. enter `python3 antibiogram_edit.py` in terminal (mac) or cmd (windows)
> 2. OR use right click on the antibiogram_edit.py file and launch with **python launcher**
* Once the script is started, select the exported excel file from Epic in the dialog window  
   * enter password to the Epic extract file
      * <img width="334" alt="image" src="https://user-images.githubusercontent.com/28236780/155064252-df332a88-c6ea-4d59-bc93-572712ab787e.png"> 
      * *sample output file for testing: Test_Antibiogram_2.xlsx*
> <img width="600" alt="image" src="https://user-images.githubusercontent.com/28236780/155063556-aa837e87-496b-414a-bd2b-62b5f6cf6581.png"> 
* Follow the displayed prompts to answer specifics of this antibiogram: 
  * Enter `year` of the antibiogram (e.g. 2021) 
  * Select if this is a `HHS` or `JBH` antibiogram
  * Select `location` of the antibiorgam (e.g. ICU)
  * Select either `gram-positive` or `gram negative`
  * Select `type` of the antibiogram (e.g. Blood)
  * Select `facility` of the antibiorgam (e.g. Hamilton General Hospital)
    * Then the script will then proceed to perform masking of the antibiotic and organism combination and then renders the file in html format. After rendering, the system will automatically take an screenshot of the window and crop the excess white spaces.
      * note: it's helpful to keep the mouse towards to bottom right portion of the screen to avoid the risk of it been taken part of the screenshot as the script will render in Firefox and take screenshot and crop image for excess white space. 
     * the outputed png will be automatically saved on the same directory 
       * *sample output png: 2001 ICU Gram Positive Urine Antibiogram - MUMC*
> <img width="699" alt="image" src="https://user-images.githubusercontent.com/28236780/155064983-c01072b9-9e97-42ca-853e-1f717547a10e.png">
