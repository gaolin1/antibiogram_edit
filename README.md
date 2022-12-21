### *Requirements*
* Python 3 (tested on versions 3.9 and 3.10 on mac)
* Required packages (to manually install packages, follow the install command specified on package websites, otherwise recommend using the install.py script below):
1. [pandas](https://pandas.pydata.org/docs/getting_started/install.html)
2. [selenium](https://selenium-python.readthedocs.io/installation.html)
3. [pillow](https://pillow.readthedocs.io/en/stable/installation.html) (install command: please refer to the library website for instruction per os)
4. [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
5. [msoffcypto](https://github.com/nolze/msoffcrypto-tool)
6. [numpy](https://numpy.org/install/)
7. [matplotlib](https://matplotlib.org/stable/users/installing/index.html)
8. [PySimpleGUI](https://www.pysimplegui.org/en/latest/)
9. [opencv_python](https://pypi.org/project/opencv-python/)
10. [pdfkit](https://pypi.org/project/pdfkit/)
11. [webdriver_manager](https://pypi.org/project/webdriver-manager/)
---
### *Installation Steps*
1. install the latest version of Python (https://www.python.org/downloads/)
2. Download the latest release from this repository
* <img width="1248" alt="image" src="https://user-images.githubusercontent.com/28236780/208823973-ef87470e-5993-4574-8db2-4a929ee17d9f.png">
3. Click on (source code (zip)) to download the folder
* <img width="631" alt="image" src="https://user-images.githubusercontent.com/28236780/208824341-aed1ca38-aad5-4808-b8b1-ba7a172b381f.png">
4. unzip the files, then right click on install.py and select launch with python
5. after the installation is completed, right click on antibiogram generator.py and select launch with python
- Additional required program: 
  - browser for rendering: [Firefox](https://www.mozilla.org/en-CA/firefox/products/) (only browser to support sideway rotation)
    - recommend rendering running script on mac if available to match rendering on the sample output file 
  - webdriver: install [geckodriver](https://github.com/mozilla/geckodriver/releases) for mac (copy to /usr/local/bin/) or windows (copy to Python/Scripts)
---
### *Intended Use*
This script'smain function: take the Epic exported antibiogram file and performs data transformation (e.g. puts data in table form, add footer and header etc.), finally ouputs with firefox engine for a png file within the same directory.

---
### *Running the script*

#### Within Epic:
> First, extract from Epic: run antibiogram report for desired location (and/or service area for Joseph Brant Hospital) and then export excel file from Epic<br>
#### Then, On local computer:
> To start the script, either 
* Right click on `antibiogram generator.py` and launch with **python launcher**
* Choose JBH button if exporting for JBH, otherwise continue to enter the required fields:
> <img width="830" alt="image" src="https://user-images.githubusercontent.com/28236780/189787187-6dadee46-dc10-4e5e-b723-7ab0d782c8f2.png">
  * Follow the displayed prompts to answer specifics of this antibiogram: 
    * Enter `year` of the antibiogram (e.g. 2021) 
    * Select `facility` of the antibiorgam (e.g. Hamilton General Hospital)
    * Select `location` of the antibiorgam (e.g. ICU)
    * Select `type` of the antibiogram (e.g. Blood)
    * Select either `gram-positive` or `gram negative` or `combination`
    * *then* click on brosw to select the desired excel output from Epic
    * enter the password below
    * *click on "Generate HHS Antibiogram"*
      * *sample output file for testing: Test_Antibiogram_2.xlsx*
    * Then the script will then proceed to perform masking of the antibiotic and organism combination and then renders the file in html format. After rendering, the system will automatically take an screenshot of the window and crop the excess white spaces.
      * note: it's helpful to keep the mouse towards to bottom right portion of the screen to avoid the risk of it been taken part of the screenshot as the script will render in Firefox and take screenshot and crop image for excess white space. 
     * the outputed png will be automatically saved on the same directory 
       * *sample output png: 2001 ICU Gram Positive Urine Antibiogram - MUMC*
