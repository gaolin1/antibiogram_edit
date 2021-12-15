from selenium import webdriver
from webdriver_manager.firefox import GeckoDriverManager

def main():
    driver = webdriver.Firefox(executable_path = '/usr/local/bin/geckodriver')
    #put in instruction to install geckodriver in install path
    driver.get("file:///Users/LinBin/Documents/repos/antibiogram_edit/output_final.html")
    driver.save_screenshot("test.png")
    driver.quit()

if __name__ == '__main__':
    main()