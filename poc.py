import pandas as pd
import openpyxl 
from selenium import webdriver
from selenium.webdriver.common.by import By
import base64
import sys
import numpy as np
import cv2
import msoffcrypto
import io
from io import BytesIO
from PIL import Image
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os
#import pdfkit
#from weasyprint import HTML, CSS

def main():
    df = import_df()
    title, title_type, gp_or_gn = title_combine()
    df = convert_to_num(df)
    df = convert_isolate_to_str(df, title_type, gp_or_gn)
    df = mask_combined(df, title_type)
    df = add_tag(df)
    df = flag_column(df, title_type, gp_or_gn)
    footer = add_footer(title_type, gp_or_gn)
    html = make_real_html(df, title, footer)
    html = change_tag(html)
    image = export_to_png(html, title)
    image = remove_br(image, title)
    crop_image(image)

    #intended script stops here
    #write_to_html(html)

def remove_br(image, title):
    if "<br>" in title:
        title = title.replace("<br>", "")
        new_image = './' + title + '.png'
        os.rename(image, new_image)
        return new_image
    else:
        pass

def get_file():
    Tk().withdraw()
    filename = askopenfilename(initialdir = "./", title = "Select file",filetypes = [("Excel Files","*.xlsx")])
    return filename

def crop_image(image):
    img = cv2.imread(image) # Read in the image and convert to grayscale
    #img = img[:-20,:-20] # Perform pre-cropping
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = 255*(gray < 128).astype(np.uint8) # To invert the text to white
    gray = cv2.morphologyEx(gray, cv2.MORPH_OPEN, np.ones((2, 2), dtype=np.uint8)) # Perform noise filtering
    coords = cv2.findNonZero(gray) # Find all non-zero points (text)
    x, y, w, h = cv2.boundingRect(coords) # Find minimum spanning bounding box
    x = 0
    y = 0 #sets starting point to be 0,0
    h = h + 50 #extends the height for extra whitespace
    rect = img[y:y+h, x:x+w] # Crop the image - note we do this on the original image
    #cv2.imshow("Cropped", rect) # Show it
    #cv2.waitKey(0)
    #cv2.destroyAllWindows()
    cv2.imwrite(image, rect) # Save the image

def export_to_png(html, title):
    if sys.platform == "win32":
        driver = webdriver.Firefox()
    if sys.platform == "darwin":
        driver = webdriver.Firefox(executable_path = '/usr/local/bin/geckodriver')
    #install geckodriver for mac (copy to /usr/local/bin/) or windows (copy to Python/Scripts)
    html_base64 = base64.b64encode(html.encode('utf-8')).decode()
    #print(html_base64)
    driver.get("data:text/html;base64," + html_base64)
    image = driver.find_element(By.TAG_NAME, 'body').screenshot(title + '.png')
    driver.quit()
    image = './' + title + '.png'
    return image

def get_concat_v_cut(im1, im2, im3):
    dst = Image.new(
        'RGB', (min(im1.width, im2.width, im3.width), im1.height + im2.height + im3.height))
    dst.paste(im1, (0, 0))
    dst.paste(im2, (0, im1.height))
    dst.paste(im3, (0, im2.height))
    return dst

def mask_combined(df, title_type):
    #this function masks the cell based on the row and column labels
    df = mask(df, "Staphylococcus aureus (includes mssa and mrsa)", "Ampicillin")
    df = mask(df, "Staphylococcus aureus (includes mssa and mrsa)", "High-Level gentamicin")
    df = mask(df, "Coagulase negative staphylococcus", "Ampicillin")
    df = mask(df, "Coagulase negative staphylococcus", "High-Level gentamicin")
    df = mask(df, "E. coli", "Ceftazidime")
    df = mask(df, "Klebsiella pneumoniae", "Ceftazidime")
    df = mask(df, "Proteus mirabilis", "Ceftazidime")
    df = mask(df, "Pseudomonas aeruginosa", "Ampicillin")
    df = mask(df, "Pseudomonas aeruginosa", "Ceftriaxone")
    df = mask(df, "Pseudomonas aeruginosa", "Ertapenem")
    df = mask(df, "Pseudomonas aeruginosa", "Cefazolin")
    df = mask(df, "Pseudomonas aeruginosa", "Cefazolin (Urinary)")
    df = mask(df, "Pseudomonas aeruginosa", "TMP/SMX")
    if title_type in [' All Specimen Types Excluding Surveillance ', " Lower Respiratory "]:
        df = mask_dot(df, "Staphylococcus aureus (includes mssa and mrsa)", "Clindamycin")
        df = mask_dot(df, "Staphylococcus aureus (includes mssa and mrsa)", "Erythromycin (predicts azithromycin)")
        df = mask_dot(df, "Staphylococcus aureus (includes mssa and mrsa)", "TMP/SMX")
        df = mask_dot(df, "Staphylococcus aureus (includes mssa and mrsa)", "Tetracycline")
        df = mask_dot(df, "Staphylococcus aureus (includes mssa and mrsa)", "Rifampin (not to be used as montherapy)")
        df = mask_dot(df, "Staphylococcus aureus (includes mssa and mrsa)", "Vancomycin")
        df = mask_dot(df, "Staphylococcus aureus (includes mssa and mrsa)", "Ciprofloxacin")
    df = mask(df, "Pseudomonas aeruginosa", "Nitrofurantoin (Urine)")
    df = mask(df, "Staphylococcus aureus MRSA", "Ampicillin")
    df = mask(df, "Staphylococcus aureus MSSA", "Ampicillin")
    df = mask(df, "Staphylococcus aureus MRSA", "Erythromycin (predicts azithromycin)")
    df = mask(df, "Staphylococcus aureus MSSA", "Erythromycin (predicts azithromycin)")
    df = mask(df, "Enterococcus spp", "Cefazolin")
    df = mask(df, "Enterococcus faecalis", "Cefazolin")
    df = mask(df, "Enterococcus faecium", "Cefazolin")
    df = mask(df, "Enterococcus spp", "Cloxacillin")
    df = mask(df, "Enterococcus faecalis", "Cloxacillin")
    df = mask(df, "Enterococcus faecium", "Cloxacillin")
    df = mask(df, "Enterococcus spp", "Clindamycin")
    df = mask(df, "Enterococcus spp", "TMP/SMX")
    if title_type != " Urine ":
        df = mask(df, "Enterococcus spp", "Ciprofloxacin")
        df = mask(df, "Enterococcus spp", "Tetracycline")
    else:
        df = mask(df, "Enterobacter spp.", "Fosfomycin (Oral)")
        df = mask(df, "Klebsiella pneumoniae", "Fosfomycin (Oral)")
        df = mask(df, "Proteus mirabilis", "Fosfomycin (Oral)")
        df = mask(df, "Pseudomonas aeruginosa", "Fosfomycin (Oral)")
    df = mask(df, "Enterococcus spp", "Erythromycin (predicts azithromycin)")
    #check this too
    df = mask(df, "Enterococcus spp", "Rifampin (not to be used as montherapy)")
    df = mask(df, "Enterobacter spp.", "Ampicillin")
    df = mask(df, "Enterobacter spp.", "Cefazolin")
    df = mask(df, "Enterobacter spp.", "Ceftriaxone")
    df = mask(df, "Enterobacter spp.", "Ceftazidime")
    df = mask(df, "Enterobacter spp.", "Piperacillin-Tazobactam")
    df = mask(df, "Klebsiella pneumoniae", "Ampicillin")
    return df

def mask_dot(data, row_label, column_label):
    data = data.reset_index()
    data = data.set_index("Name")
    if row_label and column_label in data:
        if row_label in data.index:
            data.loc[row_label, column_label] = "*"
            return data
        else:
            return data
    else:
        return data

def mask(data, row_label, column_label):
    data = data.reset_index()
    data = data.set_index("Name")
    if row_label and column_label in data:
        if row_label in data.index:
            data.loc[row_label, column_label] = "N/R"
            return data
        else:
            return data
    else:
        return data

def flag_column(df, title_type, gp_or_gn):
    df = df.rename(columns={"Isolates": "Number of Isolates"})
    if title_type == ' All Specimen Types Excluding Surveillance ' or title_type == ' Lower Respiratory ':
        if "Gram Positive" in gp_or_gn:
            df = df.rename(columns={"Ciprofloxacin": "Ciprofloxacin**"})
    if title_type == " Blood Cultures ":
        if "Gram Positive" in gp_or_gn:
            df = df.rename(columns={"High-Level gentamicin": "High-Level gentamicin**"})
    if " Urine " in title_type:
        if "Gram Negative" in gp_or_gn:
            df = df.rename(columns={"Cefazolin": "Cefazolin (Urinary)**"})
        if "Gram Positive" in gp_or_gn:
            df = df.rename(columns={"Fosfomycin (Oral)": "Fosfomycin (Oral)**"})
            df = df.rename(columns={"Ciprofloxacin": "Ciprofloxacin***"})
    df = df.reset_index()
    df = df.set_index("dummy")
    return df

def change_tag(html):
    #print(type(html))
    html = html.replace("<td>red", "<td class = 'color_red'>")
    html = html.replace("<td>yellow", "<td class = 'color_yellow'>")
    html = html.replace("<td>green", "<td class = 'color_green'>")
    html = html.replace("<td>grey", "<td class = 'color_grey'>")
    html = html.replace("<th>Name", "<th class = 'hide'>Name")
    html = html.replace("N/R", "")
    return html

def add_tag(df):
    df = df.applymap(apply_color)
    #print(df)
    return df

def convert_isolate_to_str(df, title_type, gp_or_gn):
    df['Isolates'] = pd.to_numeric(df['Isolates'], errors="coerce", downcast="integer")
    df['Isolates'] = df['Isolates'].fillna(0).astype(int)
    df["Isolates"] = df["Isolates"].apply(apply_less_than_30)
    df['Isolates'] = df['Isolates'].replace(to_replace=0, value=' ')
    df['Isolates'] = df['Isolates'].astype(str)
    # print(df['Isolates'].head())
    return df

def convert_to_num(df):
    df = df.set_index("Name")
    df2 = df.apply(lambda x: pd.to_numeric(x.astype(str).str.replace(',', ''), errors='coerce', downcast="integer"))
    df3 = df2.fillna(999)
    df4 = df3.astype(int)
    df5 = df4.replace(to_replace=999, value=' ')
    df5 = df5.reset_index()
    df5["dummy"] = df5["Name"] != ""
    df5 = df5.set_index("dummy")
    return df5

def apply_less_than_30(val):
    hashed = "#" + str(val)
    if type(val) in [float, int]:
        if val >= 30:
            return val
        elif val == 0:
            return val
        else:
            return hashed

def apply_style(df):
    df_color = df.style.applymap(apply_color)
    return df_color

def apply_color(val):
    red = 'red' + str(val)
    green = 'green' + str(val)
    yellow = 'yellow' + str(val)
    grey = "grey" + str(val)
    if type(val) in [float, int]:
        if val >= 90:
            return green
        elif val >= 60:
            return yellow
        return red
    if val == "N/R":
        return grey
    return val

def add_footer(title_type, gp_or_gn):
    footer_first = '<p>* A shaded box indicates that the particular antibiotic/microorganism combinations are not recommended. </p>'
    footer_last = "<p># Fewer than 30 isolates may not be reliable for guiding empiric treatment decisions and cannot be used to statistically compare results to another year. </p>"
    if title_type == ' All Specimen Types Excluding Surveillance ' or title_type == ' Lower Respiratory ':
        if "Gram Negative" in gp_or_gn:
            footer = "<p> ** Cefazolin is not included in this table as automated susceptibility results are not reliable. <br>Refer to the table on blood cultures where alternative methods (Kirby-Bauer) are used for testing. <p>"
            footer = footer_first + footer + footer_last
            return footer
        else:
            pass
        if "Gram Positive" in gp_or_gn:
            footer_1 = "<p> * See MSSA and MRSA <p>"
            footer_2 = "<p> ** Ciprofloxacin monotherapy is NOT recommended for serious infections caused by Staphylococcus spp. <p>"
            footer_first = '<p>*** A shaded box indicates that the particular antibiotic/microorganism combinations are not recommended. </p>'
            footer = footer_1 + footer_2 + footer_first + footer_last
            return footer
        else:
            pass
    else:
        pass
    if title_type == " Blood Cultures ":
        if "Gram Positive" in gp_or_gn:
            footer = "<p> ** For use in combination with Ampicillin or Vancomycin for synergy (for enterococcus species only). </p>"
            footer = footer_first + footer + footer_last
            return footer
        else:
            pass
    else:
        pass
    if title_type == " Urine ":
        if "Gram Negative" in gp_or_gn:
            footer = "<p> ** Cefazolin (urinary) predicts for cephalexin and cefprozil when used for treatment of uncomplicated UTIs due to E. coli, K. pneumoniae, and P. mirabilis but not for  therapy of infections other than uncomplicated UTIs. <p>"
            footer = footer_last + footer_first + footer
            return footer
        elif "Gram Positive" in gp_or_gn:
            footer_1 = "<p> ** For Enterococcus faecalis only. </p>"
            footer_2 = "<p> *** For urine only </p>"
            footer = footer_first + footer_1 + footer_2 + footer_last
            return footer
        else:
            pass
    else:
        pass
    footer = footer_first + footer_last
    return footer

def write_to_html(html):
    write_path = input("Enter location and name for output html file: ")
    write_path = write_path + '.html'
    text_file = open(write_path, "w")
    text_file.write(html)
    text_file.close()

def title_combine():
    title_year = ask_for_year()
    title_area = ask_for_area()
    if title_area == "Hamilton Health Sciences":
        title_location = ask_for_location()
        title_type = ask_for_type()
        title_gp_or_gn = skip_if_strep(title_type)
        title_site = ask_for_site()
        title = title_year + " " + title_site + title_location + title_type + title_gp_or_gn + ' Antibiogram'
    else:
        title_type = JBH_ask_for_type()
        title_gp_or_gn = skip_if_strep(title_type)
        title = title_year + title_area + "<br>" + title_type + title_gp_or_gn + " Antibigram"
    return title, title_type, title_gp_or_gn

def skip_if_strep(title_type):
    if "Streptococcus Pneumoniae" in title_type:
        title_gp_or_gn = ""
    else:
        title_gp_or_gn = gp_or_gn()
    return title_gp_or_gn

def ask_for_area():
    area = input("\n1. Hamilton Health Sciences"
                 "\n2. Joseph Brant Hospital"
                 "\nSelect the service area (Please enter a number): ")
    if area == "1":
        area = "Hamilton Health Sciences"
    if area == "2":
        area = " Joseph Brant Hospital"
    return area

def ask_for_location():
    location = input('\n1. Hospital-Wide'
                     '\n2. ICU'
                     '\n3. CF Clinic'
                     '\nSelect location (Please enter a number): ')
    if location == "1":
        location = ' <br>Hospital-Wide'
    if location == "2":
        location = ' <br>ICU'
    if location == "3":
        location == ' <br>CF Clinic' 
    return location


def gp_or_gn():
    gp_or_gn = input("\n1. Gram Positive\n"
                    "2. Gram Negative\n"
                    "3. Combination\n"
                    "Select type (Please enter a number): ")
    if gp_or_gn == "1":
        gp_or_gn = "Gram Positive"
    if gp_or_gn == "2":
        gp_or_gn = "Gram Negative"
    if gp_or_gn == "3":
        gp_or_gn = "Combination"
    return gp_or_gn  


def ask_for_year():
    year = input('\nPlease enter year for the Antibiogram (e.g. 2020): ')
    return year


def ask_for_type():
    output_type = input('\n1. All specimen types excluding surveillance\n'
                        '2. Urine\n'
                        '3. Blood Cultures\n'
                        '4. Lower Respiratory\n'
                        '5. (S. Pneumoniae) All Specimens Excluding Blood Cultures and Spinal Fluids\n'
                        '6. (S. Pneumoniae) Blood Cultures and Spinal Fluids\n'
                        'Select antibiogram type (Please enter a number): ')
    if output_type == "1":
        output_type = ' All Specimen Types Excluding Surveillance '
    if output_type == "2":
        output_type = ' Urine '
    if output_type == "3":
        output_type = ' Blood Cultures '
    if output_type == "4":
        output_type = ' Lower Respiratory '
    if output_type == "5":
        output_type = ' All Specimens Excluding Blood Cultures and Spinal Fluids Specimemens - Streptococcus Pneumoniae'
    if output_type == "6":
        output_type = ' Blood Cultures and Spinal Fluids Specimens - Streptococcus Pneumoniae'
    return output_type

def JBH_ask_for_type():
    output_type = input('\n1. All specimen types excluding surveillance\n'
                        '2. (S. Pneumoniae) All Specimens Excluding Blood Cultures and Spinal Fluids\n'
                        '3. (S. Pneumoniae) Blood Cultures and Spinal Fluids\n'
                        'Select antibiogram type (Please enter a number): ')
    if output_type == "1":
        output_type = ' All Specimen Types Excluding Surveillance '
    if output_type == "2":
        output_type = ' All Specimens Excluding Blood Cultures and Spinal Fluids Specimemens - Streptococcus Pneumoniae'
    if output_type == "3":
        output_type = ' Blood Cultures and Spinal Fluids Specimens - Streptococcus Pneumoniae'
    return output_type


def ask_for_site():
    title_selection = input("\n1. Hamilton General Hospital\n"
                            "2. McMaster University Medical Centre\n"
                            "3. Juravinski Hospital\n"
                            "4. St. Peter's hospital\n"
                            "5. West Lincoln Memorial Hospital\n"
                            "6. St. Joseph's Healthcare Hamilton\n"
                            "Select Facility (Please enter a number): ")
    if title_selection == "1":
        title_selection = 'Hamilton General Hospital'
    if title_selection == "2":
        title_selection = "McMaster University Medical Centre"
    if title_selection == "3":
        title_selection = "Juravinski Hospital"
    if title_selection == "4":
        title_selection = "St. Peter's hospital"
    if title_selection == "5":
        title_selection = "West Lincoln Memorial Hospital"
    if title_selection == "6":
        title_selection = ""
    return title_selection


def make_real_html(df, title,footer):
    message_start = f"""
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>{ title }</title>"""

    message_style = """
<style type="text/css" media="screen">
    table {
        font-family: 'Times New Roman', Times, serif;
        border-collapse: collapse;
        table-layout: fixed;
        word-wrap: break-word;
    }

    tr {
        border: 1px solid;
        font-size: 8pt;
        text-align: center;
    }

    td {
        border: 1px solid;
        padding: 0px;
    }

    td:nth-child(1) {
        font-style: italic;
        text-align: left;
        padding: 3px;
    }

    .color_grey {
        background-color: #d3d3d3;
        text-align: center;
    }

    .color_red {
        background-color: #fe0000;
        text-align: center;
    }

    .color_yellow {
        background-color: #ffff00;
        text-align: center;
    }

    .color_green {
        background-color: #9acc00;
        text-align: center;
    }

    th {
        font-weight: normal;
        writing-mode: sideways-lr;
        min-width: 20px;
        font-size: 8pt;
        padding: 5px 5px 5px 5px;
        border: 1px solid;
    }

    td:hover {
        font-weight: bold;
        border-style: solid;
        border-width: 2px;
        border-color: black;
        font-size: 7.8pt;
    }

    h2 {
        font-family: 'Times New Roman', Times, serif;
        font-size: 10pt;
    }

    .hide {
        color: white;
        border: 1px solid;
        border-right-color: black;
        border-bottom-color: black;
        text-align: center;
    }

    p {
        font-family: 'Times New Roman', Times, serif;
        font-size: 8pt;
    }
  </style>
</head>
<body>
"""
    message_chart_title = f"""
            <h2>{ title }</h2>
    """
    message_body = df.to_html(index=False)
    message_end = f"""
    { footer }
    </body>"""
    messages = (message_start + message_style + message_chart_title + message_body + message_end)
    return messages

def import_df():
    decrypted = io.BytesIO()
    read_path = get_file()
    key = input("\nEnter File Password: ")
    with open(read_path, "rb") as f:
        file = msoffcrypto.OfficeFile(f)
        file.load_key(password=key)
        file.decrypt(decrypted)
    df = pd.read_excel(decrypted)
    #print(df)
    return df

if __name__ == '__main__':
    main()