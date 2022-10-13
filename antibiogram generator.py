from cgitb import text
from this import d
from turtle import color
from xml.etree.ElementTree import TreeBuilder
import PySimpleGUI as sg
import os
import io
import msoffcrypto
import pandas as pd
import cv2
import numpy as np
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
import base64

def main():
    ##general utility functions
    #returns the selected variable per string value
    def get_selected_var(window, list, initial_value, rel_pos):
        var_value = window[initial_value].TKIntVar.get()
        var_selected = list[int(var_value/1000%100)-rel_pos] if var_value else None    
        #return var_value #(for troubleshooting)    
        return var_selected
    
    #subfunction to convert isolate to string
    def apply_less_than_30(val):
        hashed = "#" + str(val)
        if type(val) in [float, int]:
            if val >= 30:
                return val
            elif val == 0:
                return val
            else:
                return hashed
    
    #subfunction to mask specific cell
    def mask(data, row_label, column_label, value):
        data = data.reset_index()
        data = data.set_index("Name")
        if row_label and column_label in data:
            if row_label in data.index:
                data.loc[row_label, column_label] = value
                return data
            else:
                return data
        else:
            return data
    
    #subfunction to apply color to antibiogram
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

    #title section
    #for title clean_up if it's Strep antibiogram
    def skip_if_strep(title_type, input):
        if "Streptococcus Pneumoniae" in title_type:
            title_gp_or_gn = ""
        else:
            title_gp_or_gn = input
        return title_gp_or_gn   

    #main function: takes read path of the file and password and returns the dataframe
    def import_df(read_path, key):
        decrypted = io.BytesIO()
        #read_path = get_file()
        #key = input("\nEnter File Password: ")
        with open(read_path, "rb") as f:
            file = msoffcrypto.OfficeFile(f)
            file.load_key(password=key)
            file.decrypt(decrypted)
        df = pd.read_excel(decrypted)
        #print(df)
        return df
    
    ##main function to generate antibiogram df
    def generate_antibiogram_df(df, title_type, title_gp_or_gn):
        df = convert_to_num(df)
        df = convert_isolate_to_str(df)
        df = swap_mssa(df)
        df = mask_combined(df, title_type)
        df = add_tag(df)
        df = flag_column(df, title_type, title_gp_or_gn)
        return df

    #converts dataframe to numbers for future manupliation
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
    
    #after the numbers are converted to number, to apply less than 30 filtering then returned back to string
    def convert_isolate_to_str(df):
        df['Isolates'] = pd.to_numeric(df['Isolates'], errors="coerce", downcast="integer")
        df['Isolates'] = df['Isolates'].fillna(0).astype(int)
        df["Isolates"] = df["Isolates"].apply(apply_less_than_30)
        df['Isolates'] = df['Isolates'].replace(to_replace=0, value=' ')
        df['Isolates'] = df['Isolates'].astype(str)
        # print(df['Isolates'].head())
        return df

    #replace string within df
    def swap_mssa(df):
        df = df.reset_index()
        df = df.set_index("Name")
        if "Staphylococcus aureus" in df.index:
            df = df.reset_index()
            df["Name"] = df["Name"].replace("Staphylococcus aureus", "Staphylococcus aureus MSSA")
            df = df.set_index("Name")
            mssa_poc, combined_poc = df.index.get_loc("Staphylococcus aureus MSSA"), df.index.get_loc("Staphylococcus aureus (includes mssa and mrsa)")
            df = df.reset_index()
            df_first = df.iloc[:mssa_poc]
            df_last = df.iloc[combined_poc+1:]
            combined = df.iloc[combined_poc]
            mssa = df.iloc[mssa_poc]
            df_first.loc["Staphylococcus aureus (includes mssa and mrsa)"] = combined
            df_first.loc["Staphylococcus aureus MSSA"] = mssa
            df = pd.concat([df_first,df_last],ignore_index=True)
            df = df.set_index("Name")
            return df
        else:
            return df

    #this function masks the cell based on the row and column labels
    def mask_combined(df, title_type):
        df = mask(df, "Staphylococcus aureus MRSA", "Cefazolin", 0)
        df = mask(df, "Staphylococcus aureus MRSA", "Cloxacillin", 0)
        df = mask(df, "Staphylococcus aureus MSSA", "Cefazolin", 100)
        df = mask(df, "Staphylococcus aureus MSSA", "Cloxacillin", 100)
        #mask for N/R
        df = mask(df, "Staphylococcus aureus (includes mssa and mrsa)", "Ampicillin", "N/R")
        df = mask(df, "Staphylococcus aureus (includes mssa and mrsa)", "High-Level gentamicin", "N/R")
        df = mask(df, "Coagulase negative staphylococcus", "Ampicillin", "N/R")
        df = mask(df, "Coagulase negative staphylococcus", "High-Level gentamicin", "N/R")
        df = mask(df, "E. coli", "Ceftazidime", "N/R")
        df = mask(df, "Klebsiella pneumoniae", "Ceftazidime", "N/R")
        df = mask(df, "Proteus mirabilis", "Ceftazidime", "N/R")
        df = mask(df, "Pseudomonas aeruginosa", "Ampicillin", "N/R")
        df = mask(df, "Pseudomonas aeruginosa", "Ceftriaxone", "N/R")
        df = mask(df, "Pseudomonas aeruginosa", "Ertapenem", "N/R")
        df = mask(df, "Pseudomonas aeruginosa", "Cefazolin", "N/R")
        df = mask(df, "Pseudomonas aeruginosa", "Cefazolin (Urinary)", "N/R")
        df = mask(df, "Pseudomonas aeruginosa", "TMP/SMX", "N/R")
        if title_type in [' All Specimen Types Excluding Surveillance ', " Lower Respiratory "]:
            df = mask(df, "Staphylococcus aureus (includes mssa and mrsa)", "Clindamycin", "*")
            df = mask(df, "Staphylococcus aureus (includes mssa and mrsa)", "Erythromycin (predicts azithromycin)", "*")
            df = mask(df, "Staphylococcus aureus (includes mssa and mrsa)", "TMP/SMX", "*")
            df = mask(df, "Staphylococcus aureus (includes mssa and mrsa)", "Tetracycline", "*")
            df = mask(df, "Staphylococcus aureus (includes mssa and mrsa)", "Rifampin (not to be used as montherapy)", "*")
            df = mask(df, "Staphylococcus aureus (includes mssa and mrsa)", "Vancomycin", "*")
            df = mask(df, "Staphylococcus aureus (includes mssa and mrsa)", "Ciprofloxacin", "*")
        df = mask(df, "Pseudomonas aeruginosa", "Nitrofurantoin (Urine)", "N/R")
        df = mask(df, "Staphylococcus aureus MRSA", "Ampicillin", "N/R")
        df = mask(df, "Staphylococcus aureus MSSA", "Ampicillin", "N/R")
        df = mask(df, "Staphylococcus aureus MRSA", "Erythromycin (predicts azithromycin)", "N/R")
        df = mask(df, "Staphylococcus aureus MSSA", "Erythromycin (predicts azithromycin)", "N/R")
        df = mask(df, "Enterococcus spp", "Cefazolin", "N/R")
        df = mask(df, "Enterococcus faecalis", "Cefazolin", "N/R")
        df = mask(df, "Enterococcus faecium", "Cefazolin", "N/R")
        df = mask(df, "Enterococcus spp", "Cloxacillin", "N/R")
        df = mask(df, "Enterococcus faecalis", "Cloxacillin", "N/R")
        df = mask(df, "Enterococcus faecium", "Cloxacillin", "N/R")
        df = mask(df, "Enterococcus spp", "Clindamycin", "N/R")
        df = mask(df, "Enterococcus spp", "TMP/SMX", "N/R")
        if title_type != " Urine ":
            df = mask(df, "Enterococcus spp", "Ciprofloxacin", "N/R")
            df = mask(df, "Enterococcus spp", "Tetracycline", "N/R")
        else:
            df = mask(df, "Enterobacter spp.", "Fosfomycin (Oral)", "N/R")
            df = mask(df, "Klebsiella pneumoniae", "Fosfomycin (Oral)", "N/R")
            df = mask(df, "Proteus mirabilis", "Fosfomycin (Oral)", "N/R")
            df = mask(df, "Pseudomonas aeruginosa", "Fosfomycin (Oral)", "N/R")
        df = mask(df, "Enterococcus spp", "Erythromycin (predicts azithromycin)", "N/R")
        #check this too
        df = mask(df, "Enterococcus spp", "Rifampin (not to be used as montherapy)", "N/R")
        df = mask(df, "Enterobacter spp.", "Ampicillin", "N/R")
        df = mask(df, "Enterobacter spp.", "Cefazolin", "N/R")
        df = mask(df, "Enterobacter spp.", "Ceftriaxone", "N/R")
        df = mask(df, "Enterobacter spp.", "Ceftazidime", "N/R")
        df = mask(df, "Enterobacter spp.", "Piperacillin-Tazobactam", "N/R")
        df = mask(df, "Klebsiella pneumoniae", "Ampicillin", "N/R")
        return df

    #add color tags to the dataframe
    def add_tag(df):
        df = df.applymap(apply_color)
        #print(df)
        return df

    #changes column name to match footer description
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
                df = df.rename(columns={"Cefazolin": "Cefazolin (Urinary)***"})
                df = df.rename(columns={"Ciprofloxacin": "Ciprofloxacin**"})
                df = df.rename(columns={"Tetracycline": "Tetracycline**"})
            if "Gram Positive" in gp_or_gn:
                df = df.rename(columns={"Fosfomycin (Oral)": "Fosfomycin (Oral)**"})
                df = df.rename(columns={"Ciprofloxacin": "Ciprofloxacin***"})
        df = df.reset_index()
        df = df.set_index("dummy")
        return df

    #main fucntion: generates HTML document
    def generate_antibiogram_html(df, title, title_type, gp_or_gn):
        footer = add_footer(title_type, gp_or_gn)
        html = make_real_html(df, title, footer)
        html = change_tag(html)
        return html

    #generates footer
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
                footer = "<p> *** Cefazolin (urinary) predicts for cephalexin and cefprozil when used for treatment of uncomplicated UTIs due to E. coli, K. pneumoniae, and P. mirabilis but not for  therapy of infections other than uncomplicated UTIs. <p>"
                footer_2 = "<p> ** For urine only </p>"
                footer = footer_last + footer_first + footer_2 + footer 
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

    #subfunction to generate unprocessed html document
    def make_real_html(df, title, footer):
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

    #subfunction to change color tag to html class
    def change_tag(html):
        #print(type(html))
        html = html.replace("<td>red", "<td class = 'color_red'>")
        html = html.replace("<td>yellow", "<td class = 'color_yellow'>")
        html = html.replace("<td>green", "<td class = 'color_green'>")
        html = html.replace("<td>grey", "<td class = 'color_grey'>")
        html = html.replace("<th>Name", "<th class = 'hide'>Name")
        html = html.replace("N/R", "")
        return html
    
    #main function to take screenshot and crop white spaces
    def screenshot_and_crop(html, title):
        image = export_to_png(html, title)
        image, img_adr = remove_br(image, title)
        #print(image, img_adr)
        crop_image(image)
        return img_adr

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

    def remove_br(image, title):
        if "<br>" in title:
            title = title.replace("<br>", "")
            new_image = './' + title + '.png'
            os.rename(image, new_image)
            working_dir = os.getcwd()
            image_address = working_dir + "/" + title + '.png'
            return new_image, image_address
        else:
            pass

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

    #global lists
    hospital_choices = [
        'Hamilton General Hospital', 
        "McMaster University Medical Centre", 
        "Juravinski Hospital",
        "St. Peter's hospital",
        "West Lincoln Memorial Hospital"
    ]
    
    jbh_type_choices = [
        " All specimen types excluding surveillance ", 
        " All Specimens Excluding Blood Cultures and Spinal Fluids Specimemens - Streptococcus Pneumoniae", 
        " Blood Cultures and Spinal Fluids Specimens - Streptococcus Pneumoniae"]
    
    gp_or_gn_choices = [
        "Gram Positive",
        "Gram Negative",
        "Combination"
    ]

    hhs_location_choices = [
        'Hospital-Wide',
        'ICU',
        'CF Clinic' 
    ]

    hhs_type_choices = [
        ' All Specimen Types Excluding Surveillance ',
        ' Urine ',
        ' Blood Cultures ',
        ' Lower Respiratory ',
        ' All Specimens Excluding Blood Cultures and Spinal Fluids Specimemens - Streptococcus Pneumoniae',
        ' Blood Cultures and Spinal Fluids Specimens - Streptococcus Pneumoniae'
    ]

    #GUI section
    sg.SetOptions(font="any 14")

    def main_window():
        layout = [
            [sg.Button("Joseph Brant Hospital (click here)")],
            [sg.Text("Enter Year", text_color="yellow")],
            [sg.InputText(key="-YEAR-")],
            [sg.Text("Choose Hospital", text_color="yellow")],
            [[sg.Radio(text, "SITE", enable_events=True, key=f"SITE {i}")] for i, text in enumerate(hospital_choices)],
            [sg.Text("Choose Location", text_color="yellow")],
            [[sg.Radio(text, "LOC", enable_events=True, key=f"LOC {i}")] for i, text in enumerate(hhs_location_choices)],
            [sg.Text("Choose Type", text_color="yellow")],
            [[sg.Radio(text, "TYPE", enable_events=True, key=f"TYPE {i}")] for i, text in enumerate(hhs_type_choices)],
            [sg.Text("Choose one of the following (*applies to non-Strep antibiograms only)", text_color="yellow")],
            [[sg.Radio(text, "GP_OR_GN", enable_events=True, key=f"GP_OR_GN {i}")] for i, text in enumerate(gp_or_gn_choices)],
            [sg.Text("Choose the Epic output file",  text_color="yellow")],
            [sg.InputText(key="-FILE_PATH-"),
            sg.FileBrowse(initial_folder="./", file_types=(["Excel files", "*.xlsx"],))],
            [sg.Text("Enter file password below",  text_color="yellow")],
            [sg.InputText(key="-PWD-", password_char="*")],
            [sg.Button("Generate HHS Antibiogram"), sg.Exit()]
        ]
        window = sg.Window("HHS Antibiogram Generator", layout, finalize=True)
        return window
    
    def JBH_window():
        layout = [
            [sg.Text("Enter Year", text_color="yellow")],
            [sg.InputText(key="-YEAR-")],
            [sg.Text("Choose Type", text_color="yellow")],
            [[sg.Radio(text, "TYPE", enable_events=True, key=f"TYPE {i}")] for i, text in enumerate(jbh_type_choices)],
            [sg.Text("Choose one of the following (*applies to non-Strep antibiograms only)", text_color="yellow")],
            [[sg.Radio(text, "GP_OR_GN", enable_events=True, key=f"GP_OR_GN {i}")] for i, text in enumerate(gp_or_gn_choices)],
            [sg.Text("Choose the Epic output file", text_color="yellow")],
            [sg.InputText(key="-FILE_PATH-"),
            sg.FileBrowse(initial_folder="./", file_types=(["Excel files", "*.xlsx"],))],
            [sg.Text("Enter file password below", text_color="yellow")],
            [sg.InputText(key="-PWD-", password_char="*")],
            [sg.Button("Generate JBH Antibiogram"), sg.Button("Cancel")]
        ]
        window = sg.Window("JBH Antibiogram Generator", layout, relative_location=[-100, -100], finalize=True)
        return window

    window1, window2 = main_window(), None

    while True:
        window, event, values = sg.read_all_windows()
        if event in (sg.WIN_CLOSED, 'Exit', 'Cancel'):
            window.close()
            if window == window2:       # if closing win 2, mark as closed
                window2 = None
            elif window == window1:     # if closing win 1, exit program
                break
        elif event == "Generate HHS Antibiogram":
            title_year = values["-YEAR-"]
            title_site = get_selected_var(window, hospital_choices, "SITE 0", 4)
            title_location = get_selected_var(window, hhs_location_choices, "LOC 0", 11)
            title_location = " <br>" + title_location
            title_type = get_selected_var(window, hhs_type_choices, "TYPE 0", 16)
            input_gp_or_gn = get_selected_var(window, gp_or_gn_choices, "GP_OR_GN 0", 24)
            title_gp_or_gn = skip_if_strep(title_type, input_gp_or_gn)
            title = title_year + " " + title_site + title_location + title_type + title_gp_or_gn + ' Antibiogram'
            df = import_df(values["-FILE_PATH-"], values["-PWD-"])
            df = generate_antibiogram_df(df, title_type, title_gp_or_gn)
            html = generate_antibiogram_html(df, title, title_type, title_gp_or_gn)
            img_dir = screenshot_and_crop(html, title)
            sg.popup("Antibiogram generation sucessful, file created under " + img_dir)
            #print(event, values)
            #print(title)
        elif event == "Joseph Brant Hospital (click here)" and not window2:
            window2 = JBH_window()
            #print(event, values)
            #hospital_name = get_hospital_name(window, hospital_choices)
            #type_name = get_type_name(window, type_choices)
            #print(hospital_name, type_name)
        elif event == "Generate JBH Antibiogram":
            title_year = values["-YEAR-"]
            title_type = get_selected_var(window, jbh_type_choices, "TYPE 0", 3)
            input_gp_or_gn = get_selected_var(window, gp_or_gn_choices, "GP_OR_GN 0", 8)
            title_gp_or_gn = skip_if_strep(title_type, input_gp_or_gn)
            title = title_year + " Joseph Brant Hospital" + "<br>" + title_type + title_gp_or_gn + " Antibigram"
            df = import_df(values["-FILE_PATH-"], values["-PWD-"])
            df = generate_antibiogram_df(df, title_type, title_gp_or_gn)
            html = generate_antibiogram_html(df, title, title_type, title_gp_or_gn)
            img_dir = screenshot_and_crop(html, title)
            sg.popup("Antibiogram generation sucessful, file created under " + img_dir)
            #print(html)

    window.close()

if __name__ == "__main__":
    main()