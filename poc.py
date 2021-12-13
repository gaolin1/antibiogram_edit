import pandas as pd
import openpyxl as openpyxl
from selenium import webdriver
import base64
#import pdfkit
#from weasyprint import HTML, CSS

def main():
    df = import_df()
    df = convert_to_num(df)
    df = convert_isolate_to_str(df)
    df = add_tag(df)
    df = mask(df, 'E. coli', 'Tobramycin')
    df = flag_column(df)
    title = title_combine()
    footer = add_footer()
    html = make_real_html(df, title, footer)
    html = change_tag(html)
    export_to_png(html)
    #write_to_html(html)

def export_to_png(html):
    driver = webdriver.Firefox()
    html_base64 = base64.b64encode(html.encode('utf-8')).decode()
    #print(html_base64)
    driver.get("data:text/html;base64," + html_base64)
    driver.save_screenshot('test.png')
    driver.quit()


def mask(data, row_label, column_label):
    data = data.reset_index()
    data = data.set_index("Name")
    if row_label and column_label in data:
        data.loc[row_label, column_label] = "N/R"
        return data
    else:
        return data

def flag_column(df):
    df = df.rename(columns={"High-Level Gentamicin": "High-Level Gentamicin**"})
    df = df.rename(columns={"Isolates": "Number of Isolates"})
    df = df.reset_index()
    df = df.set_index("dummy")
    return df

def change_tag(html):
    #print(type(html))
    html = html.replace("<td>red", "<td class = 'color_red'>")
    html = html.replace("<td>yellow", "<td class = 'color_yellow'>")
    html = html.replace("<td>green", "<td class = 'color_green'>")
    html = html.replace("<th>Name", "<th class = 'hide'>Name")
    return html


def add_tag(df):
    df = df.applymap(apply_color)
    #print(df)
    return df

def convert_isolate_to_str(df):
    df['Isolates'] = pd.to_numeric(df['Isolates'], errors="coerce", downcast="integer")
    df['Isolates'] = df['Isolates'].fillna(0).astype(int)
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
    #print(df5)
    return df5

def apply_style(df):
    df_color = df.style.applymap(apply_color)
    return df_color

def apply_color(val):
    red = 'red' + str(val)
    green = 'green' + str(val)
    yellow = 'yellow' + str(val)
    if type(val) in [float, int]:
        if val >= 90:
            return red
        elif val >= 60:
            return green
        return yellow
    return val

def add_footer():
    footer = '<p>* Fewer than 30 isolates may not be reliable for guiding empiric treatment decisions and cannot be used to statistically compare results to another year. </p><p>** For use in combination with Ampicillin or Vancomycin for synergy. </p><p>*** N/R = Not Recommended. </p>'
    return footer

def write_to_html(html):
    write_path = input("Enter location and name for output html file: ")
    write_path = write_path + '.html'
    text_file = open(write_path, "w")
    text_file.write(html)
    text_file.close()

def title_combine():
    title_year = ask_for_year()
    title_type = ask_for_type()
    title_location = ask_for_location()
    title_site = ask_for_site()
    title = title_year + title_location + title_type + ' Antibiogram - ' + title_site
    return title

def ask_for_location():
    location = input('\n1. Hospital-Wide'
                     '\n2. ICU'
                     '\n3. CF Clinic'
                     '\nSelect location (Please enter a number): ')
    if location == "1":
        location = ' Hospital-Wide'
        return location
    if location == "2":
        location = ' ICU '
        return location
    if location == "3":
        location == ' CF Clinic'
        return location


def ask_for_year():
    year = input('\nPlease enter year for the Antibiogram (e.g. 2020): ')
    year = year + ' '
    return year


def ask_for_type():
    output_type = input('\n1. All specimen types excluding surveillance\n'
                        '2. Urine\n'
                        '3. Blood\n'
                        '4. Lower Respiratory\n'
                        'Select antibiogram type (Please enter a number): ')
    if output_type == "1":
        output_type = ' All Specimen Types Excluding Surveillance '
        return output_type
    if output_type == "2":
        output_type = ' Urine '
        return output_type
    if output_type == "3":
        output_type = ' Blood '
        return output_type
    if output_type == "4":
        output_type = ' Lower Respiratory '
        return output_type


def ask_for_site():
    title_selection = input("\n1. Hamilton General Hospital\n"
                            "2. McMaster University Medical Centre\n"
                            "3. Juravinski Hospital\n"
                            "4. St. Peter's hospital\n"
                            "5. West Lincoln Memorial Hospital\n"
                            "Select Facility (Please enter a number): ")
    if title_selection == "1":
        title_selection = 'Hamilton General Hospital'
        return title_selection
    if title_selection == "2":
        title_selection = "McMaster University Medical Centre"
        return title_selection
    if title_selection == "3":
        title_selection = "Juravinski Hospital"
        return title_selection
    if title_selection == "4":
        title_selection = "St. Peter's hospital"
        return title_selection
    if title_selection == "5":
        title_selection = "West Lincoln Memorial Hospital"
        return title_selection


def make_real_html(df, title,footer):
    message_start = f"""
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>{ title }</title>"""

    message_style = """
<style type="text/css" media="screen">
    table {
        border-collapse: collapse;
        table-layout: fixed;
        word-wrap: break-word;
    }

    tr {
        border: 1px solid;
        font-size: 12pt;
        text-align: center;
    }

    td {
        font-size: 12pt;
        border: 1px solid;
    }

    .color_red {
        background-color: red;
        text-align: center;
    }

    .color_yellow {
        background-color: yellow;
        text-align: center;
    }

    .color_green {
        background-color: green;
        text-align: center;
    }

    th {
        writing-mode: sideways-lr;
        max-width: 90px;
        min-width: 30px;
        font-size: 12pt;
        padding: 5px 5px 5px 5px;
        border: 1px solid;
    }

    td:hover {
        font-weight: bold;
        border-style: solid;
        border-width: 2px;
        border-color: black;
        font-size: 12pt;
    }

    h2 {
        text-align: left;
    }

    .hide {
        color: white;
        border: 1px solid;
        border-right-color: black;
        border-bottom-color: black;
        text-align: center;
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
    read_path = input("Enter antibiogram file path: ")
    df = pd.read_excel(read_path, "Sheet1")
    return df

if __name__ == '__main__':
    main()