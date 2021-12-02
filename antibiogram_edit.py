import pandas as pd
import openpyxl as openpyxl
#from bs4 import BeautifulSoup
#import dataframe_image as dfi


def main():
    df = import_df()
    # print(df)
    df = convert_to_num(df)
    df = convert_isolate_to_str(df)
    df = mask(df, 'E. coli organism grouper', 'TOB -  TOBRAMYCIN')
    df = flag_column(df)
    #print(df)
    title = title_combine()
    df_color = apply_style(df,title)  
    html = df_color.render()    
    html = add_footer(html)
    to_html(html)

"useless stuff below"
    # write_to_html(html)
    # dfi.export(df_color, 'D:\Downloads\color.png')

    # export_df(df_color)

def apply_style(df,title):
    df_color = df.style.set_caption(title) \
        .applymap(apply_color) \
        .set_table_styles(style()) \
        # .applymap(apply_color) \
    return df_color


def add_footer(html):
    html = html + '<p>' \
                  '* Fewer than 30 isolates may not be reliable for guiding empiric ' \
                  'treatment decisions and cannot be used to statistically compare results' \
                  'to another year.' \
                  '</p>' \
                  '<p>' \
                  '** For use in combination with Ampicillin or Vancomycin for synergy.' \
                  '</p>'
    return html


def title_combine():
    title_year = ask_for_year()
    title_type = ask_for_type()
    title_location = ask_for_location()
    title_site = ask_for_site()
    title = title_year + title_location + title_type + ' Antibiogram - ' + title_site
    return title


def flag_column(df):
    df = df.rename(columns={"HLG - High-Level Gentamicin": "HLG - High-Level Gentamicin**"})
    return df


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


def write_to_html(html):
    write_path = input("Enter location and name for output html file: ")
    write_path = write_path + '.html'
    text_file = open(write_path, "w")
    text_file.write(html)
    text_file.close()


def mask(data, row_label, column_label):
    if row_label and column_label in data:
        data.loc[row_label, column_label] = "Not recommended"
        return data
    else:
        return data


def apply_color(val):
    red = 'background-color: red; text-align: center;font-weight: bold; border-style:solid ;border-width: 2.5px'
    green = 'background-color: green; text-align: center;font-weight: bold; border-style:solid ;border-width: 2.5px'
    yellow = 'background-color: yellow; text-align: center;font-weight: bold; border-style:solid ;border-width: 2.5px'
    default = 'background-color: #f7f7f9'
    if type(val) in [float, int]:
        if val >= 90:
            return green
        elif val >= 60:
            return yellow
        return red
    return default


def convert_to_num(df):
    df2 = df.apply(lambda x: pd.to_numeric(x.astype(str).str.replace(',', ''), errors='coerce', downcast="integer"))
    df3 = df2.fillna(999)
    df4 = df3.astype(int)
    df5 = df4.replace(to_replace=999, value=' ')
    return df5


def import_df():
    read_path = input("Enter antibiogram file path: ")
    df = pd.read_excel(read_path, "Sheet1", index_col='Name')
    return df


def export_df(df):
    write_path = input("Enter location for output html file: ")
    df.to_excel(write_path, sheet_name='Sheet1')


def convert_isolate_to_str(df):
    df['Isolates'] = pd.to_numeric(df['Isolates'], errors="coerce", downcast="integer")
    df['Isolates'] = df['Isolates'].fillna(0).astype(int)
    df['Isolates'] = df['Isolates'].replace(to_replace=0, value=' ')
    df['Isolates'] = df['Isolates'].astype(str)
    # print(df['Isolates'].head())
    return df


def to_html(html):
    write_path = input("Enter location and name for output html file: ")
    write_path = write_path + '.html'
    text_file = open(write_path, "w")
    text_file.write(html)
    text_file.close()


def style():
    def magnify():
        return [dict(selector="th",
                     props=[("font-size", "4pt")]),
                dict(selector="td",
                     props=[('padding', "0em 0em")]),
                dict(selector="th:hover",
                     props=[("font-size", "12pt")]),
                dict(selector="tr:hover td:hover",
                     props=[('max-width', '200px'),
                            ('font-size', '12pt')])
                ]

    th_props = [
        ('font-size', '15px'),
        ('text-align', 'center'),
        ('font-weight', 'bold'),
        # ('color', '#6d6d6d'),
        ('background-color', '#f7f7f9'),
        ('border-style', 'solid'),
        ('border-width', '1px'),
        ('padding', '4px 4px'),
        ("border-collapse", "collapse"),
        ('margin', '0px'),
    ]

    # Set CSS properties for td elements in dataframe
    td_props = [
        ('font-size', '15px'),
        ('border-width', '0.01em'),
        ('border-style', 'solid'),
        # ("border-collapse", "collapse"),
        ('margin', '0px')
    ]

    caption_props = [
        ('font-size', '20px'),
        ('text-align', 'center'),
        ('font-weight', 'bold'),
    ]

    index_props = [
        ('color', 'darkgrey'),
    ]

    table_props = [
        ('border-collapse', 'collapse'),
        ('table-layout', 'fixed'),
        ('width', '100%'),
        ('word-wrap', 'break-word')
    ]

    # Set table styles
    styles = [
        dict(selector="th", props=th_props),
        dict(selector="td", props=td_props),
        # dict(selector="td:hover th", props=[("background-color", '#ffffb3')]),
        dict(selector="td:hover", props=[("background-color",
                                          'background-color: red; '
                                          'font-color: #ffffff'
                                          'text-align: center;'
                                          'font-weight: bold; '
                                          'border-style: solid;'
                                          'border-width: 3px;'
                                          'border-color: black;'
                                          'font-size: 16pt')]),
        dict(selector="tr:hover th", props=[("background-color", '#ffffb3')]),
        dict(selector="th:hover", props=[("background-color", '#ffffb3')]),
        dict(selector='caption', props=caption_props),
        dict(selector='tr: .index_name', props=index_props),
        dict(selector="", props=table_props),
        #dict(selector="", props=('border-collapse', 'collapse'))
        # dict(selector="tr:hover td", props=[("background-color", '#ffffb3')]),
        # magnify()

    ]
    return styles


if __name__ == '__main__':
    main()
