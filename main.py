import pandas as pd
import sys
import datetime as dt
import numpy as np
import os
import urllib
import urllib.request


sys.path.insert(0, '../Key Talents Management/powerpoint-class/assets/scripts')
from ppt_class import Ppt

Ppt.template = '../Key Talents Management/powerpoint-class/assets/Slide Template.pptx'

# To create the First slide
slide_title = 'Key Talent Management Profile Tracking'
today = dt.datetime.today().strftime(format='%b-%Y')
example = Ppt(title=slide_title, subtitle1=today)

Ppt.header_dimensions = [1.42, 0.48, 21.76, 1.45]


''' To create Table slides '''
df = pd.read_excel('C:/Work/People Analytics/Key Talents Management/With dummy data_Key talent list_2023.xlsx')

# To calculate tenure from group join date
df['Today'] = pd.Timestamp.today().strftime('%Y-%m-%d')
df['Today'] = pd.to_datetime(df['Today'])
df['Group Join Date'] = pd.to_datetime(df['Group Join Date'])
df['Group Join Date (alt)'] = df['Group Join Date'].dt.strftime("%b %Y")

# To create table slides
table_df = df[['Name', 'Region/ Office', 'Business Unit', 'Rank', 'Tenure (Years)', 'Duration in Current Rank',
               'Reporting Manager', 'CountryHead/HOD', '2022 Grade', 'Long-Term Potential (H1 2023 - Updated)']]

table_df.index = np.arange(1, len(df) + 1)
table_df['#'] = table_df.index
table_df['Location'] = table_df['Region/ Office']
table_df['2022 Performance'] = table_df['2022 Grade']
table_df['Tenure current rank (Years)'] = table_df['Duration in Current Rank']
table_df['2023 Long-Term Potential'] = table_df['Long-Term Potential (H1 2023 - Updated)']
table_df = table_df[['#', 'Name', 'Location', 'Business Unit', 'Rank', 'Tenure (Years)', 'Tenure current rank (Years)',
                     'Reporting Manager', 'CountryHead/HOD', '2022 Performance', '2023 Long-Term Potential']]

size = 14
list_of_dfs = [table_df.loc[i:i+size-1, :] for i in range(0, len(table_df), size)]

for _ in range(0, len(list_of_dfs)):
    table_slide = example.add_slide(f'Overview - Page {_+1}', layout=6)
    table_slide.add_df_to_table(list_of_dfs[_],
                                size,
                                columns_width=[0.5, 0.88, 0.88, 0.88, 0.88, 0.88, 0.88, 0.88, 0.88, 1, 1],
                                headers_height=0.5,
                                dimensions=[0.14, 0.77, 9.67, 1.29])

# To fill in text boxes

# To avoid .0 at the end of year value by turning foat
df['Highest Grad Year'] = df['Highest Grad Year'].astype("Int64")
df['First Grad Year'] = df['First Grad Year'].astype("Int64")

df['First Grad Year'] = df['First Grad Year'].apply(str)
df['Highest Grad Year'] = df['Highest Grad Year'].apply(str)



def create_slide(x):

    name = x['Name']
    slide = example.add_slide(f'Key Talent Profile - {name}', layout=5)
    slide.add_textbox(f'{name}', [3.03, 0.73, 2.29, 0.34], cm=False, halign='left', fit=(True, 9))

    designation = x['Title']
    slide.add_textbox(f'{designation}', [3.03, 0.97, 2.29, 0.34], cm=False, halign='left', fit=(True, 9))

    location = x['Region/ Office']
    slide.add_textbox(f'{location}', [3.03, 1.22, 2.29, 0.34], cm=False, halign='left', fit=(True, 9))

    reporting_manager = x['Reporting Manager']
    slide.add_textbox(f'{reporting_manager}', [3.03, 1.47, 2.29, 0.34], cm=False, halign='left', fit=(True, 9))

    country_head_hod = x['CountryHead/HOD']
    slide.add_textbox(f'{country_head_hod}', [3.03, 1.71, 2.29, 0.34], cm=False, halign='left', fit=(True, 9))

    total_team_size = x['Team Size (# of direct report)']
    slide.add_textbox(f'{total_team_size}', [6.87, 0.74, 2.75, 0.29], cm=False, halign='left', fit=(True, 9))

    tenure = x['Tenure (Years)']
    group_join_date = x['Group Join Date (alt)']
    slide.add_textbox(f'{group_join_date} ({tenure} yrs)', [6.87, 0.98, 2.75, 0.29], cm=False, halign='left', fit=(True, 9))

    age = x['Age']
    slide.add_textbox(f'{age}', [6.87, 1.22, 2.75, 0.29], cm=False, halign='left', fit=(True, 9))

    grade2022 = x['2022 Grade']
    slide.add_textbox(f'{grade2022}', [6.87, 1.47, 2.75, 0.29], cm=False, halign='left', fit=(True, 9))

    long_term_potential = x['Long-Term Potential (H1 2023 - Updated)']
    slide.add_textbox(f'{long_term_potential}', [6.87, 1.7, 2.75, 0.29], cm=False, halign='left', fit=(True, 9))

    if pd.isnull(x['First University']) or (pd.isnull(x['First Degree']) and pd.isnull(x['First Discipline'])):
        first_degree = None
    elif pd.notnull(x['First University']) and pd.notnull(x['First Discipline']) and pd.notnull(x['First Degree']):
        first_degree = f"{x['First Degree']} of {x['First Discipline']} in {x['First University']}" if pd.isnull(x['First Grad Year']) else \
            f"{x['First Degree']} of {x['First Discipline']} in {x['First University']} ({x['First Grad Year']})"
    elif pd.notnull(x['First University']) and pd.notnull(x['First Degree']) and pd.isnull(x['First Discipline']):
        first_degree = f"{x['First Degree']} in {x['First University']}" if pd.isnull(x['First Grad Year']) else \
            f"{x['First Degree']} in {x['First University']} ({x['First Grad Year']})"
    elif pd.notnull(x['First University']) and pd.notnull(x['First Discipline']) and pd.isnull(x['First Degree']):
        first_degree = f"{x['First Discipline']} in {x['First University']}" if pd.isnull(x['First Grad Year']) else \
            f"{x['First Discipline']} in {x['First University']} ({x['First Grad Year']})"

    if first_degree:
        slide.add_textbox(f'{first_degree}', [1.19, 2.3, 2.36, 0.46], cm=False, halign='left', fit=(True, 9))

    if pd.isnull(x['Highest University']) or (pd.isnull(x['Highest Degree']) and pd.isnull(x['Highest Discipline'])):
        highest_degree = None
    elif pd.notnull(x['Highest University']) and pd.notnull(x['Highest Discipline']) and pd.notnull(x['Highest Degree']):
        highest_degree = f"{x['Highest Degree']} of {x['Highest Discipline']} in {x['Highest University']}" if pd.isnull(x['Highest Grad Year']) else \
            f"{x['Highest Degree']} of {x['Highest Discipline']} in {x['Highest University']} ({x['Highest Grad Year']})"
    elif pd.notnull(x['Highest University']) and pd.notnull(x['Highest Degree']) and pd.isnull(x['Highest Discipline']):
        highest_degree = f"{x['Highest Degree']} in {x['Highest University']}" if pd.isnull(x['Highest Grad Year']) else \
            f"{x['Highest Degree']} in {x['Highest University']} ({x['Highest Grad Year']})"
    elif pd.notnull(x['Highest University']) and pd.notnull(x['Highest Discipline']) and pd.isnull(x['Highest Degree']):
        highest_degree = f"{x['Highest Discipline']} in {x['Highest University']}" if pd.isnull(x['Highest Grad Year']) else \
            f"{x['Highest Discipline']} in {x['Highest University']} ({x['Highest Grad Year']})"

    if highest_degree:
        slide.add_textbox(f'{highest_degree}', [1.19, 2.75, 2.36, 0.45], cm=False, halign='left', fit=(True, 9))

    career_path = x['Career Path and Past Experiences']
    slide.add_textbox(f'{career_path}', [1.19, 3.24, 2.36, 0.92], cm=False, halign='left', fit=(True, 9))

    performance_goal = x['2023 Performance Goals']
    slide.add_textbox(f'{performance_goal}', [1.19, 4.22, 2.36, 1.16], cm=False, halign='left', fit=(True, 9))

    strengths = x['HOD/ HOD-1 general remarks on individual\'s potential and strengths']
    slide.add_textbox(f'{strengths}', [5.2, 2.37, 4.55, 0.57], cm=False, halign='left', fit=(True, 9))

    development_area = x['HOD/ HOD-1 general remarks on individual\'s improvement areas']
    slide.add_textbox(f'{development_area}', [5.2, 2.93, 4.5, 0.57], cm=False, halign='left', fit=(True, 9))

    current_role = x['Current Role']
    duration = x['Duration in Current Role']
    slide.add_textbox(f'(Duration in current role: {current_role}) {duration} years', [5.2, 3.77, 4.59, 0.38], cm=False, halign='left', fit=(True, 9))

    role_in_6_months = x['Potential Role (in 6 months\' time)']
    slide.add_textbox(f'{role_in_6_months}', [5.2, 4.13, 4.59, 0.38], cm=False, halign='left', fit=(True, 9))

    role_in_1_2_years = x['Potential Role (in 1-2 years\' time)']
    slide.add_textbox(f'{role_in_1_2_years}', [5.2, 4.5, 4.59, 0.38], cm=False, halign='left', fit=(True, 9))

    development_plan_2023 = x['Development Plan 2023']
    slide.add_textbox(f'{development_plan_2023}', [5.2, 4.87, 4.59, 0.38], cm=False, halign='left', fit=(True, 9))

# to add picture

    def photo_download(x):
        url = x['Photolink']
        name = x['HRIS ID']
        with urllib.request.urlopen(f'{url}') as fh:
            with open(f'../Key Talents Management/Photos/{name}.jpg', 'wb') as out:
                out.write(fh.read())

    df.apply(photo_download, axis=1)

    photo = x['HRIS ID']
    try:
        slide.add_picture(f'../Key Talents Management/Photos/{photo}.jpg', [0.2, 0.75, 1.29, 1.22])
    except FileNotFoundError:
        slide.add_textbox('Photo not found', [0.2, 0.75, 1.29, 1.22], cm=False)

df.apply(create_slide, axis=1)

example.save('Hoang_test.pptx')
os.startfile('Hoang_test.pptx')

