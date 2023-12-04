import pandas as pd
import sys
import datetime as dt
import numpy as np

sys.path.insert(0, '../assets/scripts')
from ppt_class import Ppt

Ppt.template = '../assets/Slide Template.pptx'

# To create the First slide
slide_title = 'Key Talent Management Profile Tracking'
today = dt.datetime.today().strftime(format='%b-%Y')
example = Ppt(title=slide_title, subtitle1=today)

Ppt.header_dimensions = [2.46, 0.91, 28.29, 1.29]


''' To create Table slides '''
df = pd.read_excel('With dummy data_Key talent list_v1_4 Jan 2023.xlsx')

# To calculate tenure from group join date
df['Today'] = pd.Timestamp.today().strftime('%Y-%m-%d')
df['Today'] = pd.to_datetime(df['Today'])
df['Group Join Date'] = pd.to_datetime(df['Group Join Date'])
df['Tenure (Year)'] = ((df['Today']-df['Group Join Date'])/np.timedelta64(1, 'Y')).round(2)

# To get the reporting line
df['Country Head'] = df['Reporting Line'].str.extract('.*>(.*)', expand=False).str.strip()

# To create table slides
table_df = df[['Name', 'Region/Office', 'Business Unit', 'Rank', 'Tenure (Year)', 'Duration in Current Rank',
               'Country Head', 'CountryHead/HOD', '2022 Grade', 'Long-Term Potential']]

size = 14
list_of_dfs = [table_df.loc[i:i+size-1, :] for i in range(0, len(table_df), size)]

for _ in range(0, len(list_of_dfs)):
    table_slide = example.add_slide(f'Overview - Page {_+1}', layout=3) #starts with 0
    table_slide.add_df_to_table(list_of_dfs[_],
                                size,
                                columns_width=[1.25, 1.25, 1.25, 1.25, 1.25, 1.25, 1.25, 1.25, 1.25, 1.25],
                                headers_height=0.5)

# To fill in text boxes

# To avoid .0 at the end of year value by turning foat
df['Highest Grad Year'] = df['Highest Grad Year'].astype("Int64")
df['First Grad Year'] = df['First Grad Year'].astype("Int64")

df['First Grad Year'] = df['First Grad Year'].apply(str)
df['Highest Grad Year'] = df['Highest Grad Year'].apply(str)


# To get the Academic Background
df['First Education'] = df['First Degree'] + ' of ' + df['First Discipline'] + ' in ' + \
                        df['First University'] + ' (' + df['First Grad Year'] + ')'

df['Highest Education'] = df['Highest Degree']
# + ' of ' + df['Highest Discipline'] + ' in ' + df['Highest University'] + ' (' + df['Highest Grad Year'] + ')'

def concat(*args):
    strs = [str(arg) for arg in args if not pd.isnull(arg)]
    return ' '.join(strs) if strs else np.nan
np_concat = np.vectorize(concat)

df['my_col'] = np_concat(df['First Degree'].astype(str),  'in ' + df['First Discipline'].astype(str))

df['Highest Education'] = df['Highest Education'].fillna('')
df['First Education'] = df['First Education'].fillna('')

def create_slide(x):

    name = x['Name']
    slide = example.add_slide(f'Key Talent Profile - {name}', layout=1)
    slide.add_textbox(f'{name}', [4, 1, 4.12, 0.37], cm=False, halign='left', fit=(True, 12))

    designation = x['Title']
    slide.add_textbox(f'{designation}', [4, 1.33, 4.12, 0.37], cm=False, halign='left', fit=(True, 12))

    location = x['Region/Office']
    slide.add_textbox(f'{location}', [4, 1.63, 4.12, 0.37], cm=False, halign='left', fit=(True, 12))

    reporting_manager = x['Country Head']
    slide.add_textbox(f'{reporting_manager}', [4.01, 1.98, 4.12, 0.37], cm=False, halign='left', fit=(True, 12))

    country_head_hod = x['CountryHead/HOD']
    slide.add_textbox(f'{country_head_hod}', [4.01, 2.29, 4.12, 0.37], cm=False, halign='left', fit=(True, 12))

    total_team_size = x['Team Size (# of direct report)']
    slide.add_textbox(f'{total_team_size}', [10.3, 1, 2.66, 0.36], cm=False, halign='left', fit=(True, 12))

    tenure = x['Tenure (Year)']
    slide.add_textbox(f'{tenure} years', [10.3, 1.33, 2.66, 0.27], cm=False, halign='left', fit=(True, 12))

    age = x['Age']
    slide.add_textbox(f'{age}', [10.3, 1.68, 2.66, 0.27], cm=False, halign='left', fit=(True, 12))

    grade2022 = x['2022 Grade']
    slide.add_textbox(f'{grade2022}', [10.3, 1.98, 2.66, 0.27], cm=False, halign='left', fit=(True, 12))

    long_term_potential = x['Long-Term Potential']
    slide.add_textbox(f'{long_term_potential}', [10.3, 2.29, 2.66, 0.27], cm=False, halign='left', fit=(True, 12))

    if pd.isnull(x['First University']) or (pd.isnull(x['First Degree']) and pd.isnull(x['First Discipline'])):
        first_degree = None
    elif x['First University'] and x['First Discipline'] and x['First Degree']:
        first_degree = f"{x['First Degree']} of {x['First Discipline']} in {x['First University']}" if pd.isnull(x['First Grad Year']) else \
            f"{x['First Degree']} of {x['First Discipline']} in {x['First University']} ({x['First Grad Year']})"
    elif x['First University'] and x['First Degree'] and pd.isnull(x['First Discipline']):
        first_degree = f"{x['First Degree']} in {x['First University']}" if pd.isnull(x['First Grad Year']) else \
            f"{x['First Degree']} in {x['First University']} ({x['First Grad Year']})"
    elif x['First University'] and x['First Discipline'] and pd.isnull(x['First Degree']):
        first_degree = f"{x['First Discipline']} in {x['First University']}" if pd.isnull(x['First Grad Year']) else \
            f"{x['First Discipline']} in {x['First University']} ({x['First Grad Year']})"

    if first_degree:
        slide.add_textbox(f'{first_degree}', [1.24, 3.08, 3.44, 0.41], cm=False, halign='left', fit=(True, 12))

    highest_degree = x['Highest Education']
    slide.add_textbox(f'{highest_degree}', [1.24, 3.52, 3.44, 0.41], cm=False, halign='left', fit=(True, 12))

    career_path = x['Career Path and Past Experiences']
    slide.add_textbox(f'{career_path}', [1.24, 3.92, 3.44, 1.72], cm=False, halign='left', fit=(True, 12))

    performance_goal = x['2023 Performance Goals']
    slide.add_textbox(f'{performance_goal}', [1.24, 5.82, 3.44, 1.33], cm=False, halign='left', fit=(True, 12))

    strengths = x['HOD/ HOD-1 general remarks on individual\'s potential and strengths']
    slide.add_textbox(f'{strengths}', [6.98, 3.1, 5.92, 0.72], cm=False, halign='left', fit=(True, 12))

    development_area = x['HOD/ HOD-1 general remarks on individual\'s improvement areas']
    slide.add_textbox(f'{development_area}', [6.98, 3.88, 5.95, 1.38], cm=False, halign='left', fit=(True, 12))

    current_role = x['Current Role']
    slide.add_textbox(f'{current_role}', [6.98, 5.74, 5.95, 0.32], cm=False, halign='left', fit=(True, 12))

    role_in_6_months = x['Potential Role (in 6 months\' time)']
    slide.add_textbox(f'{role_in_6_months}', [6.98, 6.14, 5.95, 0.32], cm=False, halign='left', fit=(True, 12))

    role_in_1_2_years = x['Potential Role (in 1-2 years\' time)']
    slide.add_textbox(f'{role_in_1_2_years}', [6.98, 6.52, 5.95, 0.32], cm=False, halign='left', fit=(True, 12))

df.apply(create_slide, axis=1)

example.save('Hoang_test.pptx')