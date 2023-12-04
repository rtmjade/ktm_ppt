#install pdtoppt
# !pip install git+https://github.com/robintw/PandasToPowerpoint.git

import os
import sys
import datetime
import pandas as pd
import numpy as np

sys.path.insert(0, '../assets/scripts')
from ppt_class import Ppt

Ppt.template = '../assets/Slide Template.pptx'

slide_title = 'Key Talent Management Profile Tracking'

today = datetime.datetime.today().strftime(format='%b-%Y')

example = Ppt(title=slide_title, subtitle1=today)


def add_slide_for_key_talent(x):
    slide = example.add_slide(f'Key Talent Profile {name} - Shopee Manager', layout=1)
    slide.add_textbox(f"{x['name']}", [10.3, 2.6, 10.46, 0.86], halign='left', fit=(True, 12))
    slide.add_textbox(f"{x['Designation']}", [10.3, 2.6, 10.46, 0.86], halign='left', fit=(True, 12))
    slide.add_textbox(f"{x['name']}", [10.3, 2.6, 10.46, 0.86], halign='left', fit=(True, 12))
    slide.add_textbox(f"{x['name']}", [10.3, 2.6, 10.46, 0.86], halign='left', fit=(True, 12))

for name in ['Keng Whye LONG TEXTLONG TEXTLONG TEXTLONG TEXTLONG TEXTLONG TEXTLONG TEXTLONG TEXTLONG TEXTLONG TEXTLONG TEXT', 'Hoang', 'Chuye']:
    slide = example.add_slide(f'Key Talent Profile {name} - Shopee Manager', layout=1)
    slide.add_textbox(f'{name}', [10.3, 2.6, 10.46, 0.86], cm=True, halign='left', fit=(True, 12))

df = pd.DataFrame(np.random.randint(0, 100,size=(100, 4)), columns=list('ABCD'))

'''Out of Green Zone Slides'''
size = 14
# Split data into sets of 14 entries
lists_of_dfs = [df.loc[i:i+size-1,:] for i in range(0, len(df), size)]

# For each dataframe, create a slide
for _ in range(0, len(lists_of_dfs)):
    table_slide = example.add_slide(f'Slide Number {_}', layout=3)
    table_slide.add_df_to_table(lists_of_dfs[_],
                                size,
                                columns_width=[0.53, 7.13, 2.42, 2.42],
                                headers_height=0.38)
example.save('Example.pptx')
