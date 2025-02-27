import streamlit as st

st.title("GOAT-L'Oreal Monthly Report")
year = st.selectbox(
  'Please select the reporting year',
  ('2024', '2025')
)
month = st.selectbox(
  'Please select the reporting month',
  ('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December')
)
month_map = {
    "January": 1, "February": 2, "March": 3, "April": 4,
    "May": 5, "June": 6, "July": 7, "August": 8,
    "September": 9, "October": 10, "November": 11, "December": 12
}

month_num = month_map.get(month, "")  # Returns '' if month is not found

division = st.selectbox(
  "Please select the reporting L'Oreal Division",
  ('CPD', 'LDB', 'LLD', 'PPD')
)
category = st.selectbox(
  "Please select the reporting L'Oreal TDK Category",
  ('Hair Care', 'Female Skin', 'Make Up', 'Fragrance', 'Men Skin', 'Hair Color')
)
brands = st.multiselect(
    "Please Select 3 L'Oreal Brands to compare in the report",
    ["BLP Skin", "Garnier", "L'Oreal Paris", "GMN Shampoo Color", "Armani", "Kiehls", "Lancome", "Shu Uemura", "Urban Decay", "YSL", "Cerave", "La Roche Posay", "L'Oreal Professionel", "Matrix", "Biolage", "Kerastase", "Maybelline"]
,max_selections=3
)

#--- DATA PROCESSING ---

import os
import pandas as pd
import numpy as np

import time
import json
from google.cloud import bigquery
from google.oauth2 import service_account
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_MARKER_STYLE
from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.enum.chart import XL_LABEL_POSITION
from pptx.enum.text import MSO_ANCHOR

def format_title(slide, text, alignment, font_name, font_size, font_bold = False, font_color = RGBColor(0, 0, 0),left=Pt(75), top=Pt(25), width=Pt(850), height=Pt(70)):
    title_shape = slide.shapes.add_textbox(left=left, top=top, width=width, height=height)
    title_text_frame = title_shape.text_frame
    title_text_frame.text = text
    title_text_frame.word_wrap = True
    for paragraph in title_text_frame.paragraphs:
        paragraph.alignment = alignment
        for run in paragraph.runs:
            run.font.name = font_name   
            run.font.bold = font_bold
            run.font.color.rgb = font_color
            run.font.size = Pt(font_size)
    return title_shape

def pie_chart(slide,df,x,y,cx,cy,fontsize=15,legend_right = True, chart_title = False, title='',fontsize_title = Pt(14)):    
    df.fillna(0, inplace = True) #fill nan
    # Convert the transposed DataFrame into chart data
    chart_data = CategoryChartData()
    # Add the brand names as categories to the chart data
    for i in df.transpose().columns:
        chart_data.add_category(i)

    # Add the SOV values as series to the chart data
    for index, row in df.transpose().iterrows():
        chart_data.add_series(index, row.values)

    chart = slide.shapes.add_chart(XL_CHART_TYPE.PIE, x, y, cx, cy, chart_data).chart

    if chart_title:
        chart.has_title = True
        chart.chart_title.text_frame.text = title 
        title_font = chart.chart_title.text_frame.paragraphs[0].font
        title_font.bold = True
        title_font.size = fontsize_title
        # chart.chart_title.text_frame.paragraphs[0].font.size = Pt(10)  # Set font size to 24pt
        # chart.chart_title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0,0,0)  # Set font color to black
    else:
        chart.has_title = False
        
    chart.has_legend = True
    chart.legend.include_in_layout = False
    chart.legend.position = XL_LEGEND_POSITION.RIGHT if legend_right else XL_LEGEND_POSITION.BOTTOM 
    # Set legend font size to 10 points
    chart.legend.font.size = Pt(fontsize) 


    for series in chart.plots[0].series:
        for i, val in enumerate(series.values):
            if val == 0:
                series.points[i].data_label.has_text_frame = True
            series.data_labels.show_value = True
            series.data_labels.font.size = Pt(fontsize)
            series.data_labels.number_format = '0%'
            series.data_labels.position = XL_LABEL_POSITION.BEST_FIT
            
    # chart.chart_title.text_frame.text = chart.chart_title.text_frame.text + 'budget distribution'
    # chart.chart_title.text_frame.paragraphs[0].font.size = Pt(fontsize)
    # chart.chart_title.text_frame.color.rgb = RGBColor(0,0,0)
          
    return chart


def line_marker_chart(slide,df,x,y,cx,cy, legend = True, legend_position = XL_LEGEND_POSITION.RIGHT, data_show = False, chart_title = False, title ="", fontsize = Pt(9), fontsize_title = Pt(14), percentage = False, line_width = Pt(1), smooth=False):
    df.fillna(0, inplace = True) #fill nan
    # Define chart data
    chart_data = CategoryChartData()
    # for i in df.columns:
    #     chart_data.add_category(i)
    # for j, row in df.iterrows():
    #     chart_data.add_series(j, row.values)
    for i in df.columns:
        chart_data.add_category(i)
    for j, row in df.iterrows():
        chart_data.add_series(j,np.where(row.values == 0, None, row.values))
    
    chart = slide.shapes.add_chart(XL_CHART_TYPE.LINE, x, y, cx, cy, chart_data).chart
    
    # smoothing
    if smooth: # Smooth the lines
        for series in chart.series:
            series.smooth = True

    
    # Add legend
    if legend:
        chart.has_legend = True
        chart.legend.include_in_layout = False
        chart.legend.position = legend_position
        chart.legend.font.size = fontsize  # assuming Pt is imported from pptx.util
    else:
        chart.has_legend = False
    
    if data_show:
        for series in chart.plots[0].series:
            for i, val in enumerate(series.values):
                if val == 0:
                    series.points[i].data_label.has_text_frame = True
                series.data_labels.show_value = True
                series.data_labels.font.size = fontsize
                # series.data_labels.number_format = '0.00%'
                series.data_labels.position = XL_LABEL_POSITION.ABOVE
                
    # Change line point to all circle
    # Iterate through each series and set marker style to circle
    for series in chart.series:
        series.marker.style = XL_MARKER_STYLE.CIRCLE
        series.format.line.width = line_width # custom line width

    # Customize y-axis format
    chart.value_axis.tick_labels.font.size = fontsize  # Set font size for tick labels

    # Set font size for category axis (months)
    chart.category_axis.tick_labels.font.size = fontsize  # Set font size for category axis labels
    
    # Find the maximum value across all series
    max_value = 0
    for series in chart.plots[0].series:
        try:
            series_max = max(series.values)
        except:
            series_max = 0
        max_value = max(max_value, series_max)  

    if max_value >= 1000:
        chart.value_axis.tick_labels.number_format = '#,##0'  # add commas to separate thousands
    
    if chart_title:
        chart.chart_title.text_frame.text = title # Set the title text   
        title_font = chart.chart_title.text_frame.paragraphs[0].font
        title_font.bold = True
        title_font.size = fontsize_title
    else:
        chart.has_title = False

    # Remove Gridlines (Line Chart Specific)
    value_axis = chart.value_axis
    category_axis = chart.category_axis
    value_axis.has_major_gridlines = False
    value_axis.has_minor_gridlines = False
    category_axis.has_major_gridlines = False
    category_axis.has_minor_gridlines = False
    
    # if percentage
    if percentage:
        category_axis.tick_labels.NumberFormat = '0"%"'

    return chart

def table_default(slide, df, left, top, width, height, width_row, height_row, header=True, upper=False, fontsize=10, alignment=PP_ALIGN.LEFT): # Function to add and format a table on a slide
    table_data = df.values.tolist()
    if header:
        header = df.columns.values.tolist()
        table_data.insert(0,header)

    table = slide.shapes.add_table(rows=len(table_data), cols=len(table_data[0]), left=left, top=top, width=width, height=height).table

    # Fill table data
    for i, row in enumerate(table_data):
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            if isinstance(val, (int, float)):  # Check if it's a number
                text = f"{val:,}"  # Format with comma separators
            else:
                text = str(val).upper() if upper else str(val)
            cell.text = text
            # for i in range(len(cell.text_frame.paragraphs)):
            #     p = cell.text_frame.paragraphs[i]
            #     p.font.size = Pt(fontsize)  # Set font size 
            #     p.alignment = alignment
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = alignment
                for run in paragraph.runs:
                    # run.font.name = font_name   
                    # run.font.bold = font_bold
                    # run.font.color.rgb = font_color
                    run.font.size = Pt(fontsize)
            cell.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE  
            # Set all margins to 0
            cell.text_frame.margin_left = 0
            cell.text_frame.margin_right = 0
            cell.text_frame.margin_top = 0
            cell.text_frame.margin_bottom = 0
            
        # Set the width of each column
        for i, column in enumerate(table.columns):
            column.width = width_row[i]

        # Adjust the height of each row
        for row in table.rows:
            row.height = height_row # Set the height of each row 

    return table

def adjust_dataframe(df, columns, index=False):
    
  # Combine existing and desired columns/index
  all_columns = list(df.columns) + [col for col in columns if col not in df.columns]
  if index == False:
    df = df.reindex(columns=all_columns,fill_value=np.nan)
  else:
    all_index = list(df.index) + [idx for idx in index if idx not in df.index]
    # Reindex to include all columns and index values
    df = df.reindex(columns=all_columns, index=all_index, fill_value=np.nan)

  return df

# Load credentials from Streamlit Secrets
credentials_dict = st.secrets["gcp_service_account"]
credentials = service_account.Credentials.from_service_account_info(credentials_dict)

# Initialize BigQuery Client
client = bigquery.Client(credentials=credentials, project=credentials.project_id)
   
# Example Query
query = """
SELECT 
Brand
,SUM(Views) as Views
,SUM(Engagement) as Engagement
FROM loreal-id-prod.loreal_storage.advocacy_tdk
WHERE TDK_Category = '{}'
AND datename(month, Date) = '{}'
AND YEAR(Date) = '{}'
GROUP BY
Brand
""".format(category, month_num, year)

# Fetch data
df = client.query(query).to_dataframe()

# Display results
st.write(df)


