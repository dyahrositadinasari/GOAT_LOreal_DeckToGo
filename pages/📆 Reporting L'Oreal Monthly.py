import streamlit as st
import os
import pandas as pd
import numpy as np
import time
from datetime import datetime
import json
import smtplib
import pytz
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
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
from pptx.enum.shapes import MSO_SHAPE

st.logo(
    "image/goat_logo.png", size="medium", link=None, icon_image=None
)
st.image("image/goat-agency-logo2.png")
st.header("L'Oreal Monthly Report")
year = st.selectbox(
  'Please select the reporting year',
  ('2024', '2025')
)
year_ = {
  '2024':24, '2025':25
}

year_map = year_.get(year, "")  # Returns '' if year is not found
year_range = (year_map -1, year_map)

quarter_ = st.selectbox(
  'Please select the reporting quarter',
  ('Quarter 1', 'Quarter 2', 'Quarter 3', 'Quarter 4')
)
q_map = {
  'Quarter 1': 'Q1', 'Quarter 2': 'Q2', 'Quarter 3': 'Q3', 'Quarter 4': 'Q4'
}
quarter = q_map.get(quarter_, "")  # Returns '' if quarter is not found

month = st.selectbox(
  'Please select the reporting month',
  ('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec')
)
month_map = {
    "Jan": "1", "Feb": "2", "Mar": "3", "Apr": "4",
    "May": "5", "Jun": "6", "Jul": "7", "Aug": "8",
    "Sep": "9", "Oct": "10", "Nov": "11", "Dec": "12"
}

month_num = month_map.get(month, "")  # Returns '' if month is not found

division_selection = st.multiselect(
  "Please select the reporting L'Oreal Division",
  ['Select All', 'CPD', 'LDB', 'LLD', 'PPD']
)
if 'Select All' in division_selection:
    division = ['CPD', 'LDB', 'LLD', 'PPD']
else: division = division_selection
	
category_selection = st.multiselect(
  "Please select the reporting L'Oreal Category",
  ['Select All', 'Pro Hair', 'Skincare', 'Makeup', 'Fragrance', 'Men Skincare', 'Haircolor', 'Haircare']
)
if 'Select All' in category_selection:
    category = ['Pro Hair', 'Skincare', 'Makeup', 'Fragrance', 'Men Skincare', 'Haircolor', 'Haircare']
else: category = category_selection

brands = st.multiselect(
    "Please Select max 5 L'Oreal Brands to compare in the report",
	["3CE", "ARM", "CRV", "GAR", "KER", "KIE", "LAN", "LP", "LRP", "MNY", "MXBIO", "OAP", "SHU", "YSL"]
	,max_selections=5
)
# brand full name -> ["BLP Skin", "Garnier", "L'Oreal Paris", "GMN Shampoo Color", "Armani", "Kiehls", "Lancome", "Shu Uemura", "Urban Decay", "YSL", "Cerave", "La Roche Posay", "L'Oreal Professionel", "Matrix", "Biolage", "Kerastase", "Maybelline"]

email_list = st.multiselect(
	"Please select the email address you wish to send this report to",
	("dyah.dinasari@groupm.com", "bandi.wijaya@groupm.com", "yahya.wahyu@groupm.com", "ana.maratu@groupm.com", "aldi.firstanto@groupm.com", "muhammad.ilham@groupm.com", )
	)

st.badge(" Optional", icon="⚠️", color="blue")

uploaded_file = st.file_uploader("Upload the PPT Template for this report, make sure the page is 16 pages or more.", type=["pptx"])
if uploaded_file is not None:
	st.success("File uploaded successfully!")
else:
	uploaded_file = "pages/Template Deck to Go - L'Oreal Indonesia.pptx"




#st.write("Selected Year : ", year)
#st.write("Selected Year_ : ", year_map)
#st.write("Selected Month : ", month)
#st.write("Month Number : ", month_num)
#st.write("Division : ", division)
#st.write("Category : ", category)
#st.write("Brands : ", brands)
#st.write("Email list : ", email_list)

if st.button("Generate Report", type="primary"):
#--- DATA PROCESSING ---

#CHART FORMATING
	def format_title(slide, text, alignment, font_name, font_size, font_bold = False, font_italic = None, font_color = RGBColor(0, 0, 0),left=Pt(75), top=Pt(25), width=Pt(850), height=Pt(70)):
		title_shape = slide.shapes.add_textbox(left=left, top=top, width=width, height=height)
		title_text_frame = title_shape.text_frame
		title_text_frame.text = text
		title_text_frame.word_wrap = True
		for paragraph in title_text_frame.paragraphs:
			paragraph.alignment = alignment
			for run in paragraph.runs:
				run.font.name = font_name
				run.font.bold = font_bold
				run.font.italic = font_italic
				run.font.color.rgb = font_color
				run.font.size = Pt(font_size)
		return title_shape

	def pie_chart(slide,df,x,y,cx,cy,fontsize=9,legend_right = True, chart_title = False, title='',fontsize_title = Pt(20)):
		df.fillna(0, inplace = True) #fill nan
		# Convert the transposed DataFrame into chart data
		chart_data = CategoryChartData()
		# Add the brand names as categories to the chart data
		for i in df.transpose().columns:
			chart_data.add_category(i)
	  	# Add the SOV values as series to the chart data
		for index, row in df.transpose().iterrows():
			chart_data.add_series(index, row.values)
  		
		# create chart  
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
	
		if chart.has_legend:
			chart.legend.include_in_layout = False
			chart.legend.position = XL_LEGEND_POSITION.RIGHT if legend_right else XL_LEGEND_POSITION.BOTTOM
			chart.legend.font.size = Pt(fontsize)
	
		for series in chart.plots[0].series:
			for i, val in enumerate(series.values):
				if val == 0:
					series.points[i].data_label.has_text_frame = True
				series.data_labels.show_value = True
				series.data_labels.font.size = Pt(fontsize)
				series.data_labels.number_format = '0%'
				series.data_labels.position = XL_LABEL_POSITION.BEST_FIT
				series.data_labels.show_category_name = True
	
		return chart

	def line_marker_chart(slide,df,x,y,cx,cy, legend = True, legend_position = XL_LEGEND_POSITION.RIGHT,
                      data_show = False, chart_title = False, title ="", fontsize = Pt(12),
                      fontsize_title = Pt(14), percentage = False, line_width = Pt(1), smooth=False):
		df.fillna(0, inplace = True) #fill nan
	  # Define chart data
		chart_data = CategoryChartData()

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
	
		if percentage:
			category_axis.tick_labels.NumberFormat = '0"%"'
	
		return chart

	def table_default(slide, df, left, top, width, height, width_row, height_row, header=True, upper=False, fontsize=10, alignment=PP_ALIGN.LEFT):
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
				for paragraph in cell.text_frame.paragraphs:
					paragraph.alignment = alignment
					for run in paragraph.runs:
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

	def horizontal_bar_chart(slide, df, x, y, cx, cy, legend=True, legend_position=XL_LEGEND_POSITION.RIGHT,
							 data_show=True, chart_title=False, title="", fontsize=Pt(12),
							 fontsize_title=Pt(14), percentage=False, bar_width=Pt(10)):
		df.fillna(0, inplace=True)  # Fill NaN values
		# Define chart data
		chart_data = CategoryChartData()
		for i in df.index:
			chart_data.add_category(i)
		for col in df.columns:
			chart_data.add_series(col, np.where(df[col].values == 0, None, df[col].values))
	
	    # Create horizontal bar chart
		chart = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, x, y, cx, cy, chart_data).chart
	
	    # Add legend
		if legend:
			chart.has_legend = True
			chart.legend.include_in_layout = False
			chart.legend.position = legend_position
			chart.legend.font.size = fontsize
		else:
			chart.has_legend = False
	
	    # Show data labels
		if data_show:
			for series in chart.series:
				for point in series.points:
					point.data_label.show_value = True
					point.data_label.font.size = fontsize
					point.data_label.position = XL_LABEL_POSITION.OUTSIDE_END
	
	    # Customize category axis (y-axis) label font size
		chart.category_axis.tick_labels.font.size = fontsize
	
	    # Customize value axis (x-axis) label font size and format
		chart.value_axis.tick_labels.font.size = fontsize
		if percentage:
			chart.value_axis.tick_labels.number_format = '0%'
		elif df.max().max() >= 1000:
			chart.value_axis.tick_labels.number_format = '#,##0'  # Add commas for thousands
	
	    # Adjust bar width
		for series in chart.series:
			series.format.line.width = bar_width
	
		# Set chart title
		if chart_title:
			chart.chart_title.text_frame.text = title
			title_font = chart.chart_title.text_frame.paragraphs[0].font
			title_font.bold = True
			title_font.size = fontsize_title
		else:
			chart.has_title = False
	
	    # Remove Gridlines
		value_axis = chart.value_axis
		category_axis = chart.category_axis
		value_axis.has_major_gridlines = False
		value_axis.has_minor_gridlines = False
		category_axis.has_major_gridlines = False
		category_axis.has_minor_gridlines = False
	
		return chart

	def vertical_bar_chart(slide, df, x, y, cx, cy, legend=True, legend_position=XL_LEGEND_POSITION.RIGHT,
							 data_show=True, chart_title=False, title="", fontsize=Pt(12),
							 fontsize_title=Pt(14), percentage=False, bar_width=Pt(10)):
		df.fillna(0, inplace=True)  # Fill NaN values
		# Define chart data
		chart_data = CategoryChartData()
		for i in df.index:
			chart_data.add_category(i)
		for col in df.columns:
			chart_data.add_series(col, np.where(df[col].values == 0, None, df[col].values))
	
	    # Create vertical bar chart
		chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data).chart
	
	    # Add legend
		if legend:
			chart.has_legend = True
			chart.legend.include_in_layout = False
			chart.legend.position = legend_position
			chart.legend.font.size = fontsize
		else:
			chart.has_legend = False
	
	    # Show data labels
		if data_show:
			for series in chart.series:
				for point in series.points:
					point.data_label.show_value = True
					point.data_label.font.size = fontsize
					point.data_label.position = XL_LABEL_POSITION.OUTSIDE_END
	
	    # Customize category axis (y-axis) label font size
		chart.category_axis.tick_labels.font.size = fontsize
	
	    # Customize value axis (x-axis) label font size and format
		chart.value_axis.tick_labels.font.size = fontsize
		if percentage:
			chart.value_axis.tick_labels.number_format = '0%'
		elif df.max().max() >= 1000:
			chart.value_axis.tick_labels.number_format = '#,##0'  # Add commas for thousands
	
	    # Adjust bar width
		for series in chart.series:
			series.format.line.width = bar_width
	
		# Set chart title
		if chart_title:
			chart.chart_title.text_frame.text = title
			title_font = chart.chart_title.text_frame.paragraphs[0].font
			title_font.bold = True
			title_font.size = fontsize_title
		else:
			chart.has_title = False
	
	    # Remove Gridlines
		value_axis = chart.value_axis
		category_axis = chart.category_axis
		value_axis.has_major_gridlines = False
		value_axis.has_minor_gridlines = False
		category_axis.has_major_gridlines = False
		category_axis.has_minor_gridlines = False
	
		return chart

	def combo_chart(slide, df, x, y, cx, cy, legend=True, legend_position=XL_LEGEND_POSITION.TOP,
                data_show=False, chart_title=False, title="", fontsize=Pt(10), fontsize_title=Pt(12),
                line_width=Pt(1), smooth=False, font_name='Neue Haas Grotesk Text Pro'):

		df.fillna(0, inplace=True)  # Fill NaN values
	
	    # Split Data: Stacked Bar Chart (Categories) & Line Chart (Total)
		df_categories = df.iloc[:, :-1]  # Exclude 'Total' column for bars
		df_total = df.iloc[:, -1]        # Only 'Total' column for line chart
	
	    # Stacked Bar Chart Data
		chart_data = CategoryChartData()
		chart_data.categories = df.index  # Use index as categories (e.g., Months)	
		for category in df_categories.columns:
			chart_data.add_series(category, df_categories[category].values.tolist())
	
	    # Add Stacked Bar Chart
		chart_shape = slide.shapes.add_chart(
			XL_CHART_TYPE.COLUMN_STACKED, x, y, cx, cy, chart_data
	    )
		chart = chart_shape.chart
	
		chart.value_axis.tick_labels.font.size = fontsize
		chart.category_axis.tick_labels.font.size = fontsize
		chart.value_axis.tick_labels.number_format = "#,##0"  # Thousands separator
		chart.value_axis.has_major_gridlines = False
	
	    # Add percentage labels inside bars
		df_percent = df_categories.div(df_total, axis=0) * 100
		for series, category in zip(chart.series, df_percent.columns):
			series.data_labels.show_category_name = False
			series.data_labels.show_value = True
			series.data_labels.font.size = fontsize
			for point, percent in zip(series.points, df_percent[category]):
				point.has_data_label = False
				point.data_label.font.size = fontsize
				point.data_label.text_frame.text = f"{percent:.0f}%"  # Format as percentage
	
	    # Add legend
		if legend:
			chart.has_legend = True
			chart.legend.include_in_layout = False
			chart.legend.position = legend_position
			chart.legend.font.size = fontsize
	
	    # Line Chart Data (Total)
		line_chart_data = CategoryChartData()
		line_chart_data.categories = df.index
		line_chart_data.add_series("Total", df_total.values.tolist())
	
	    # Add Line Chart Overlay
		line_chart_shape = slide.shapes.add_chart(
			XL_CHART_TYPE.LINE, x, y, cx, cy, line_chart_data
	    )
		line_chart = line_chart_shape.chart
	
	    # Remove X-axis and Gridlines from Line Chart
		line_chart.category_axis.visible = False # Remove X-axis
		line_chart.value_axis.visible = False # Remove Y-axis
		line_chart.value_axis.has_major_gridlines = False # Remove gridlines
		line_chart.category_axis.has_major_gridlines = False # Remove gridlines
		line_chart.has_legend = False  # Remove legend
		line_chart.has_title = False  # Remove chart title
		line_chart.value_axis.tick_labels.number_format = '""'  # Hide Y-axis labels
	
	
	    # Add Data Labels to Line Chart
		for series in line_chart.series:
			series.has_data_labels = True
			for point in series.points:
				point.has_data_label = True
				point.data_label.number_format = "#,##0"  # Thousands separator
	
	    # Apply Smooth Line if required
		if smooth:
			for series in line_chart.series:
				series.smooth = True
	
	    # Set Line Chart Color and Style
		line_series = line_chart.series[0]
		line_series.marker.style = XL_MARKER_STYLE.CIRCLE
		line_series.marker.format.fill.solid()
		line_series.marker.format.fill.fore_color.rgb = RGBColor(78, 167, 46)
		line_series.format.line.width = line_width
		line_series.format.line.fill.solid()
		line_series.format.line.fill.fore_color.rgb = RGBColor(78, 167, 46)  # Green lime
	
	    # Add Chart Title
		if chart_title:
			chart.chart_title.text_frame.text = title
			chart.chart_title.text_frame.paragraphs[0].font.size = fontsize_title
			chart.chart_title.text_frame.paragraphs[0].font.bold = True
		else:
			chart.has_title = False
		
		return chart

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
   
# df for Competitive data (TDK)
	# Query Code
	query = """
 	SELECT 
	date
	,month
	,years
	,brand
	,category
	,division
	,manufacturer
	,advertiser_name
	,SUM(views_float) as views
	,SUM(engagements) as engagements
	,SUM(content) as content
	FROM loreal-id-prod.loreal_storage.advocacy_tdk_df
	WHERE years = {} or years = {}
	GROUP BY 
	date
	,month
	,years
	,brand
	,division
	,category
	,manufacturer
	,advertiser_name
	""".format(year_map, (year_map-1))

	# Fetch data
	df = client.query(query).to_dataframe()

	# Add Date & Quarter column
	df['date'] = pd.to_datetime(df['date'], format='%y-%m-%d')

	df['quarter'] = (df['date'].dt.quarter).map({1: 'Q1', 2: 'Q2', 3: 'Q3', 4: 'Q4'})

# df2 for INTERNAL DATA
	# Query Code 2
	query2 = """
 	SELECT 
	date_post as date
 	,grid_month as month
  	,grid_year as years
 	,brand
 	,division
  	,main_category as category
  	,tier
	,SUM(actual_views) as views
	,SUM(engagement) as engagements
	,COUNT(brand) as content
	FROM loreal-id-prod.loreal_storage.advocacy_campaign_df
	WHERE grid_year = {}
	GROUP BY 
	date_post
  	,grid_month
  	,grid_year
 	,brand
 	,division
  	,main_category
  	,tier
	""".format(year)

	# Fetch data
	df2 = client.query(query2).to_dataframe()

	# Add Date & Quarter column
	df2['date'] = pd.to_datetime(df2['date'], format='%y-%m-%d')

	df2['quarter'] = (df2['date'].dt.quarter).map({1: 'Q1', 2: 'Q2', 3: 'Q3', 4: 'Q4'})
	st.write(df2.head())

	df2['years'] = df2['years'].astype(str)
	
#---- SLIDES PRESENTATION ----
#	ppt_temp_loc = "GOAT_LOreal_DeckToGo/pages/Template Deck to Go - Loreal Indonesia.pptx"
  
	ppt = Presentation(uploaded_file)

#------------PAGE 1--------------
	page_no = 0 #PAGE1

## Title Slide
# Add a title to the slide
	format_title(ppt.slides[page_no], "MONTHLY BRAND & DIVISION REPORT", alignment=PP_ALIGN.LEFT, font_name= 'Aptos Display', font_size=50, font_bold=True,left=Pt(35), top=Pt(150), width=Pt(500), height=Pt(70), font_color=RGBColor(255, 255, 255))

#------------PAGE 2--------------
	page_no = page_no + 1 #PAGE2

	# Add a title to the slide
	format_title(ppt.slides[page_no], "MONTHLY SOV & SOE", alignment=PP_ALIGN.LEFT, font_name= 'Neue Haas Grotesk Text Pro', font_size=28, font_bold=True,left=Inches(0.5), top=Inches(0.5), width=Inches(12), height=Inches(0.3), font_color=RGBColor(0, 0, 0))
	format_title(ppt.slides[page_no], "Total Views", alignment=PP_ALIGN.CENTER, font_name= 'Neue Haas Grotesk Text Pro', font_size=18, font_italic=True,left=Inches(5), top=Inches(2), width=Inches(1.43), height=Inches(1.01), font_color=RGBColor(255, 255, 255))
	format_title(ppt.slides[page_no], "Total Eng.", alignment=PP_ALIGN.CENTER, font_name= 'Neue Haas Grotesk Text Pro', font_size=18, font_italic=True,left=Inches(11), top=Inches(2), width=Inches(1.43), height=Inches(1.01), font_color=RGBColor(255, 255, 255))

	# Filter the dataframe
	df_m = df[(df['division'].isin(division)) & (df['category'].isin(category)) & (df['years'] == year_map) &  (df['month'] == month)]

	# Perform groupby and aggregation with handling for datetime64 columns
	grouped_df_m = df_m.groupby('brand').agg({
    	# Numerical columns: sum them
    	'views': 'sum',
    	'engagements': 'sum',
    	'content': 'sum'
	})
	grouped_df_m = grouped_df_m.reset_index()

	# Calculate Total Views
	total_views_m = grouped_df_m['views'].sum().astype(int)
	total_engagement_m = grouped_df_m['engagements'].sum().astype(int)

	# Calculate SOV (%)
	grouped_df_m['SOV%'] = (grouped_df_m['views'] / total_views_m)
	grouped_df_m['SOE%'] = (grouped_df_m['engagements'] / total_engagement_m)

	sov_df_m = grouped_df_m[['brand', 'SOV%']].sort_values(by='SOV%', ascending=False)
	soe_df_m = grouped_df_m[['brand', 'SOE%']].sort_values(by='SOE%', ascending=False)

	# Sort by 'SOV%' and keep the top 10 brands
	top_brands = sov_df_m.head(10)

	# Group all other brands into "Others"
	others_sov = 1 - top_brands["SOV%"].sum()  # Remaining percentage

	# Append "Others" row if applicable
	top_brands_m = top_brands.copy()
	if others_sov > 0:
		others_row = pd.DataFrame([{"brand": "Others", "SOV%": others_sov}])
		top_brands_m = pd.concat([top_brands, others_row], ignore_index=True)

	# Add pie chart
	### st.write(top_brands_m.set_index('brand'))
	pie_chart(ppt.slides[page_no], top_brands_m.set_index('brand'), Inches(0.5), Inches(1.5), Inches(6), Inches(6), chart_title=True, title='SOV', fontsize_title = Pt(20), fontsize=9)
	pie_chart(ppt.slides[page_no], top_brands_m.set_index('brand'), Inches(7), Inches(1.5), Inches(6), Inches(6), chart_title=True, title='SOE', fontsize_title = Pt(20), fontsize=9)

	format_title(ppt.slides[page_no], "Total Views", alignment=PP_ALIGN.CENTER, font_name= 'Neue Haas Grotesk Text Pro', font_size=18, font_italic=True,left=Inches(5.3), top=Inches(3), width=Inches(1.3), height=Inches(1.01), font_color=RGBColor(0, 0, 0))
	format_title(ppt.slides[page_no], format(total_views_m, ","), alignment=PP_ALIGN.CENTER, font_name= 'Neue Haas Grotesk Text Pro', font_size=18, font_bold=True,left=Inches(5.3), top=Inches(2.7), width=Inches(1.3), height=Inches(0.5), font_color=RGBColor(0, 0, 0))

	format_title(ppt.slides[page_no], "Total Eng.", alignment=PP_ALIGN.CENTER, font_name= 'Neue Haas Grotesk Text Pro', font_size=18, font_italic=True,left=Inches(11.7), top=Inches(3), width=Inches(1.3), height=Inches(1.01), font_color=RGBColor(0, 0, 0))
	format_title(ppt.slides[page_no], format(total_engagement_m, ","), alignment=PP_ALIGN.CENTER, font_name= 'Neue Haas Grotesk Text Pro', font_size=18, font_bold=True,left=Inches(11.7), top=Inches(2.7), width=Inches(1.3), height=Inches(0.5), font_color=RGBColor(0, 0, 0))

#------------PAGE 3--------------
	page_no = page_no + 1 #PAGE3

# Add a title to the slide
	format_title(ppt.slides[page_no], "QUARTERLY SOV & SOE", alignment=PP_ALIGN.LEFT, font_name= 'Neue Haas Grotesk Text Pro', font_size=28, font_bold=True,left=Inches(0.5), top=Inches(0.5), width=Inches(12), height=Inches(0.3), font_color=RGBColor(0, 0, 0))

# Filter the dataframe
	df_q = df[(df['division'].isin(division)) & (df['category'].isin(category)) & (df['years'] == year_map) & (df['quarter'] == quarter)]

# Perform groupby and aggregation with handling for datetime64 columns
	grouped_df_q = df_q.groupby('brand').agg({
    	# Numerical columns: sum them
    	'views': 'sum',
    	'engagements': 'sum',
    	'content': 'sum'
	})
	grouped_df_q = grouped_df_q.reset_index()

# Calculate Total Views
	total_views_q = (grouped_df_q['views'].sum()).astype(int)
	total_engagement_q = (grouped_df_q['engagements'].sum()).astype(int)

	# Calculate SOV (%)
	grouped_df_q['SOV%'] = (grouped_df_q['views'] / total_views_q)
	grouped_df_q['SOE%'] = (grouped_df_q['engagements'] / total_engagement_q)

	sov_df_q = grouped_df_q[['brand', 'SOV%']].sort_values(by='SOV%', ascending=False)
	soe_df_q = grouped_df_q[['brand', 'SOE%']].sort_values(by='SOE%', ascending=False)
	
	# Sort by 'SOV%' and keep the top 10 brands
	top_brands = sov_df_q.head(10)

	# Group all other brands into "Others"
	others_sov = 1 - top_brands["SOV%"].sum()  # Remaining percentage

	# Append "Others" row if applicable
	top_brands_q= top_brands.copy()
	if others_sov > 0:
		others_row = pd.DataFrame([{"brand": "Others", "SOV%": others_sov}])
		top_brands_q = pd.concat([top_brands, others_row], ignore_index=True)
		
	# Add pie chart
	### st.write(top_brands_q.set_index('brand'))
	pie_chart(ppt.slides[page_no], top_brands_q.set_index('brand'), Inches(0.5), Inches(1.5), Inches(6), Inches(6), chart_title=True, title='SOV', fontsize_title = Pt(20), fontsize=9)
	pie_chart(ppt.slides[page_no], top_brands_q.set_index('brand'), Inches(7), Inches(1.5), Inches(6), Inches(6), chart_title=True, title='SOE', fontsize_title = Pt(20), fontsize=9)

	format_title(ppt.slides[page_no], "Total Views", alignment=PP_ALIGN.CENTER, font_name= 'Neue Haas Grotesk Text Pro', font_size=18, font_italic=True,left=Inches(5.3), top=Inches(3), width=Inches(1.3), height=Inches(1.01), font_color=RGBColor(0, 0, 0))
	format_title(ppt.slides[page_no], format(total_views_q, ","), alignment=PP_ALIGN.CENTER, font_name= 'Neue Haas Grotesk Text Pro', font_size=18, font_bold=True,left=Inches(5.3), top=Inches(2.7), width=Inches(1.3), height=Inches(0.5), font_color=RGBColor(0, 0, 0))

	format_title(ppt.slides[page_no], "Total Eng.", alignment=PP_ALIGN.CENTER, font_name= 'Neue Haas Grotesk Text Pro', font_size=18, font_italic=True,left=Inches(11.7), top=Inches(3), width=Inches(1.3), height=Inches(1.01), font_color=RGBColor(0, 0, 0))
	format_title(ppt.slides[page_no], format(total_engagement_q, ","), alignment=PP_ALIGN.CENTER, font_name= 'Neue Haas Grotesk Text Pro', font_size=18, font_bold=True,left=Inches(11.7), top=Inches(2.7), width=Inches(1.3), height=Inches(0.5), font_color=RGBColor(0, 0, 0))

#------------PAGE 4--------------
	page_no = page_no + 1 #PAGE4

# Add a title to the slide
	format_title(ppt.slides[page_no], "ANNUAL SOV & SOE", alignment=PP_ALIGN.LEFT, font_name= 'Neue Haas Grotesk Text Pro', font_size=28, font_bold=True,left=Inches(0.5), top=Inches(0.5), width=Inches(12), height=Inches(0.3), font_color=RGBColor(0, 0, 0))

# Filter the dataframe
	df_y = df[(df['division'].isin(division)) & (df['category'].isin(category)) & (df['years'] == year_map)]

# Perform groupby and aggregation with handling for datetime64 columns
	grouped_df_y = df_y.groupby('brand').agg({
		# Numerical columns: sum them
		'views': 'sum',
		'engagements': 'sum',
		'content': 'sum'
	})
	grouped_df_y = grouped_df_y.reset_index()

# Calculate Total Views
	total_views_y = (grouped_df_y['views'].sum()).astype(int)
	total_engagement_y = (grouped_df_y['engagements'].sum()).astype(int)

# Calculate SOV (%)
	grouped_df_y['SOV%'] = (grouped_df_y['views'] / total_views_y)
	grouped_df_y['SOE%'] = (grouped_df_y['engagements'] / total_engagement_y)

	sov_df_y = grouped_df_y[['brand', 'SOV%']].sort_values(by='SOV%', ascending=False)
	soe_df_y = grouped_df_y[['brand', 'SOE%']].sort_values(by='SOE%', ascending=False)
	
	# Sort by 'SOV%' and keep the top 10 brands
	top_brands = sov_df_y.head(10)

	# Group all other brands into "Others"
	others_sov = 1 - top_brands["SOV%"].sum()  # Remaining percentage

	# Append "Others" row if applicable
	top_brands_y= top_brands.copy()
	if others_sov > 0:
		others_row = pd.DataFrame([{"brand": "Others", "SOV%": others_sov}])
		top_brands_y = pd.concat([top_brands, others_row], ignore_index=True)

# Add pie chart
	### st.write(top_brands_y.set_index('brand'))
	pie_chart(ppt.slides[page_no], top_brands_y.set_index('brand'), Inches(0.5), Inches(1.5), Inches(6), Inches(6), chart_title=True, title='SOV', fontsize_title = Pt(20), fontsize=9)
	pie_chart(ppt.slides[page_no], top_brands_y.set_index('brand'), Inches(7), Inches(1.5), Inches(6), Inches(6), chart_title=True, title='SOE', fontsize_title = Pt(20), fontsize=9)

	format_title(ppt.slides[page_no], "Total Views", alignment=PP_ALIGN.CENTER, font_name= 'Neue Haas Grotesk Text Pro', font_size=18, font_italic=True,left=Inches(5.3), top=Inches(3), width=Inches(1.3), height=Inches(1.01), font_color=RGBColor(0, 0, 0))
	format_title(ppt.slides[page_no], format(total_views_y, ","), alignment=PP_ALIGN.CENTER, font_name= 'Neue Haas Grotesk Text Pro', font_size=18, font_bold=True,left=Inches(5.3), top=Inches(2.7), width=Inches(1.3), height=Inches(0.5), font_color=RGBColor(0, 0, 0))

	format_title(ppt.slides[page_no], "Total Eng.", alignment=PP_ALIGN.CENTER, font_name= 'Neue Haas Grotesk Text Pro', font_size=18, font_italic=True,left=Inches(11.7), top=Inches(3), width=Inches(1.3), height=Inches(1.01), font_color=RGBColor(0, 0, 0))
	format_title(ppt.slides[page_no], format(total_engagement_y, ","), alignment=PP_ALIGN.CENTER, font_name= 'Neue Haas Grotesk Text Pro', font_size=18, font_bold=True,left=Inches(11.7), top=Inches(2.7), width=Inches(1.3), height=Inches(0.5), font_color=RGBColor(0, 0, 0))

#------------PAGE 5--------------
	page_no = page_no + 1 #PAGE5

# Add a title to the slide
	format_title(ppt.slides[page_no], "YTD BEAUTY MARKET VIEWS AND ENGAGEMENT SUMMARY", alignment=PP_ALIGN.LEFT, font_name= 'Neue Haas Grotesk Text Pro', font_size=28, font_bold=True,left=Inches(0.5), top=Inches(0.5), width=Inches(12.3), height=Inches(0.3), font_color=RGBColor(0, 0, 0))

# Aggregate Views by Month and Manufacturer
	stacked_data_views = df_y.pivot_table(index="month", columns="manufacturer", values="views", aggfunc="sum", fill_value=0)
	stacked_data_views['Total'] = stacked_data_views.sum(axis=1)

# Aggregate Views by Month and Manufacturer
	stacked_data_eng = df_y.pivot_table(index="month", columns="manufacturer", values="engagements", aggfunc="sum", fill_value=0)
	stacked_data_eng['Total'] = stacked_data_eng.sum(axis=1)

# Sort months correctly
	month_order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
	stacked_data_eng = stacked_data_eng.reindex(month_order)
	stacked_data_views = stacked_data_views.reindex(month_order)

# Add combo stacked bar chart
	combo_chart(ppt.slides[page_no], stacked_data_views, Inches(.1), Inches(1.5), Inches(9), Inches(2.8), chart_title=True, title="Market Movement - Views",
            fontsize=Pt(10), fontsize_title=Pt(12), smooth=True, data_show=True)
	combo_chart(ppt.slides[page_no], stacked_data_eng, Inches(.1), Inches(4.5), Inches(9), Inches(2.8), chart_title=True, title="Market Movement - Engagements",
            fontsize=Pt(10), fontsize_title=Pt(12), smooth=True)
	
#------------PAGE 6--------------
	page_no = page_no + 1 #PAGE6

	format_title(ppt.slides[page_no], "VIEWS PERFORMANCE: YTD and YTG with SOV PROJECTION", alignment=PP_ALIGN.LEFT, font_name= 'Neue Haas Grotesk Text Pro', font_size=28, font_bold=True,left=Inches(0.5), top=Inches(0.5), width=Inches(12.3), height=Inches(0.3), font_color=RGBColor(0, 0, 0))

# Filter the dataframe
	df_y2 = df[(df['division'].isin(division)) & (df['years'] == year_map)]

# Calculate Total Views per Category
	total_views = df_y2.groupby('category')['views'].sum().reset_index()
	total_views.rename(columns={'views': 'Total_Views'}, inplace=True)

# Aggregate Views and calculate SOV
	df_grouped = df_y2.groupby(['brand', 'category', 'advertiser_name'])['views'].sum().reset_index()
	df_grouped = df_grouped.merge(total_views, on='category', how='left')
	df_grouped['SOV'] = (df_grouped['views'] / df_grouped['Total_Views']).map('{:.0%}'.format)
	df_grouped['views'] = df_grouped['views'].astype(int)

# Rank Brands within each Category
	df_grouped['#Rank'] = df_grouped.groupby('category')['views'].rank(method='dense', ascending=False).astype(int)

# Filter for L'Oreal Advertiser
	df_final = df_grouped[df_grouped['advertiser_name'] == "L'Oreal"][['category', 'brand', 'views', 'SOV', '#Rank']]
	df_final = df_final.sort_values(['category', 'brand'])

	table_default(ppt.slides[page_no], df_final, Inches(1), Inches(1.2), Inches(12.2), Inches(5.2),
				  [Inches(1.5)]*2+[Inches(0.75)]*3+[Inches(1),Inches(0.5)], Inches(0.5), header=True, upper=True, fontsize=12, alignment=PP_ALIGN.LEFT)

#------------PAGE 7--------------
	page_no = page_no + 1 #PAGE7

	format_title(ppt.slides[page_no], "MONTHLY TIMELINE", alignment=PP_ALIGN.LEFT, font_name= 'Neue Haas Grotesk Text Pro', font_size=28, font_bold=True,left=Inches(0.5), top=Inches(0.5), width=Inches(12.3), height=Inches(0.3), font_color=RGBColor(0, 0, 0))

# Aggregate total views per brand and rank them
	brand_ranking = df_y.groupby('brand')['views'].sum().reset_index()
	brand_ranking['Brand_Rank'] = brand_ranking['views'].rank(method='dense', ascending=False)

# Keep only the top 7 brands
	top_brands = brand_ranking[brand_ranking['Brand_Rank'] <= 7]['brand']

# Retrieve daily views for these top 7 brands
	df_top_brands = df_y[df_y['brand'].isin(top_brands)]
	df_m_views = pd.pivot_table(df_top_brands[['date','brand','views']], columns = 'date', index = 'brand', aggfunc = 'sum', fill_value = 0)
	df_m_views.columns = df_m_views.columns.droplevel() # drop column level
	df_m_views.columns = df_m_views.columns.strftime('%b')

# Reshape the DataFrame before grouping to have 'views' as a column
	df_m_views2 = df_m_views.reset_index().melt(id_vars=['brand'], var_name='month', value_name='views')

	df_views = df_m_views2.groupby(['brand'], as_index=False)['views'].sum()
	df_views['SOV%'] = (df_views['views'] / total_views_y).map('{:.0%}'.format)
	df_views['views'] = df_views['views'].astype(int)
	df_views['rank'] = df_views['views'].rank(method='dense', ascending=False).astype(int)
	df_views = df_views.sort_values('rank')

# Add line chart
	line_marker_chart(ppt.slides[page_no], df_m_views, Inches(0.5), Inches(1.7), Inches(7), Inches(5), legend=True, legend_position=XL_LEGEND_POSITION.BOTTOM, chart_title = False, fontsize_title = Pt(12), line_width = Pt(2), smooth=True)
# Add table
	table_default(ppt.slides[page_no], df_views, Inches(8), Inches(1.7), Inches(5), Inches(5),
								[Inches(1.5)]+[Inches(0.75)]*3+[Inches(1),Inches(0.5)], Inches(0.5), header=True, upper=True, fontsize=12, alignment=PP_ALIGN.LEFT)

#------------PAGE 8--------------
	page_no = page_no + 1 
	format_title(ppt.slides[page_no], "QUARTERLY TIMELINE", alignment=PP_ALIGN.LEFT, font_name= 'Neue Haas Grotesk Text Pro', font_size=28, font_bold=True,left=Inches(0.5), top=Inches(0.5), width=Inches(12.3), height=Inches(0.3), font_color=RGBColor(0, 0, 0))

# Aggregate total views per brand and rank them
	df_q_views = pd.pivot_table(df_top_brands[['quarter','brand','views']], columns = 'quarter', index = 'brand', aggfunc = 'sum', fill_value = 0)
	df_q_views.columns = df_q_views.columns.droplevel() # drop column level

# Add line chart
	line_marker_chart(ppt.slides[page_no], df_q_views, Inches(0.5), Inches(1.7), Inches(7), Inches(5), legend=True, legend_position=XL_LEGEND_POSITION.BOTTOM, chart_title = False, fontsize_title = Pt(12), line_width = Pt(2), smooth=True)
# Add table
	table_default(ppt.slides[page_no], df_views, Inches(8), Inches(1.7), Inches(5), Inches(5),
								[Inches(1.5)]+[Inches(0.75)]*3+[Inches(1),Inches(0.5)], Inches(0.5), header=True, upper=True, fontsize=12, alignment=PP_ALIGN.LEFT)

#------------PAGE 9--------------
	page_no = page_no + 1 
	format_title(ppt.slides[page_no], "BAR CHART COMPARISON QTR VERSUS", alignment=PP_ALIGN.LEFT, font_name= 'Neue Haas Grotesk Text Pro', font_size=28, font_bold=True,left=Inches(0.5), top=Inches(0.5), width=Inches(12.3), height=Inches(0.3), font_color=RGBColor(0, 0, 0))

	quarter_mapping = {'Q1': ['Q4', 'Q1'], 'Q2': ['Q1', 'Q2'], 'Q3': ['Q2', 'Q3'], 'Q4': ['Q3', 'Q4']}
	quarter_compare = quarter_mapping.get(quarter, [quarter])
	title_q = quarter_compare[0] + " vs " + quarter_compare[1]

	df_q_views = pd.pivot_table(df_top_brands[['quarter','brand','views']], columns = 'quarter', index = 'brand', aggfunc = 'sum', fill_value = 0)
	df_q_content = pd.pivot_table(df_top_brands[['quarter','brand','content']], columns = 'quarter', index = 'brand', aggfunc = 'sum', fill_value = 0)

	df_q_views = df_q_views[[('views', q) for q in quarter_compare]]
	df_q_views = np.ceil(df_q_views * 10) / 10 # Round up with 1 decimal place
	df_q_views.columns = quarter_compare

	df_q_content = df_q_content[[('content', q) for q in quarter_compare]]
	df_q_content.columns = quarter_compare

# Add horizontal bar chart views
	horizontal_bar_chart(ppt.slides[page_no], df_q_views, Inches(0.5), Inches(1.9), Inches(5.5), Inches(5),
                     chart_title = True, title= f"{title_q} Views Contribution", fontsize_title = Pt(16),
                     legend=True, legend_position=XL_LEGEND_POSITION.TOP,
                     bar_width = Pt(8), percentage=False, fontsize=Pt(10))

# Add horizontal bar chart content
	horizontal_bar_chart(ppt.slides[page_no], df_q_content, Inches(7), Inches(1.9), Inches(5.5), Inches(5),
                     chart_title = True, title= f"{title_q} Content Contribution", fontsize_title = Pt(16),
                     legend=True, legend_position=XL_LEGEND_POSITION.TOP,
                     bar_width = Pt(8), percentage=False, fontsize=Pt(10))

#------------PAGE 10--------------
	page_no = page_no + 1
	format_title(ppt.slides[page_no], "MONTHLY PER BRAND CONTRIBUTION", alignment=PP_ALIGN.LEFT, font_name= 'Neue Haas Grotesk Text Pro', font_size=28, font_bold=True,left=Inches(0.5), top=Inches(0.5), width=Inches(12.3), height=Inches(0.3), font_color=RGBColor(0, 0, 0))

# Filter the dataframe
	df_10 = df2[(df2['brand'].isin(brands)) & (df2['years'] == year) &  (df2['month'] == month)]
	df_10_views = pd.pivot_table(df_10[['brand', 'views']], index = 'brand', aggfunc = 'sum', fill_value = 0)
	df_10_eng = pd.pivot_table(df_10[['brand', 'engagements']], index = 'brand', aggfunc = 'sum', fill_value = 0)
	df_10_content = pd.pivot_table(df_10[['brand', 'content']], index = 'brand', aggfunc = 'sum', fill_value = 0)
	st.write("df_10_views :", df_10_views)

# Add vertical bar chart
	vertical_bar_chart(ppt.slides[page_no], df_10_views, Inches(0.9), Inches(1.9), Inches(3.7), Inches(3),
                     chart_title = True, title= "View Performance", fontsize_title = Pt(16),
                     legend=True, legend_position=XL_LEGEND_POSITION.TOP,
                     bar_width = Pt(8), percentage=False, fontsize=Pt(10))
	vertical_bar_chart(ppt.slides[page_no], df_10_eng, Inches(4.9), Inches(1.9), Inches(3.7), Inches(3),
                     chart_title = True, title= "Engagement Performance", fontsize_title = Pt(16),
                     legend=True, legend_position=XL_LEGEND_POSITION.TOP,
                     bar_width = Pt(8), percentage=False, fontsize=Pt(10))
	vertical_bar_chart(ppt.slides[page_no], df_10_content, Inches(8.9), Inches(1.9), Inches(3.7), Inches(3),
                     chart_title = True, title= "Content Performance", fontsize_title = Pt(16),
                     legend=True, legend_position=XL_LEGEND_POSITION.TOP,
                     bar_width = Pt(8), percentage=False, fontsize=Pt(10))
# Add rectangle
	ppt.slides[page_no].shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.9), Inches(5), Inches(3.7), Inches(2))
	ppt.slides[page_no].shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.9), Inches(5), Inches(3.7), Inches(2))
	ppt.slides[page_no].shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.9), Inches(5), Inches(3.7), Inches(2))
	
#------------PAGE 11--------------
	page_no = page_no + 1 
	format_title(ppt.slides[page_no], "QUARTERLY PER BRAND CONTRIBUTION", alignment=PP_ALIGN.LEFT, font_name= 'Neue Haas Grotesk Text Pro', font_size=28, font_bold=True,left=Inches(0.5), top=Inches(0.5), width=Inches(12.3), height=Inches(0.3), font_color=RGBColor(0, 0, 0))

	df_11 = df2[(df2['brand'].isin(brands)) & (df2['years'] == year) &  (df2['quarter'] == quarter)]
	df_11_views = pd.pivot_table(df_11[['brand', 'views']], index = 'brand', aggfunc = 'sum', fill_value = 0)
	df_11_eng = pd.pivot_table(df_11[['brand', 'engagements']], index = 'brand', aggfunc = 'sum', fill_value = 0)
	df_11_content = pd.pivot_table(df_11[['brand', 'content']], index = 'brand', aggfunc = 'sum', fill_value = 0)
	st.write("df_11_views :", df_11_views)

# Add vertical bar chart
	vertical_bar_chart(ppt.slides[page_no], df_11_views, Inches(0.9), Inches(1.9), Inches(3.7), Inches(3),
                     chart_title = True, title= "View Performance", fontsize_title = Pt(16),
                     legend=True, legend_position=XL_LEGEND_POSITION.TOP,
                     bar_width = Pt(8), percentage=False, fontsize=Pt(10))
	vertical_bar_chart(ppt.slides[page_no], df_11_eng, Inches(4.9), Inches(1.9), Inches(3.7), Inches(3),
                     chart_title = True, title= "Engagement Performance", fontsize_title = Pt(16),
                     legend=True, legend_position=XL_LEGEND_POSITION.TOP,
                     bar_width = Pt(8), percentage=False, fontsize=Pt(10))
	vertical_bar_chart(ppt.slides[page_no], df_11_content, Inches(8.9), Inches(1.9), Inches(3.7), Inches(3),
                     chart_title = True, title= "Content Performance", fontsize_title = Pt(16),
                     legend=True, legend_position=XL_LEGEND_POSITION.TOP,
                     bar_width = Pt(8), percentage=False, fontsize=Pt(10))
	
# Add rectangle
	shapes1 = ppt.slides[page_no].shapes
	rectangle1 = shapes1.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(5), Inches(3.7), Inches(2))
	rectangle1.adjustments[0] = 0.5
	fill1 = rectangle1.fill.background()
	text_frame1 = rectangle1.text_frame
	text_frame1.text = "1. ..."
	text_frame1.paragraphs[0].font.size = Pt(13)

	shapes2 = ppt.slides[page_no].shapes
	rectangle2 = shapes2.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.9), Inches(5), Inches(3.7), Inches(2))
	rectangle2.adjustments[0] = 0.5
	fill2 = rectangle2.fill.background()
	text_frame2 = rectangle2.text_frame
	text_frame2.text = "1. ..."
	text_frame2.paragraphs[0].font.size = Pt(13)

	shapes3 = ppt.slides[page_no].shapes
	rectangle3 = shapes3.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.9), Inches(5), Inches(3.7), Inches(2))
	rectangle3.adjustments[0] = 0.5
	fill3 = rectangle3.fill.background()
	text_frame3 = rectangle3.text_frame
	text_frame3.text = "1. ..."
	text_frame3.paragraphs[0].font.size = Pt(13)

	
	
#------------PAGE 12--------------
	page_no = page_no + 1
	format_title(ppt.slides[page_no], "BRAND SCORE-CARD", alignment=PP_ALIGN.LEFT, font_name= 'Neue Haas Grotesk Text Pro', font_size=28, font_bold=True,left=Inches(0.5), top=Inches(0.5), width=Inches(12.3), height=Inches(0.3), font_color=RGBColor(0, 0, 0))

#------------PAGE 13--------------
	page_no = page_no + 1

#------------PAGE 14--------------
	page_no = page_no + 1
	if 'Select All' in category_selection:
		category_title = "ALL CATEGORY"
	else: category_title = "| ".join(category_selection)
	
	format_title(ppt.slides[page_no], category_title.upper()+" KOL MIX", alignment=PP_ALIGN.LEFT, font_name= 'Neue Haas Grotesk Text Pro', font_size=28, font_bold=True,left=Inches(0.5), top=Inches(0.5), width=Inches(12.3), height=Inches(0.3), font_color=RGBColor(0, 0, 0))

	# Filter the dataframe
	df_14 = df2[(df2['category'].isin(category)) & (df2['years'] == year) &  (df2['month'] == month)]
	df_14_views = pd.pivot_table(df_14[['tier', 'views']], index = 'tier', aggfunc = 'sum', fill_value = 0)
	df_14_views = df_14_views.sort_values(by=['tier'], ascending=False)
	st.write("df_14_views :", df_14_views)

# Add vertical bar chart
	vertical_bar_chart(ppt.slides[page_no], df_14_views, Inches(1), Inches(1.7), Inches(11), Inches(5),
                     chart_title = True, title= f"{category_title}", fontsize_title = Pt(16),
                     legend=True, legend_position=XL_LEGEND_POSITION.TOP,
                     bar_width = Pt(8), percentage=False, fontsize=Pt(10))
	
#-----------PAGE 15---------------
	page_no = page_no + 1

#-----------PAGE 16---------------
	page_no = page_no + 1	
	

	page_no = page_no + 1
	for i in range(page_no,len(ppt.slides)):
		try:
			xml_slides = ppt.slides._sldIdLst
			slides = list(xml_slides)
			xml_slides.remove(slides[page_no])
		except:
			pass

####### ------------- SAVE PPT -----------------
	file = (f'{category_title.upper()} MONTHLY REPORT - {month} {year}')
	filename = (f'{category_title.upper()} MONTHLY REPORT - {month} {year}.pptx')
	files = ppt.save(filename)

# ✅ This code runs **only after clicking the submit button**
	st.success(f"✅ Your report is ready : '{category_title} Monthly Report - {month} {year}.pptx'")

	# Store filename in session state for later use
	st.session_state["report_filename"] = filename

# Send Email Button (Only appears after submitting)
if "report_filename" in st.session_state:
	filename = st.session_state["report_filename"]
else:
	st.error("can't send email 📧 : No file generated yet. Please generate the report first.")
	st.stop()  # Stop execution if filename is missing
		
if st.button("Send Report", type= "secondary"):
	wib = pytz.timezone("Asia/Jakarta")
	now = datetime.now(wib)
	formatted_date = now.strftime("%Y-%m-%d %H:%M")
	
    # Gmail SMTP Configuration
	SMTP_SERVER = "smtp.gmail.com"
	SMTP_PORT = 587
	EMAIL_USER = "dyah.dinasari.groupm@gmail.com"
	EMAIL_PASS = "koxp pzgm ixws ihek"
    # Email Details
	send_to = email_list #["dyah.dinasari@groupm.com"]		
	subject = "[Test] m-Slide Goat-L'Oreal"
	body = """Hi team,
        
		This is a test email sent via Python SMTP.
        
	Regards,
	Dyah Dinasari"""

	# ✅ Create Email Message
	msg = MIMEMultipart()
	msg["From"] = EMAIL_USER
	msg["To"] = ", ".join(send_to)
	msg["Subject"] = subject
	msg.attach(MIMEText(body, "plain"))
    # ✅ Attach the PowerPoint File
	if os.path.exists(filename):  # Check if the file exists
		with open(filename, "rb") as file:
			part = MIMEApplication(file.read(), Name=os.path.basename(filename))
			part["Content-Disposition"] = f'attachment; filename="{os.path.basename(filename)}"'
			msg.attach(part)
	else:
		st.write(f"⚠ Warning: File '{filename}' not found. Email will be sent without attachment.")

    # ✅ Send Email via Gmail SMTP
	try:
		smtp = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
		smtp.starttls()  # Secure the connection
		smtp.login(EMAIL_USER, EMAIL_PASS)  # Login with App Password
		smtp.sendmail(EMAIL_USER, send_to, msg.as_string())  # Send email
		smtp.quit()
		st.success(f"✅ Email sent successfully! on: {formatted_date} please expect email delay around 10 minutes")
	except Exception as e:
		st.error(f"❌ Error: {e}")
			
#else:
	#st.warning("⚠ Please fill in the details and click 'Submit'.")
  
