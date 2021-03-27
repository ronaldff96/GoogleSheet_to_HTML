#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# =============================================================================
# Created By  : Ronald Fonseca
# Created Date: March 27, 2021
# =============================================================================
"""
This Script extracts information from a Google Sheets' example spredsheet and
converts some required information into a HTML file using a template.

This script shows a basic use of the Google Sheets API and some commonly used 
python modules. 
"""
# =============================================================================
# Imports
# =============================================================================
from __future__ import print_function
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import itertools
from collections import Counter
from tabulate import tabulate
from time import sleep
import webbrowser

#This is required by the Google Sheets API to request permissions.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and ranges of interest of the Example Spreadsheet.
Spreadsheet_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
ranges_to_lookup = ['Class Data!A2:A','Class Data!B2:B','Class Data!C2:C','Class Data!E2:E']

def get_service():
	"""
	Basic configuration to execute the Google Sheets API
	This function returns the service after requesting permisions to the user.
	"""
	creds = None
	# The file token.json stores the user's access and refresh tokens, and is
	# created automatically when the authorization flow completes for the first
	# time.
	if not os.path.exists('credentials.json'):
		print('''There is no credentials.json file in the current directory and it is necessary to execute the script. 
You will be redirected to: https://developers.google.com/sheets/api/quickstart/python so you can create your credentials to use the script.

Find more information in how to use this script in https://github.com/ronaldff96/GoogleSheet_to_HTML.
''')
		sleep(2)
		webbrowser.open('https://developers.google.com/sheets/api/quickstart/python')
		sleep(10)
		exit()
	if os.path.exists('token.json'):
	    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
	# If there are no (valid) credentials available, let the user log in.
	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
		    creds.refresh(Request())
		else:
		    flow = InstalledAppFlow.from_client_secrets_file(
		        'credentials.json', SCOPES)
		    creds = flow.run_local_server(port=0)
		# Save the credentials for the next run
		with open('token.json', 'w') as token:
			token.write(creds.to_json())

	service = build('sheets', 'v4', credentials=creds)
	return service

def get_values(service):
	#Get the title of the Google Spreadsheet.
	title = service.spreadsheets().get(spreadsheetId=Spreadsheet_ID).execute()['properties']['title']

	#Get the values of the ranges that we want.
	result = service.spreadsheets().values().batchGet(spreadsheetId=Spreadsheet_ID, ranges=ranges_to_lookup).execute()
	values = result.get('valueRanges', [])

	#Assign variables for the values that we are interested in.
	name = list(itertools.chain(*values[0]['values']))
	gender = list(itertools.chain(*values[1]['values']))
	level = list(itertools.chain(*values[2]['values']))
	major = list(itertools.chain(*values[3]['values']))

	#Count and separate Male and Female for the gender statistics table.
	gender_counter = dict(Counter(gender))
	ladies = gender_counter['Female']
	gentlemen = gender_counter['Male']
	total = ladies + gentlemen
	ladies = f"{ladies/total*100:.1f}%"
	gentlemen = f"{gentlemen/total*100:.1f}%"

	#Create an array of lists with the values: Student Name, Class Level, and Major. Obtained from the Google Spreadsheet.
	class_table = [[name[x],level[x],major[x]] for x in range(len(name))]

	#Convert the array into HTML using the tabulate module.
	headers = ['Student Name', 'Class Level', 'Major']
	class_table_html = tabulate(class_table, tablefmt='html', headers= headers)
	return title, ladies, gentlemen, class_table_html

def generate_HTML(title, ladies, gentlemen, class_table_html):
	#Open the HTML template and extract its content.
	with open('Template.html', 'r', encoding='utf-8') as f:
		html = f.read()

	#Modify the template with the values obtained with the get_values function.
	html = html.replace('%SHEET_NAME%', title)
	html = html.replace('%LADIES%', ladies)
	html = html.replace('%GENTLEMEN%', gentlemen)
	html = html.replace('%CLASS_TABLE%', class_table_html)

	#Generate the HTML file
	file_directory = 'Output File'
	filename = title.lower().replace(' ', '_')
	if not os.path.exists(file_directory):
		os.makedirs(file_directory)
	with open(f'{file_directory}/{filename}.html', 'w', encoding='utf-8') as f:
		f.write(html)
	print (f'File {filename}.html created and stored in "{file_directory}" folder.')

if __name__ == '__main__':
	service = get_service()
	title, ladies, gentlemen, class_table_html = get_values(service)
	generate_HTML(title, ladies, gentlemen, class_table_html)
	sleep(5)