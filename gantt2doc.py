import datetime
#import docx 
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.http import MediaFileUpload
import io
from googleapiclient.errors import HttpError
import pandas as pd
from gspread_formatting import *
import requests
import json
from google.oauth2 import service_account
import googleapiclient.discovery
import streamlit as st
from datetime import datetime


scope = ['https://www.googleapis.com/auth/drive']
service_account_json_key = 'omer-python-prac1-6e5aa86322e2.json'
credentials = service_account.Credentials.from_service_account_file(
                              filename=service_account_json_key, 
                              scopes=scope)
service = build('drive', 'v3', credentials=credentials)

client = gspread.authorize(credentials)

sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1p2iuT7jrXTM1y1drInAs-6biDqlL8VjRF_igfzmKT8M/edit#gid=0')

worksheet = client.open('Schedule').worksheet('Scheduler')

service = googleapiclient.discovery.build('sheets', 'v4', credentials=credentials)

spreadsheet_id = '1p2iuT7jrXTM1y1drInAs-6biDqlL8VjRF_igfzmKT8M'
range_name = 'B2:AA17'  

request = service.spreadsheets().get(spreadsheetId=spreadsheet_id, ranges=[range_name], includeGridData=True)
response = request.execute()

response_dict = response['sheets'][0]['data'][0]['rowData']

json_data_object = json.dumps(response['sheets'][0]['data'][0]['rowData'], skipkeys = True)
 
# Writing to json
with open("response_data.json", "w") as outfile:
    outfile.write(json_data_object)
    
persons = ['Alice','Bob','Charlie','Dave','Fred','Greg','Henry','Indira','Juliet','Ken','Larry','Mike','Neom','Oscar','Pepe']

class Job:
    def __init__(self, task):
        #if not task:
        #    raise ValueError("Task value must be provided.")        
        self.task = task
        self.team = []
        self.start_date = None
        self.end_date = None
    
    def add_person(self, name):
        assert name in persons
        if name not in self.team:
            self.team.append(name)
        
    def set_start_date(self, start_date):
        if self.start_date is None:
            self.start_date = start_date
        
    def set_end_date(self, end_date):
        self.end_date = end_date
        
    def to_dict(self):
        return {
            'Task': self.task,
            'Start date': self.start_date,
            'End date': self.end_date,
            'Team': self.team
        }    
            
    def __str__(self):
        team_str = ", ".join(self.team)
        return f"Job instance {id(self)}\nTask: {self.task}\nStart Date: {self.start_date}\nEnd Date: {self.end_date}\nTeam: {team_str}"


jobs = []

def create_jobs(response_dict):
    
    # iterate over the rows from top to bottom
    for i in range(len(response_dict)):
    
    # determine length of each row. empty cells towards the end arent registered
    # so each row length is different depending upon if cell contents/format are non null
    
        row_len = len(response_dict[i]["values"])
       
        previous_task = None
        previous_job = None
        previous_date = None
        previous_color = None
        bkgrd_color = None
        current_job = None
        current_task = None
        task = None
    
    # iterate across the columns for each row  
        for j in range(row_len):

            current_date = str(response_dict[0]["values"][j]['formattedValue'])
            cell_keys = response_dict[i]["values"][j].keys()
            cell_text = None            
        
        # Determine start and end dates for the entire schedule block.
        # This goes into heading of word document like
        # "Schedule for the period from start_date till end_date"
            if i == 0 and j ==0:
                schedule_start_date = response_dict[i]["values"][j]['formattedValue']
                
            if i ==0 and j == row_len-1:
                schedule_end_date = response_dict[i]["values"][j]['formattedValue']
                print(f"Schedule from {schedule_start_date} to {schedule_end_date}")            
            
            if i>0 :
                                
                if 'formattedValue' in cell_keys:
                    cell_text = response_dict[i]["values"][j]['formattedValue']
                    task = cell_text
                    
                if 'effectiveFormat' in cell_keys :
                    bkgrd_color = response_dict[i]["values"][j]['effectiveFormat']['backgroundColor'].values()
                else:
                    bkgrd_color = None
                
                if task == 'vac':
                    pass 
                
                if bkgrd_color:
                    
                    # Block 1. Change from previous blank cell to colored cell
                    if not previous_color:
                        job = get_job_by_task(task)
                        if not job:
                            job = Job(task)
                            job.set_start_date(current_date)
                            
                            jobs.append(job)
                        job.add_person(persons[i-1])
                        #print(f"{current_date}: from block 1, current {job}")
                        previous_color = bkgrd_color
                        previous_date = current_date
                        previous_task = task
                        
                    # Block 2. Change from one color cell to another color cell
                    elif previous_color!=bkgrd_color:
                        
                        job = get_job_by_task(task)
                        if not job:
                            job = Job(task)
                            job.set_start_date(current_date)
                            jobs.append(job)
                        job.add_person(persons[i-1])
                        #print(f"{current_date}: from block 2, current {job}")
                        #print(f"current col: {bkgrd_color}, prev col: {previous_color}")
                        previous_job=get_job_by_task(previous_task) #THIS IS FAULTY LOGIC
                        previous_job.set_end_date(previous_date) #ONLY TO FIRE CONDITIONALLY
                        #print(f"{current_date}: from block 2, previous {job}")
                        previous_color = bkgrd_color
                        previous_date = current_date
                        previous_task = task
                        
                    # Block 3. Continue from previous colored cell    
                    else:
                        previous_color = bkgrd_color
                        previous_date = current_date
                        previous_task = task
                        
                # Block 4. Change from job to blank cell
                elif previous_color:
                    previous_job=get_job_by_task(previous_task)
                    previous_job.set_end_date(previous_date)
                    #print(f"{current_date}: from block 4, previous {job}")
                    previous_color = None
                    previous_date = None
                    previous_task = None
                    
                # Block 5. Continue from blank cell to blank cell    
                else:
                    previous_color = None
                    previous_date = current_date
                    previous_task = None
                                        
    return jobs

def get_job_by_task(task_to_find):
    for job in jobs:
        if job.task == task_to_find:
            return job

jobs = create_jobs(response_dict)    
 
jobs = sorted(jobs, key=lambda x: x.start_date)

job_dict = [job.to_dict() for job in jobs]


def callback(request_id, response, exception):
    ids = []
    if exception:
        # Handle error
        print(exception)
    else:
        print(f"Request_Id: {request_id}")
        print(f'Permission Id: {response.get("id")}')
        ids.append(response.get("id"))
    return ids


def create_doc_file(creds):
    try:
        current_timestamp = datetime.now().strftime("%Y%m%d%H%M")
        service = build('docs', 'v1', credentials=creds)
        drive_service = build('drive', 'v3', credentials=credentials)
        title = f"Schedule"+current_timestamp
        body = {'title': title}
        document = service.documents().create(body=body).execute()
        document_id = document['documentId']
        print(f"Created new document with ID: {document_id}")
        EMAIL_ADDRESS = 'ofayyaz@gmail.com'
    
        batch = drive_service.new_batch_http_request(callback=callback)
        user_permission = {
            "type": "user",
            "role": "writer",
            "emailAddress": EMAIL_ADDRESS,
        }
        batch.add(
            drive_service.permissions().create(
                fileId=document_id,
                body=user_permission,
                fields="id",
            )
        )
        batch.execute()
        #print(f'Permission added. Permission ID: {permission["id"]}')  
        return service, document_id
       
        
    except Exception as e:
        print(f'Error adding permission: {e}')
    
    
def write_to_google_docs(service, document_id, content):        
    # Update the content of the document
    # Convert data to text
    formatted_text = ''
    for entry in content:
        formatted_text += f"Task: {entry['Task']}\n"
        formatted_text += f"Start date: {entry['Start date']}\n"
        formatted_text += f"End date: {entry['End date']}\n"
        formatted_text += f"Team: {', '.join(entry['Team'])}\n\n"


    # Append the formatted text to the Google Docs document
    requests = [
        {
            'insertText': {
                'location': {'index': 1},
                'text': formatted_text,
            }
        }
    ]
    result = service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()

    print(f"Data has been written to Google Docs document with document ID: {document_id}")

if __name__ == '__main__':
    svc, doc_id = create_doc_file(credentials)
    write_to_google_docs(svc, doc_id,job_dict)

    
