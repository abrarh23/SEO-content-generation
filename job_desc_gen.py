from openai import OpenAI
import os
from dotenv import load_dotenv
load_dotenv()
import json
import traceback
from pydantic import BaseModel, Field, ValidationError
from typing import List
import pandas as pd
import time
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from googleapiclient.discovery import build

def read_input_csv() -> pd.DataFrame:
    all_job_titles = pd.read_csv(r".\data\HR Templates  - Job titles.csv")[['clean_job_titles']].copy()
    return all_job_titles

class FocusArea(BaseModel):
    focus_area: str = Field(..., description="Name of the KPI focus area")
    description: str = Field(..., description="Description of the KPI focus area")
    class Config:
        extra = "ignore"

class TeamStructure(BaseModel):
    reports_to: str = Field(..., description="The person or role this position reports to")
    collaborates_with: str = Field(..., description="Other teams or roles that this position collaborates with")
    leads: str = Field(..., description="Teams or roles this position leads (if applicable)")
    class Config:
        extra = "ignore"

class JobDetails(BaseModel):
    job_title: str = Field(..., description="The job title for the position")
    job_description: str = Field(..., description="A detailed description of the job role")
    key_responsibilities: List[str] = Field(..., description="List of key responsibilities for the role")
    skills: List[str] = Field(..., description="List of skills required for the role")
    kpis: str = Field(..., description="Key performance indicators summary")
    kpis_focus: List[FocusArea] = Field(..., description="List of KPI focus areas and descriptions")
    team_structure: TeamStructure = Field(..., description="The reporting and collaboration structure of the team")
    tools: List[str] = Field(..., description="List of tools required for the role")
    qualification: str = Field(..., description="Required qualifications for the role")
    class Config:
        extra = "ignore"

def connect_to_openai() -> OpenAI:
    OPENAI_API_KEY = os.getenv("OPENAI_API_KEY") 
    client = OpenAI(api_key=OPENAI_API_KEY)
    return client

def connect_to_google_sheets_docs():
    # Define the scope to include both Google Sheets and Google Docs
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/documents" 
    ]
    
    # Load the credentials from the JSON key file
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        r"C:\Users\Abrar\Desktop\Programs\Github\Qureos-Workspace\Modules\qureos-engineering.json", 
        scope
    )
    
    # Connect to Google Sheets
    client = gspread.authorize(credentials)
    spreadsheet_url = "https://docs.google.com/spreadsheets/d/1b3s7oy_9KLLrB46qxCVAQ4pLm4-T3RFMU-msGkovp40/edit?usp=sharing"
    spreadsheet = client.open_by_url(spreadsheet_url)
    sheet = spreadsheet.worksheet("Python")
    
    # Connect to Google Docs
    docs_service = build('docs', 'v1', credentials=credentials)
    drive_service = build('drive', 'v3', credentials=credentials) 
    
    return sheet, docs_service, drive_service

def process_response(response, prompt_tokens, completion_tokens):
    try:
        return json.loads(response), prompt_tokens, completion_tokens
    except json.JSONDecodeError as e:
        print("JSON decoding error:", e)
        traceback.print_exc()  # Print the full traceback for debugging purposes
        return {}, None, None  # Return None for tokens if there's an error
    except AttributeError as e:
        print("Attribute error:", e)
        traceback.print_exc()
        return {}, None, None  # Handle missing attributes if needed
    except Exception as e:
        print("Unexpected error:", e)
        traceback.print_exc()  # Catch any other unexpected errors
        return {}, None, None

def convert_list_to_html_bullets(data_list):
    return "<ul>" + "".join([f"<li>{item}</li>" for item in data_list]) + "</ul>"

def convert_data_to_html(content: dict):
    # Convert lists to HTML bullet points
    key_responsibilities_html = convert_list_to_html_bullets(content.get("key_responsibilities", []))
    skills_html = convert_list_to_html_bullets(content.get("skills", []))
    tools_html = convert_list_to_html_bullets(content.get("tools", []))
    return key_responsibilities_html, skills_html, tools_html

def prepare_data_for_upload(content: dict, key_responsibilities_html: str, skills_html: str, tools_html: str):
    # Helper function to convert lists to newline-separated strings
    def convert_to_string(value):
        if isinstance(value, list):
            return "\n".join(value)
        return value

    sheet_data = [
        convert_to_string(content.get("job_title", "N/A")),
        "",  # Slug (empty)
        "",  # Collection ID (empty)
        "",  # Locale ID (empty)
        "",  # Item ID (empty)
        "",  # Created On (empty)
        "",  # Updated On (empty)
        "",  # Published On (empty)
        convert_to_string(content.get("job_description", "N/A")),
        convert_to_string(content.get("key_responsibilities", "N/A")),
        key_responsibilities_html,
        convert_to_string(content.get("skills", "N/A")),
        skills_html,
        convert_to_string(content.get("kpis", "N/A")),
        convert_to_string(content["kpis_focus"][0].get("focus_area", "KPI")) if len(content["kpis_focus"]) > 0 else "",
        convert_to_string(content["kpis_focus"][0].get("description", "N/A")) if len(content["kpis_focus"]) > 0 else "",
        convert_to_string(content["kpis_focus"][1].get("focus_area", "KPI")) if len(content["kpis_focus"]) > 1 else "",
        convert_to_string(content["kpis_focus"][1].get("description", "N/A")) if len(content["kpis_focus"]) > 1 else "",
        convert_to_string(content["kpis_focus"][2].get("focus_area", "KPI")) if len(content["kpis_focus"]) > 2 else "",
        convert_to_string(content["kpis_focus"][2].get("description", "N/A")) if len(content["kpis_focus"]) > 2 else "",
        convert_to_string(content.get("team_structure", {}).get("reports_to", "N/A")),
        convert_to_string(content.get("team_structure", {}).get("collaborates_with", "N/A")),
        convert_to_string(content.get("team_structure", {}).get("leads", "N/A")),
        convert_to_string(content.get("tools", "N/A")),
        tools_html,
        convert_to_string(content.get("qualification", "N/A")),
        ""  # Google Doc Link (empty)
    ]
    
    return sheet_data

def push_to_gs(sheet: gspread.Worksheet, sheet_data: list) -> None:
    # Clear existing data and insert new data
    sheet.append_rows(sheet_data)

def get_gen_content(job_title: str):

    response = connect_to_openai().chat.completions.create(
    model="gpt-4o",
    messages=[
        {
        "role": "system",
        "content": [
            {
            "type": "text",
            "text": "Create a JSON-formatted job description with the Job title as the focal point.\n\n# Questions and Answers\n\n**Question: What does a job title provided by the user do?**  \nAnswer: Provide a concise overview of the Job title role, including its contribution to company goals and unique value. This description should be 2-3 sentences, highlighting the job title's importance. \n\n**Question: Key responsibilities of a job title provided by the user**  \nAnswer: List 6-10 main responsibilities for the Job title, action-oriented and relevant to the job's core functions. Structure as bullet points for easy reading. \n\n**Question: Skills required for a job title provided by the user**  \nAnswer: List essential skills. Include both technical skills specific to the role and soft skills like teamwork. \n\n**Question: What are the KPIs for the job title?**  \nAnswer: In a paragraph of 40-45 words.\n\n**Question: what are the 3 Key Performance Indicators for the job title\nAnswers: In a list, nested inside kpis_focus key.\n\n**Question: What is the team structure, and who does the Job title reports to, collaborates with and leads?**  \nAnswer: In json object, provide your answer corresponding to keys reports_to, collaborates_with and leads\n\n**Question: Are there any specific tools or software required for the Job title role?**  \nAnswer: List of tools in an array.\n\n**Question: What is the qualification for the job title?**  \nAnswer: Mention the education and experience required. \n\n{\n    \"job_title\": \"SEO Manager\",\n    \"job_description\": \"The SEO Manager plays a pivotal role in driving organic traffic and enhancing the online visibility of the company's digital assets. This position is crucial in achieving company growth objectives by optimizing website content and collaborating with various teams to implement effective SEO strategies.\",\n    \"key_responsibilities\": [\n        \"Develop and execute successful SEO strategies.\",\n        \"Conduct keyword research to guide content teams.\",\n        \"Review technical SEO issues and recommend fixes.\",\n        \"Optimize website content, landing pages, and paid search copy.\",\n        \"Monitor SEO performance metrics to forecast trends.\",\n        \"Collaborate with web developers and marketing teams.\",\n        \"Direct off-page optimization projects (e.g., link-building).\",\n        \"Collect data and report on traffic, rankings, and other SEO aspects.\",\n        \"Stay up to date with the latest SEO and digital marketing trends.\"\n    ],\n    \"skills\": [\n        \"Strong understanding of SEO, SEM, and digital marketing.\",\n        \"Proficient in SEO tools like Google Analytics, Ahrefs, and SEMrush.\",\n        \"Excellent analytical, problem-solving, and decision-making skills.\",\n        \"Effective communication and collaboration skills.\",\n        \"Ability to work with cross-functional teams.\"\n    ],\n    \"kpis\": \"The SEO Manager's performance is evaluated through measurable improvements in organic search rankings, increased website traffic and conversions, and successful implementation of SEO strategies.\",\n    \"kpis_focus\": [\n        {\n            \"focus_area\": \"Website Traffic\",\n            \"description\": \"Maintain or increase organic site traffic.\"\n        },\n        {\n            \"focus_area\": \"Search Rankings\",\n            \"description\": \"Improvement in search engine ranking positions.\"\n        },\n        {\n            \"focus_area\": \"Lead Generation\",\n            \"description\": \"Enhance lead conversion rates through organic channels.\"\n        }\n    ],\n    \"team_structure\": \n        {\n            \"reports_to\": \"Digital Marketing Director\",\n            \"collaborates_with\": \"Content Team, Web Developers\",\n            \"leads\": \"SEO Specialists\"\n        },\n    \"tools\": [\n        \"Google Analytics\",\n        \"Google Search Console\",\n        \"SEMrush\",\n        \"Ahrefs\",\n        \"Moz\"\n    ],\n    \"qualification\": \"Bachelor's degree in Marketing, Business, or a related field, with 3-5 years of proven experience in SEO management.\"\n}\n"
            }
        ]
        },
        {
        "role": "user",
        "content": [
            {
            "type": "text",
            "text": f"{job_title}"
            }
        ]
        },
        {
        "role": "assistant",
        "content": [
            {
            "type": "text",
            "text": "{\"job_title\":\"HR Manager\",\"job_description\":\"The HR Manager is integral to fostering a positive workplace environment by managing employee relations, recruitment, and compliance with HR policies. This role supports company growth by nurturing talent and aligning human resources practices with organizational goals.\",\"key_responsibilities\":[\"Oversee the hiring process from recruitment to onboarding.\",\"Implement HR strategies that support company objectives.\",\"Manage employee relations, including conflict resolution and performance management.\",\"Ensure compliance with employment laws and regulations.\",\"Develop training programs for employee development.\",\"Collaborate with department heads on workforce planning needs.\",\"Administer compensation and benefits programs.\"],\"skills\":[\"Strong understanding of HR principles and employment law.\",\"Proficient in human resources software like Workday or SAP SuccessFactors.\",\"Excellent communication and interpersonal skills.\",\"Strong leadership abilities.\"],\"kpis\":\"The performance of the HR Manager is measured through successful talent acquisition, reduction in employee turnover rates, enhancement of staff satisfaction, and effectively addressing workplace issues within established timelines.\",\"kpis_focus\":[{\"focus_area\":\"Talent Acquisition\",\"description\":\"Efficient filling of job vacancies as per target timeframes.\"},{\"focus_area\":\"Employee Turnover\",\"description\":\"Reduction in turnover rates year-over-year.\"},{\"focus_area\":\"Employee Satisfaction\",\"description\":\"Improvement in staff satisfaction survey scores\"}],\"team_structure\":{\"reports_to\":\"Director of Human Resources\",\"collaborates_with\":\"Department Managers, Recruitment Teams\",\"leads\":\"HR Coordinators\"},\"tools\":[\"Workday\",\"SAP SuccessFactors\",\"ADP Workforce Now\"],\"qualification\":\"Bachelor's degree in Human Resources Management or related field; 5-7 years experience managing HR functions.\"}"
            }
        ]
        }
    ],
    temperature=0,
    max_tokens=4048,
    top_p=1,
    frequency_penalty=1,
    presence_penalty=0,
    response_format={
        "type": "json_schema",
        "json_schema": {
        "name": "job_description",
        "strict": True,
        "schema": {
            "type": "object",
            "properties": {
            "job_title": {
                "type": "string",
                "description": "The title of the job position."
            },
            "job_description": {
                "type": "string",
                "description": "A detailed description of the job role and its importance."
            },
            "key_responsibilities": {
                "type": "array",
                "description": "A list of key responsibilities associated with the job role.",
                "items": {
                "type": "string"
                }
            },
            "skills": {
                "type": "array",
                "description": "A list of skills required for the job position.",
                "items": {
                "type": "string"
                }
            },
            "kpis": {
                "type": "string",
                "description": "Key performance indicators for evaluating the job performance."
            },
            "kpis_focus": {
                "type": "array",
                "description": "A list outlining the focus areas for key performance indicators.",
                "items": {
                "type": "object",
                "properties": {
                    "focus_area": {
                    "type": "string",
                    "description": "The focus area for KPI."
                    },
                    "description": {
                    "type": "string",
                    "description": "Description of the KPI focus area."
                    }
                },
                "required": [
                    "focus_area",
                    "description"
                ],
                "additionalProperties": False
                }
            },
            "team_structure": {
                "type": "object",
                "description": "The structure of the team for the job position.",
                "properties": {
                "reports_to": {
                    "type": "string",
                    "description": "The position that the job role reports to."
                },
                "collaborates_with": {
                    "type": "string",
                    "description": "Teams or individuals that the job role collaborates with."
                },
                "leads": {
                    "type": "string",
                    "description": "Positions or teams that the job role is responsible for leading."
                }
                },
                "required": [
                "reports_to",
                "collaborates_with",
                "leads"
                ],
                "additionalProperties": False
            },
            "tools": {
                "type": "array",
                "description": "A list of tools and software used in the job role.",
                "items": {
                "type": "string"
                }
            },
            "qualification": {
                "type": "string",
                "description": "Educational qualifications and experience required for the role."
            }
            },
            "required": [
            "job_title",
            "job_description",
            "key_responsibilities",
            "skills",
            "kpis",
            "kpis_focus",
            "team_structure",
            "tools",
            "qualification"
            ],
            "additionalProperties": False
        }
        }
    }
    )

    completion_tokens = response.usage.completion_tokens
    prompt_tokens = response.usage.prompt_tokens
    return process_response(response.choices[0].message.content, prompt_tokens, completion_tokens)

def get_template_structure(docs_service):
    template_doc_id = "1vhd0lkcFT0qOzAhM3ya9Ix3rc6N6hj1NlTvH4CPFc7c"
    document = docs_service.documents().get(documentId=template_doc_id).execute()
    return document.get('body').get('content')

def create_google_doc_with_formatting(docs_service, drive_service, job_title: str, template_content: str) -> str:
    # Create a new Google Doc
    document = docs_service.documents().create(body={'title': job_title + " JD Template"}).execute()
    document_id = document.get('documentId')
    print(f"Created document with ID: {document_id}. Job Title:", job_title)
    
    # Initialize the current index to track position in the document
    current_index = 1
    requests = []
    
    # Loop through template content to preserve order
    for element in template_content:
        if 'paragraph' in element:
            for text_run in element['paragraph'].get('elements', []):
                if 'textRun' in text_run:
                    text_content = text_run['textRun'].get('content', "")
                    requests.append({
                        'insertText': {
                            'location': {'index': current_index},
                            'text': text_content
                        }
                    })
                    
                    # Update current index for next insertion
                    current_index += len(text_content)

                    # Apply text style if present
                    if 'textStyle' in text_run['textRun']:
                        text_style = text_run['textRun']['textStyle']
                        style_request = {
                            'updateTextStyle': {
                                'range': {
                                    'startIndex': current_index - len(text_content),
                                    'endIndex': current_index
                                },
                                'textStyle': text_style,
                                # Updated fields parameter:
                                'fields': 'bold,italic,underline,strikethrough,link(url)' 
                            }
                        }

                        requests.append(style_request)

    # Apply the requests to the new document
    docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()

    # Grant public read access to the document
    public_permission = {
        'type': 'anyone',
        'role': 'reader'  # Change to 'writer' for public edit access
    }
    drive_service.permissions().create(
        fileId=document_id,
        body=public_permission,
        fields='id'
    ).execute()
    
    return document_id

def push_to_docs(docs_service, document_id, replacements):
    requests = []
    for placeholder, new_text in replacements.items():
        requests.append({
            'replaceAllText': {
                'containsText': {
                    'text': '{{' + placeholder + '}}',  # Ensure this matches the placeholder format in your template
                    'matchCase': True,
                },
                'replaceText': new_text  # The replacement text
            }
        })
    try:
        # Execute the batch update to replace text in Google Docs
        result = docs_service.documents().batchUpdate(
            documentId=document_id, body={'requests': requests}).execute()
        return result
    except Exception as e:
        print("An error occurred:", e)

if __name__ == "__main__":
    all_job_titles = read_input_csv()

    for idx, each_job_title in all_job_titles[283:].iterrows():
        start = time.time()
        response = get_gen_content(each_job_title["clean_job_titles"])  # Pass the correct value here
        content, prompt_tokens, completion_tokens = response
        print(json.dumps(content, indent=4))
        print("Prompt tokens:", prompt_tokens)
        print("Completion tokens:", completion_tokens)

        # Validate data
        try:
            job_details = JobDetails(**content)
        except ValidationError as e:
            print(e)

        key_responsibilities_html, skills_html, tools_html = convert_data_to_html(content)
        sheet_data = prepare_data_for_upload(content, key_responsibilities_html, skills_html, tools_html)
        # print(sheet_data)
        
        # Prepare DataFrame
        push_df = pd.DataFrame([sheet_data], columns=['job_title', 'slug', 'collection_id', 'locale_id', 'item_id', 'created_on', 'updated_on', 'published_on', 'job_description', 'key_responsibilities_text','key_responsibilities_html', 'skills_text', 'skills_html', 'kpis', 'kpis_focus_1', 'description_1', 'kpis_focus_2', 'description_2', 'kpis_focus_3', 'description_3', 'reports_to', 'collaborates_with', 'leads', 'tools_text', 'tools_html', 'qualification', 'link'])     
        sheet_data = push_df.values.tolist()  
        sheet, docs_service, drive_service = connect_to_google_sheets_docs()

        # Template handling
        template_content = get_template_structure(docs_service)
        doc_document_id = create_google_doc_with_formatting(docs_service, drive_service, each_job_title["clean_job_titles"], template_content)

        # Update Google Doc
        replacements = {col: push_df[col].iloc[0] for col in push_df.columns}
        push_to_docs(docs_service, doc_document_id, replacements)

        # Assign Google Doc link to sheet data
        google_doc_link = "https://docs.google.com/document/d/" + doc_document_id + "/copy"
        print("Google doc link:", google_doc_link)
        sheet_data[0][-1] = google_doc_link  # Assuming the link should be in the last column of the row

        # Push to Google Sheets
        print("Sheet data:", sheet_data)
        push_to_gs(sheet, sheet_data)
        
        print("Data has been pushed successfully")   
        end = time.time()
        print("Time taken:", round(end - start, 2), "seconds")
