from openai import OpenAI
import os
from dotenv import load_dotenv
load_dotenv()
import json
import traceback
from pydantic import BaseModel, Field
from typing import List, Optional
import pandas as pd
import time
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from googleapiclient.discovery import build
import html
import warnings
warnings.filterwarnings("ignore")

class SkillsToAdd(BaseModel):
    technical_skills: List[str] = Field(..., min_items=3, description="Must have at least 3 skills")
    soft_skills: List[str] = Field(..., min_items=3, description="Must have at least 3 skills")

class KPIsAndOKRs(BaseModel):
    kpis: List[str] = Field(..., min_items=3, description="Must have at least 3 kpis")
    okrs: List[str] = Field(..., min_items=3, description="Must have at least 3 okrs")

class Experience(BaseModel):
    right_example: List[str] = Field(..., min_items=3, description="Must have at least 3 right example")
    wrong_example: List[str] = Field(..., min_items=3, description="Must have at least 3 wrong example")

class Education(BaseModel):
    degree_name: str
    institution: str
    year: int
    relevant_coursework: Optional[List[str]]

class Project(BaseModel):
    project_name:str
    role: str
    tools: List[str]
    outcome: List[str]

class BasicSections(BaseModel):
    job_title_and_role_significance: str
    summary: str
    skills_to_add: SkillsToAdd
    kpis_and_okrs: KPIsAndOKRs
    experience: Experience
    education: Education
    project: Project

def read_input_csv() -> pd.DataFrame:
    all_job_titles = pd.read_csv(r"data\HR Templates  - Job titles (B2C).csv")
    all_job_titles = all_job_titles.apply(lambda x: x.str.title()).copy()
    return all_job_titles

def connect_to_openai() -> OpenAI:
    OPENAI_API_KEY = os.getenv("OPENAI_API_KEY") 
    client = OpenAI(api_key=OPENAI_API_KEY)
    return client

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

def get_openai_resp(job_title: str):
    response = connect_to_openai().chat.completions.create(
    model="gpt-4o",
    messages=[
        {
        "role": "system",
        "content": [
            {
            "type": "text",
            "text": "Your job is to write a resume template article divided into separate sections.  Return the response in JSON format.\n\n1. The first section talks briefly about the job title and its related attributes:\n\nProvide a brief description of the role's significance in the industry and relevant statistics (e.g., projected growth, average salary).\nWhen you are mentioning the statistics for project growth and average salary, mentioned that these statistics are for 2025.\nEnd the section with a new line: 'Now, we will guide you on how to write a great resume for [Job Title].'\n\nExample: \n[Job Title] professionals are essential for [brief description of the role's significance, e.g., driving business success, creating impactful designs, or leading technical innovations]. The demand for [Job Title] roles is projected to grow/shrink by [insert percentage trend in Middle East region], and the average salary ranges from [insert salary range according to Middle East region].\nA well-crafted resume is the first step toward showcasing your skills, achievements, and experience to potential employers. Now, we will guide you on how to write an impressive resume tailored for a [Job Title] role.\n\n2. Provide an example of a strong summary that highlights key skills, achievements, and career goals.\n\n3. What Skills to Add to Your [Job Title] Resume\n\nCategorize skills into two sections:\nTechnical Skills: Job-specific tools, software, or certifications.\nSoft Skills: Transferable skills like communication, problem-solving, or time management.\n\n4. What are [Job Title] KPIs and OKRs, and How Do They Fit Your Resume?\n\nWhat are top 3 KPIs pf this job title?\nWhat are top 3 OKRs of this job title?\n\n5. How to Describe Your [Job Title] Experience\n\nProvide examples of how to format the experience section using quantifiable achievements.\nUse bullet points starting with action verbs and emphasize measurable outcomes.\nInclude 3 'Right' and 'Wrong' examples to illustrate the difference.\n\n6. How to Present Your Education as a [Job Title]\n\nInclude relevant degrees, certifications, and training programs.\n\nExample structure:\nDegree/Certification Name: [Insert degree or certification name]\nInstitution: [Insert institution name]\nYear: [Insert graduation or completion year]\nRelevant Coursework (optional): [List key courses if relevant to the role].\n\n7. How to Highlight Your Projects as a [Job Title]\nDescribe key projects you've worked on that demonstrate your expertise and impact.\nInclude the project name, your role, tools/technologies used, and quantifiable outcomes.\n\nExample structure:\nProject Name: [Insert project name]\nRole: [Describe your role in the project]\nTools/Technologies: [List relevant tools or technologies used]\nOutcome: [Highlight measurable results or impact]."
            }
        ]
        },
        {
        "role": "user",
        "content": [
            {
            "type": "text",
            "text": f"job title: {job_title}"
            }
        ]
        }
    ],
    response_format={
        "type": "json_schema",
        "json_schema": {
        "name": "job_title_mentioned_by_the_user",
        "strict": True,
        "schema": {
            "type": "object",
            "properties": {
            "job_title_and_role_significance": {
                "type": "string",
                "description": "Overview of the significance and demand of job title role mentioned by the user."
            },
            "summary": {
                "type": "string",
                "description": "A compelling summary should highlight your key skills, experience, and measurable achievements in the field. It serves as your elevator pitch to grab the employer's attention according to the job title mentioned by the user"
            },
            "skills_to_add": {
                "type": "object",
                "description": "Skills categories for job title mentioned by the user.",
                "properties": {
                "technical_skills": {
                    "type": "array",
                    "description": "List of technical skills relevant to job title mentioned by the user.",
                    "items": {
                    "type": "string"
                    }
                },
                "soft_skills": {
                    "type": "array",
                    "description": "List of soft skills relevant to job title mentioned by the user.",
                    "items": {
                    "type": "string"
                    }
                }
                },
                "required": [
                "technical_skills",
                "soft_skills"
                ],
                "additionalProperties": False
            },
            "kpis_and_okrs": {
                "type": "object",
                "description": "Key Performance Indicators (KPIs) and Objectives and Key Results (OKRs) for job title mentioned by the user.",
                "properties": {
                "kpis": {
                    "type": "array",
                    "description": "Top 3 Important KPIs for a job title mentioned by the user.",
                    "items": {
                    "type": "string"
                    }
                },
                "okrs": {
                    "type": "array",
                    "description": "Top 3 OKRs for a job title mentioned by the user.",
                    "items": {
                    "type": "string"
                    }
                }
                },
                "required": [
                "kpis",
                "okrs"
                ],
                "additionalProperties": False
            },
            "experience": {
                "type": "object",
                "description": "Examples of how to present experience related to your job title.",
                "properties": {
                "right_example": {
                    "type": "array",
                    "description": "Correct examples of describing experience of job title mentioned by the user.",
                    "items": {
                    "type": "string"
                    }
                },
                "wrong_example": {
                    "type": "array",
                    "description": "Incorrect examples of experience of job title mentioned by the user.",
                    "items": {
                    "type": "string"
                    }
                }
                },
                "required": [
                "right_example",
                "wrong_example"
                ],
                "additionalProperties": False
            },
            "education": {
                "type": "object",
                "description": "Education details for a particular job title mentioned by the user.",
                "properties": {
                "degree_name": {
                    "type": "string",
                    "description": "The degree or certification obtained."
                },
                "institution": {
                    "type": "string",
                    "description": "Name of the educational institution."
                },
                "year": {
                    "type": "string",
                    "description": "Year of graduation or completion."
                },
                "relevant_coursework": {
                    "type": "array",
                    "description": "Courses relevant to job title mentioned by the user.",
                    "items": {
                    "type": "string"
                    }
                }
                },
                "required": [
                "degree_name",
                "institution",
                "year",
                "relevant_coursework"
                ],
                "additionalProperties": False
            },
            "project": {
                "type": "object",
                "description": "Project details for a particular job title mentioned by the user.",
                "properties": {
                "project_name": {
                    "type": "string",
                    "description": "Write the project name relevant to the job title mentioned by the user"
                },
                "role": {
                    "type": "string",
                    "description": "Describe your role in the project"
                },
                "tools": {
                    "type": "array",
                    "description": "List relevant tools or technologies used in this project.",
                    "items": {
                    "type": "string"
                    }
                },
                "outcome": {
                    "type": "array",
                    "description": "Highlight measurable results or impact relevant to job title mentioned by the user.",
                    "items": {
                    "type": "string"
                    }
                }
                },
                "required": [
                "project_name",
                "role",
                "tools",
                "outcome"
                ],
                "additionalProperties": False
            }
            },
            "required": [
            "job_title_and_role_significance",
            "summary",
            "skills_to_add",
            "kpis_and_okrs",
            "experience",
            "education",
            "project"
            ],
            "additionalProperties": False
        }
        }
    },
    temperature=1,
    max_completion_tokens=2048,
    top_p=1,
    frequency_penalty=0,
    presence_penalty=0
    )

    completion_tokens = response.usage.completion_tokens
    prompt_tokens = response.usage.prompt_tokens
    return process_response(response.choices[0].message.content, prompt_tokens, completion_tokens)

def convert_list_html(class_list: list) -> str:
    if isinstance(class_list, list):
        # Ensure each item is a string (if it's a list, join the items first)
        html_encoded_list = []
        for item in class_list:
            if isinstance(item, list):  # Handle if item is a list
                item = ", ".join(map(str, item))  # Join list items into a string
            html_encoded_list.append(html.escape(str(item)))  # Escape the string
        html_list = "<ul>\n"
        for item in html_encoded_list:
            html_list += f"  <li>{item}</li>\n"
        html_list += "</ul>"
        return html_list
    return "N/A"

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
        r".\qureos-engineering.json", 
        scope
    )
    
    # Connect to Google Sheets
    client = gspread.authorize(credentials)
    spreadsheet_url = "https://docs.google.com/spreadsheets/d/1b3s7oy_9KLLrB46qxCVAQ4pLm4-T3RFMU-msGkovp40/edit?gid=1823102495#gid=1823102495"
    spreadsheet = client.open_by_url(spreadsheet_url)
    sheet = spreadsheet.worksheet("Python (resume)")
    
    # Connect to Google Docs
    docs_service = build('docs', 'v1', credentials=credentials)
    drive_service = build('drive', 'v3', credentials=credentials) 
    
    return sheet, docs_service, drive_service

def convert_dict_to_df(job_title: str) -> pd.DataFrame:
    openai_resp = get_openai_resp(job_title)
    text_resp = openai_resp[0]
    print(json.dumps(text_resp))
    validated_data = BasicSections(**text_resp)
    resume_df = pd.DataFrame(index=[0])
    resume_df['job_title'] = f'{job_title}'
    resume_df['job_title_and_role_significance'] = text_resp.get("job_title_and_role_significance")
    resume_df['summary'] = text_resp.get("summary")
    resume_df['technical_skills'] = convert_list_html(text_resp.get("skills_to_add").get("technical_skills", []))
    resume_df['soft_skills'] = convert_list_html(text_resp.get("skills_to_add").get("soft_skills", []))
    resume_df['kpis_lst'] = convert_list_html(text_resp.get("kpis_and_okrs").get("kpis", []))
    resume_df['okrs_lst'] = convert_list_html(text_resp.get("kpis_and_okrs").get("okrs", []))
    resume_df['exp_right_ex'] = convert_list_html(text_resp.get("experience").get("right_example", []))
    resume_df['exp_wrong_ex'] = convert_list_html(text_resp.get("experience").get("wrong_example", []))
    resume_df['edu_degree_name'] = text_resp.get("education").get("degree_name", "N/A")
    resume_df['edu_institution'] = text_resp.get("education").get("institution", "N/A")
    resume_df['edu_year'] = text_resp.get("education").get("year", "N/A")
    resume_df['edu_relevant_coursework'] = convert_list_html(text_resp.get("education").get("relevant_coursework", []))
    resume_df['project_name'] = text_resp.get("project").get("project_name", "N/A")
    resume_df['project_role'] =  text_resp.get("project").get("role", "N/A")
    resume_df['project_tools'] =  convert_list_html(text_resp.get("project").get("tools", []))
    resume_df['project_outcome'] =  convert_list_html(text_resp.get("project").get("outcome", []))

    return resume_df

def push_to_gs(sheet: gspread.Worksheet, sheet_data: list) -> None:
    # Clear existing data and insert new data
    sheet.append_rows(sheet_data)



if __name__=="__main__":
    job_title_list = read_input_csv()

    for _, row in job_title_list[:1].iterrows():
        resume_df = convert_dict_to_df(row['job_titles'])
        sheet_data = resume_df.values.tolist() 
        sheet, docs_service, drive_service = connect_to_google_sheets_docs()
        # Push to Google Sheets
        push_to_gs(sheet, sheet_data)
