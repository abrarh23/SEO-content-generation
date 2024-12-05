from openai import OpenAI
import os
from dotenv import load_dotenv
load_dotenv()
import json
import traceback
from pydantic import BaseModel, Field, ValidationError
from typing import List, Dict
import pandas as pd
import time
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from googleapiclient.discovery import build

class Question(BaseModel):
    interview_question: str = Field(..., description="The interview question text.")
    model_answer: str = Field(..., description="An ideal model answer for the question.")
    example: str = Field(..., description="An example scenario that answers the question.")
    what_hiring_managers_should_pay_attention_to: List[str] = Field(
        ..., description="Key points hiring managers should focus on during the interview."
    )

class LevelQuestions(BaseModel):
    questions: List[Question] = Field(..., description="A list of questions for this job level.")

class JobTitleQuestions(BaseModel):
    entry_level: LevelQuestions = Field(..., description="Entry-level interview questions.")
    mid_level: LevelQuestions = Field(..., description="Mid-level interview questions.")
    senior_level: LevelQuestions = Field(..., description="Senior-level interview questions.")

class InterviewData(BaseModel):
    job_title: Dict[str, JobTitleQuestions] = Field(
        ..., description="Questions categorized by job title and hierarchy levels."
    )

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

def get_openai_resp(job_title):
    response = connect_to_openai().chat.completions.create(
    model="gpt-4o",
    messages=[
        {
        "role": "system",
        "content": [
            {
            "type": "text",
            "text": "Create a comprehensive interview questions template for a specific job title provided by the user. This template is designed for recruiters or hiring managers to assess candidates' skills, abilities, and suitability for the specified role. The focus will be on both technical expertise and relevant soft skills, providing real-world context and emphasizing measurable outcomes.\n\n# Template Structure\n\n- **Job Title Hierarchy**: Organize questions based on job title levels such as entry-level, mid-level, and senior-level.\n\n- **Types of Questions**: Suggest and include questions for technical skills, behavioral insights, and soft-skill evaluation.\n\n# Components of Each Question\n\n1. **Interview Question**: Specific question tailored to assess skills relevant to the job title.\n   \n2. **Model Answer**: Provide a comprehensive, detailed, and realistic example of how a strong candidate might respond.\n   \n3. **Example**: Offer a specific, practical scenario illustrating the model answer in action.\n   \n4. **What Hiring Managers Should Pay Attention To**: Highlight key points and red flags hiring managers should evaluate when listening to the candidate’s response.\n\n# Additional Guidelines\n\n- Maintain a professional, clear, and concise tone.\n- Ensure questions are optimal for how hiring managers might search for them as resources.\n\n# Output Format\n\nOrganize the output in a structured bullet-point format for easy readability and quick reference. Where applicable, provide additional context or examples using placeholders.\n\n# Examples\n\n*Example for Entry-Level Position:*\n\n- **Question**: \"Describe a situation wide additional context or examples using placeholders.\nhere you had to quickly learn new skills to complete a task.\"\n  \n- **Model Answer**: \"A strong candidate might explain how they identified the necessary skills, resources used (such as online courses or mentorship), steps taken to acquire these skills, and the outcome of their efforts.\"\n\n- **Example**: \"For instance, I had to learn a new software tool within two weeks to assist my team in a project.\"\n\n- **What Hiring Managers Should Pay Attention To**: Listen for adaptability, the initiative to seek resources, and problem-solving abilities.\n\n*Example for Senior-Level Position (Include more detailed scenarios):*\n\n- **Question**: \"How do you manage conflicts within your team?\"\n  \n- **Model Answer**: \"A strong candidate could describe implementing structured conflict-resolution strategies, involving identifying the root cause and mediating a resolution.\"\n\n- **Example**: \"In a software development project, there were differing opinions on the implementation strategy, which I addressed by facilitating a team meeting to discuss compromises.\"\n\n- **What Hiring Managers Should Pay Attention To**: Notice leadership abilities, communication skills, and effectiveness in resolving conflicts.\n\n# Notes\n\nEnsure the template remains adaptable to different job titles by using placeholders where specific context might change. Tailor each question and associated details to the hierarchical level and specific requirements of the role while maintaining domain relevance, practicality and at least 2 questions per hierarchial level. Return the response in json format."
            }
        ]
        },
        {
        "role": "user",
        "content": [
            {
            "type": "text",
            "text": f"job_title: {job_title}"
            }
        ]
        }
    ],
    response_format={
        "type": "json_schema",
        "json_schema": {
        "name": "job_interview_schema",
        "strict": True,
        "schema": {
            "type": "object",
            "properties": {
            "job_title": {
                "type": "object",
                "properties": {
                "entry_level": {
                    "type": "object",
                    "properties": {
                    "questions": {
                        "type": "array",
                        "description": "List of interview questions for entry-level candidates.",
                        "items": {
                        "type": "object",
                        "properties": {
                            "interview_question": {
                            "type": "string",
                            "description": "Specific question tailored to assess skills relevant to the job title."
                            },
                            "model_answer": {
                            "type": "string",
                            "description": "Provide a comprehensive, detailed, and realistic example of how a strong candidate might respond."
                            },
                            "example": {
                            "type": "string",
                            "description": "Offer a specific, practical scenario illustrating the model answer in action."
                            },
                            "what_hiring_managers_should_pay_attention_to": {
                            "type": "array",
                            "description": "Highlight key points and red flags hiring managers should evaluate when listening to the candidate’s response.",
                            "items": {
                                "type": "string"
                            }
                            }
                        },
                        "required": [
                            "interview_question",
                            "model_answer",
                            "example",
                            "what_hiring_managers_should_pay_attention_to"
                        ],
                        "additionalProperties": False
                        }
                    }
                    },
                    "required": [
                    "questions"
                    ],
                    "additionalProperties": False
                },
                "mid_level": {
                    "type": "object",
                    "properties": {
                    "questions": {
                        "type": "array",
                        "description": "List of interview questions for mid-level candidates.",
                        "items": {
                        "type": "object",
                        "properties": {
                            "interview_question": {
                            "type": "string",
                            "description": "Specific question tailored to assess skills relevant to the job title."
                            },
                            "model_answer": {
                            "type": "string",
                            "description": "Provide a comprehensive, detailed, and realistic example of how a strong candidate might respond."
                            },
                            "example": {
                            "type": "string",
                            "description": "Offer a specific, practical scenario illustrating the model answer in action."
                            },
                            "what_hiring_managers_should_pay_attention_to": {
                            "type": "array",
                            "description": "Highlight key points and red flags hiring managers should evaluate when listening to the candidate’s response.",
                            "items": {
                                "type": "string"
                            }
                            }
                        },
                        "required": [
                            "interview_question",
                            "model_answer",
                            "example",
                            "what_hiring_managers_should_pay_attention_to"
                        ],
                        "additionalProperties": False
                        }
                    }
                    },
                    "required": [
                    "questions"
                    ],
                    "additionalProperties": False
                },
                "senior_level": {
                    "type": "object",
                    "properties": {
                    "questions": {
                        "type": "array",
                        "description": "List of interview questions for senior-level candidates.",
                        "items": {
                        "type": "object",
                        "properties": {
                            "interview_question": {
                            "type": "string",
                            "description": "Specific question tailored to assess skills relevant to the job title."
                            },
                            "model_answer": {
                            "type": "string",
                            "description": "Provide a comprehensive, detailed, and realistic example of how a strong candidate might respond."
                            },
                            "example": {
                            "type": "string",
                            "description": "Offer a specific, practical scenario illustrating the model answer in action."
                            },
                            "what_hiring_managers_should_pay_attention_to": {
                            "type": "array",
                            "description": "Highlight key points and red flags hiring managers should evaluate when listening to the candidate’s response.",
                            "items": {
                                "type": "string"
                            }
                            }
                        },
                        "required": [
                            "interview_question",
                            "model_answer",
                            "example",
                            "what_hiring_managers_should_pay_attention_to"
                        ],
                        "additionalProperties": False
                        }
                    }
                    },
                    "required": [
                    "questions"
                    ],
                    "additionalProperties": False
                }
                },
                "required": [
                "entry_level",
                "mid_level",
                "senior_level"
                ],
                "additionalProperties": False
            }
            },
            "required": [
            "job_title"
            ],
            "additionalProperties": False
        }
        }
    },
    temperature=1,
    max_tokens=16383,
    top_p=1,
    frequency_penalty=0,
    presence_penalty=2
    )

    completion_tokens = response.usage.completion_tokens
    prompt_tokens = response.usage.prompt_tokens
    return process_response(response.choices[0].message.content, prompt_tokens, completion_tokens)

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
        "/home/qureos/GrowthTeam/Qureos-Workspace/Modules/qureos-a1006.json", 
        scope
    )
    
    # Connect to Google Sheets
    client = gspread.authorize(credentials)
    spreadsheet_url = "https://docs.google.com/spreadsheets/d/1b3s7oy_9KLLrB46qxCVAQ4pLm4-T3RFMU-msGkovp40/edit?gid=1823102495#gid=1823102495"
    spreadsheet = client.open_by_url(spreadsheet_url)
    sheet = spreadsheet.worksheet("Python (interview)")
    
    # Connect to Google Docs
    docs_service = build('docs', 'v1', credentials=credentials)
    drive_service = build('drive', 'v3', credentials=credentials) 
    
    return sheet, docs_service, drive_service

def get_template_structure(docs_service):
    template_doc_id = "10TYSRLcjeYudPNx3QzWSXwL2q4gTIcWzsaKFiKzjnHs"
    document = docs_service.documents().get(documentId=template_doc_id).execute()
    return document.get('body').get('content')

def create_google_doc_with_formatting(docs_service, drive_service, job_title: str, template_content: str) -> str:
    # Create a new Google Doc
    document = docs_service.documents().create(body={'title': job_title + " Interview Questions Template"}).execute()
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

if __name__ == "__main__":
    time_start = time.time()
    job_title = "data analyst"
    # response = get_openai_resp(job_title)
    # content, prompt_tokens, completion_tokens = response
    # print(json.dumps(content, indent=4))
    # print("Prompt tokens:", prompt_tokens)
    # print("Completion tokens:", completion_tokens)

    # validated_data = InterviewData(job_title=content['job_title'])

    sheet, docs_service, drive_service = connect_to_google_sheets_docs()
    template_content = get_template_structure(docs_service)
    doc_document_id = create_google_doc_with_formatting(docs_service, drive_service, job_title, template_content)

    google_doc_link = "https://docs.google.com/document/d/" + doc_document_id + "/copy"
    print(google_doc_link)