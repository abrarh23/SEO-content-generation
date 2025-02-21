# import libraries
import openai
import os
import json
import time
from dotenv import load_dotenv
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import gspread

load_dotenv()

response_format = {
        "type": "json_schema",
        "json_schema": {
        "name": "skills_schema",
        "strict": True,
        "schema": {
            "type": "object",
            "properties": {
                "introduction": {
                    "type": "object",
                    "properties": {
                        "overview": {"type": "string"},
                        "impact_on_success": {"type": "string"},
                        "adaptation_importance": {"type": "string"}
                    },
                    "required": ["overview", "impact_on_success", "adaptation_importance"],
                    "additionalProperties": False
                },
                "skill_progression": {
                    "type": "object",
                    "properties": {
                        "beginner": {
                            "type": "object",
                            "properties": {
                                "skills": {"type": "array", "items": {"type": "string"}, "min":4, "max":4},
                                "examples_with_action_steps": {"type": "array", "items": {"type": "string"}, "min":4, "max":4}
                                # "action_steps": {"type": "array", "items": {"type": "string"}, "min":4, "max":4}
                            },
                            "required": ["skills", "examples_with_action_steps"],
                            "additionalProperties": False
                        },
                        "intermediate": {
                            "type": "object",
                            "properties": {
                                "skills": {"type": "array", "items": {"type": "string"}, "min":4, "max":4},
                                "examples_with_action_steps": {"type": "array", "items": {"type": "string"}, "min":4,
                                "max":4}
                                # "action_steps": {"type": "array", "items": {"type": "string"}, "min":4, "max":4}
                            },
                            "required": ["skills", "examples_with_action_steps"],
                            "additionalProperties": False
                        },
                        "advanced": {
                            "type": "object",
                            "properties": {
                                "skills": {"type": "array", "items": {"type": "string"}, "min":4, "max":4},
                                "examples_with_action_steps": {"type": "array", "items": {"type": "string"}, "min":4,
                                "max":4}
                                # "action_steps": {"type": "array", "items": {"type": "string"}, "min":4, "max":4}
                            },
                            "required": ["skills", "examples_with_action_steps"],
                            "additionalProperties": False
                        }
                    },
                    "required": ["beginner", "intermediate", "advanced"],
                    "additionalProperties": False
                },
                "top_skills_2025": {
                    "type": "object",
                    "properties": {
                        "technical_skills": {"type": "array", "items": {"type": "string"}},
                        "soft_skills": {"type": "array", "items": {"type": "string"}},
                        "industry_trends": {"type": "array", "items": {"type": "string"}},
                        "future_requirements": {"type": "array", "items": {"type": "string"}}
                    },
                    "required": ["technical_skills", "soft_skills", "industry_trends", "future_requirements"],
                    "additionalProperties": False
                },
                "top_influencers": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string"},
                            "expertise": {"type": "string"},
                            "why_follow": {"type": "string"}
                        },
                        "min": 1,
                        "max": 4,
                        "required": ["name", "expertise", "why_follow"],
                        "additionalProperties": False
                    },
                },
                "learning_resources": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "course_link": {"type": "string"},
                            # "offerings": {"type": "array", "items": {"type": "string"}},
                            # "specializations": {"type": "array", "items": {"type": "string"}},
                            "why_recommended": {"type": "string"}
                        },
                        "required": ["course_link", "why_recommended"],
                        "additionalProperties": False
                    },
                }
            },
            "required": [
                "introduction",
                "skill_progression",
                "top_skills_2025",
                "top_influencers",
                "learning_resources"
            ],
            "additionalProperties": False
            }
        }
    }

# connect to openai
def connect_to_openai():
    openai.api_key = os.getenv("OPENAI_API_KEY")
    return openai

def connect_to_google_sheets_docs() -> gspread.Worksheet:
    # Define the scope to include both Google Sheets and Google Docs
    scope = [
        "https://www.googleapis.com/auth/spreadsheets"
    ]
    
    # Load the credentials from the JSON key file
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        r"/home/abdrafay/AllWork/Qureos/AllWork/Modules/qureos-a1006.json", 
        scope
    )
    
    # Connect to Google Sheets
    client = gspread.authorize(credentials)
    spreadsheet_url = "https://docs.google.com/spreadsheets/d/1b3s7oy_9KLLrB46qxCVAQ4pLm4-T3RFMU-msGkovp40/edit?usp=sharing"
    spreadsheet = client.open_by_url(spreadsheet_url)
    sheet = spreadsheet.worksheet("Python (Skills)")
    
    return sheet

# request to openai for skills generation
def skills_openai(profession: str) -> tuple:
    response = connect_to_openai().chat.completions.create(
    model="gpt-4o",
    messages=[
        {
        "role": "system",
        "content": [
            {
            "type": "text",
            "text": """Create a comprehensive skills guide for a specific profession. The focus should be on current and future skill requirements, career progression, and learning resources.

# Template Structure
The response should be organized in a clear JSON format with the following sections:

1. Introduction
- Overview of why skills matter in this profession
- Impact on success and innovation
- Industry adaptation importance

2. Skill Progression
- Beginner level skills with examples with actionable steps
- Intermediate level skills with examples with actionable steps
- Advanced level skills with examples with actionable steps

3. Top Skills for 2025
- Technical skills specific to the profession
- Essential soft skills
- Industry trends and their impact
- Future skill requirements

4. Top Influencers
- List of 5 influential professionals
- Their areas of expertise
- Reasons to follow them

5. Learning Resources
- List of link of minimum 1 and maximum 2 top courses
- Types of courses/certifications offered
- Specialization areas
- Why they are recommended

Return the response in JSON format with snake_case naming convention."""
            }
        ]
        },
        {
        "role": "user",
        "content": [
            {
            "type": "text",
            "text": f"profession: {profession}"
            }
        ]
        }
    ],
    response_format=response_format,
    temperature=1,
    max_tokens=4096,
    top_p=1,
    frequency_penalty=0,
    presence_penalty=0
    )

    completion_tokens = response.usage.completion_tokens
    prompt_tokens = response.usage.prompt_tokens
    return response.choices[0].message.content, prompt_tokens, completion_tokens

def flatten_dict(d: dict, parent_key='', sep='_') -> dict:
    items = []
    if isinstance(d, list) or isinstance(d, dict):
        for k, v in d.items():
            new_key = parent_key + sep + k if parent_key else k
            # Stop flattening deeper for `skill_progression` keys
            if parent_key.startswith("skill_progression") and parent_key.count(sep) >= 1:
                items.append((parent_key, d))
                break
            elif isinstance(v, dict):
                items.extend(flatten_dict(v, new_key, sep=sep).items())
            elif isinstance(v, list):
                for i, l in enumerate(v):
                    items.extend(flatten_dict(l, new_key + sep + str(i), sep=sep).items())
            else:
                items.append((new_key, v))
    return dict(items)

def process_skill_progression(data: dict) -> dict:
    for key, value in data["skill_progression"].items():
        obj = []
        # print(len(value['skills']), len(value['examples']), len(value['action_steps']))
        for i in range(0, len(value['skills'])):
            dic = {}
            dic['name'] = value['skills'][i]
            dic['examples_with_action_steps'] = value['examples_with_action_steps'][i]
            # dic['action_steps'] = value['action_steps'][i]
            # dic[value['skills'][i]] = [value['examples'][i], value['action_steps'][i]] 
            obj.append(dic)
        data['skill_progression'][key] = obj
    print(data, 'data')
    
    for key, value in data['skill_progression'].items():
        for i in range(0, len(value)):
            for k, v in value[i].items():
                data[f'skills_progression_{key}_{i}_{k}'] = v
    del data['skill_progression']
    return data

def convert_to_dataframe(data: dict) -> pd.DataFrame:
    new_df = pd.DataFrame({k: [str(v)] for k, v in data.items()})
    return new_df

def push_to_google_sheet(sheet: gspread.Worksheet, data: pd.DataFrame) -> None:
    sheet.append_rows(data.values.tolist())

jobtitles = ['Software Engineer', 'Data Analyst', 'Product Manager', 'UX Designer', 'Digital Marketer']
count = 0
for skill in jobtitles:
    print("Started: Profession:", skill)
    start = time.time()
    response = skills_openai(skill)
    content, prompt_tokens, completion_tokens = response
    content = json.loads(content)
    
    print(json.dumps(content, indent=4))
    print("Prompt tokens:", prompt_tokens)
    print("Completion tokens:", completion_tokens)
    
    output_dict = process_skill_progression(content)
    output_dict = flatten_dict(output_dict)
    
    df = convert_to_dataframe(output_dict)

    # Read the output.csv into out_df
    out_df = pd.read_csv("output.csv")

    # Concatenate both DataFrames
    out_df_combined = pd.concat([out_df, df], ignore_index=True, sort=True)

    # Reorder the columns to match the column order of out_df
    out_df_combined = out_df_combined[out_df.columns.tolist()]

    # Fill missing values with None
    out_df_combined = out_df_combined.where(pd.notnull(out_df_combined), None)
    
    sheet = connect_to_google_sheets_docs()
    
    push_to_google_sheet(sheet, out_df_combined)
    
    # append to csv as well
    if count == 0:
        out_df_combined.to_csv("skills.csv", index=False)
    out_df_combined.to_csv("skills.csv", mode='a', header=False, index=False)
    count += 1
    
    print("Data has been pushed successfully")   
    end = time.time()
    print("Time taken:", round(end - start, 2), "seconds")
        
    
    