import shutil
import docx2txt
import docx
from docx.text.hyperlink import Hyperlink
import json
import os
import re
import pandas as pd

dir = r'C:\Users\DavyBellens\VLIR-UOS\Digital Factory - APR-indicatoren 2022-2023\TEAM'
output = r'C:\Apps\vliruos-scripts\x-tractor\output'
tmp = output + '\\tmp'
files = {}

if not os.path.exists(tmp):
    os.makedirs(tmp)

for i in os.listdir(dir):
    base = f"{dir}\\{i}\\"
    for j in os.listdir(base):
        if j.endswith('.docx') and "2019" not in j:
            files[i] = base+j

headings = [
    'General information\n',
    'Project strategy\n',
    'Organisation\n',
    'Stakeholder management and coherence\n',
    'Planning and budgeting\n',
    'Monitoring and evaluation\n',
    'Learning and steering\n',
    'Feel free to attach pictures, stories, blogs, articles in the written press, etc. of or about your project.\n'
]

general_information_headings = [
    'Project code:\n',
    'Project title:\n',
    'Summary of progress made: \n'
]

project_strategy_headings = [
    "What have you achieved? Explain how your project is on its way to reach the outcomes that were described in your proposal.\n\nFor final year reporting: Elaborate to what extent the outcomes have been achieved. Describe how the project has contributed to the LNOB (incl. gender equality), environmental sustainability, and interconnectedness principles.\n",
    "What have you done? Report on the progress you have made or setbacks you have encountered regarding the intermediate changes (and the anticipated results) set out in your proposal. Explain the key results of the past year, referring to the underlying deliverables. \n\n\nFor final year reporting: please include a retrospective on the entire project period.\n"
]

organisation_headings = [] # to indicate its empty

stakeholder_management_and_coherence_headings = [
    'Describe how you have engaged with key stakeholders during the past year. What was the result?\n',
    'Explain how you have collaborated with other organisations? What was the result? Describe active linkages with other (Belgian) development actors, and other VLIR-UOS funded activities.\n'
]

planning_and_budgeting_headings = [] # to indicate its empty

monitoring_and_evaluation_headings = [] # to indicate its empty

learning_and_steering_headings = [
    "Learning - Share the lessons learned that you have acquired in the past year. What advice would you give to researchers who propose a similar project?\n",
    "Steering"
]

steering_subheadings = [
    "A – Modification of the budget lines “investment or personnel costs”, exceeding more than 30% and more than 10,000 EUR as compared to the respective initial budget line total for the whole project period, needs to be validated by VLIR-UOS.\n",
    "B – Other important reorientations of the project: important changes in the management / organisation of the project; Major changes in activities / planning that have important consequences for the project design and progress.\n",
]

def extract_text(file_path):
    images_before = len(os.listdir(tmp))
    file = docx2txt.process(file_path, tmp)
    file_images = os.listdir(tmp)[images_before:]
    return file, file_images

def clean_text(text: str) -> str:
    # This regex will match any line starting with "Indicative max." until the first '\n'
    cleaned_text = re.sub(r'Indicative max\..*?\n', '', text)
    # Strip leading and trailing whitespace and extra newlines
    return cleaned_text.strip().strip('\n').strip('\n')

def extract_sections(text: str):
    sections = {}
    for i in range(len(headings)):
        start = text.find(headings[i])+len(headings[i])
        if i == len(headings) - 1:
            end = len(text)
        else:
            end = text.find(headings[i + 1])
        sections[headings[i]] = text[start:end]
    return sections

def process_general_information(text: str, project_code: str, general_information: dict):
    general_information['Project code'] = project_code
    general_information['Project title'] = clean_text(text[text.find('Project title:')+len('Project title:'):text.find('Summary of progress made:')].strip().strip('\n').strip(project_code))
    general_information['Summary of progress made'] = clean_text(text[text.find(general_information_headings[2])+len(general_information_headings[2]):])

    return general_information

def process_project_strategy(text: str, project_strategy: dict):

    project_strategy['What have you achieved'] = clean_text(text[text.find(project_strategy_headings[0])+len(project_strategy_headings[0]):text.find(project_strategy_headings[1])])
    project_strategy["What have you done"] = clean_text(text[text.find(project_strategy_headings[1])+len(project_strategy_headings[1]):])

    return project_strategy

def process_no_subheadings(text: str):
    return clean_text(text)
    

def process_stakeholder_management_and_coherence(text: str, smc: dict):

    smc["Engagement with key stakeholders"] = clean_text(text[text.find(stakeholder_management_and_coherence_headings[0])+len(stakeholder_management_and_coherence_headings[0]):text.find(stakeholder_management_and_coherence_headings[1])])
    smc["Collaboration with other organisations"] = clean_text(text[text.find(stakeholder_management_and_coherence_headings[1])+len(stakeholder_management_and_coherence_headings[1]):])

    return smc           

def process_learning_and_steering(text: str, ls: dict):

    ls["Learning"] = clean_text(text[text.find(learning_and_steering_headings[0])+len(learning_and_steering_headings[0]):text.find(learning_and_steering_headings[1])])

    ls["Steering-A"] = clean_text(text[text.find("A –")+239:text.find(steering_subheadings[1])])
    ls["Steering-B"] = clean_text(text[text.find(steering_subheadings[1])+len(steering_subheadings[1]):])

    return ls

def process_attachment(text: str):
    return clean_text(text[:text.find('Annual Progress Report')])



def process(index, sections, filename, apr):
    if index == 0:
        return process_general_information(sections[headings[index]], filename, apr)
    elif index == 1:
        return process_project_strategy(sections[headings[index]], apr)
    elif index == 2:
        return apr
    elif index == 3:
        return process_stakeholder_management_and_coherence(sections[headings[index]], apr)
    elif index == 4:
        return apr
    elif index == 5:
        return apr
    elif index == 6:
        return process_learning_and_steering(sections[headings[index]], apr)
    elif index == 7:
        attachments = process_attachment(sections[headings[index]])
        if 'http' in attachments:
            return attachments
        else:
            t = docx.Document(files[filename])
            for item in t.paragraphs[-1].iter_inner_content():
                if isinstance(item, Hyperlink):
                    return item.url

old = {}

for j in files.keys():
    apr = {}
    extracted_text, images = extract_text(files[j])
    sections = extract_sections(extracted_text)

    for i in range(len(headings)):
        if i == len(headings) - 1:
            apr["Attachments"] = process(i, sections, j, apr)
        else:  
            apr = process(i, sections, j, apr)
            
    apr['Pictures'] = images
    old[str(j)] = apr
    
    with open(output+'\\data-2022-TEAM.json', 'w', encoding='utf-8') as f:
        json.dump(old, f, indent=4, ensure_ascii=False)

# zip the image directory


shutil.make_archive(tmp, 'zip', tmp)

# remove the temporary image directory

shutil.rmtree(tmp)


json_file = output+'\\data-2022-TEAM.json'
output_file = output+'\\data-2022-TEAM.xlsx'

df = pd.DataFrame(old)
df = df.transpose()

os.remove(output_file)

df.to_excel(output_file, index=False)