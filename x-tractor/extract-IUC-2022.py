import shutil
import docx2txt
import os
import json
import re
import pandas as pd

cd = r'C:\Apps\vliruos-scripts\x-tractor'
output = cd + '\\output'
tmp = output + '\\tmp'
zipped_images = output + '\\images.zip'

if not os.path.exists(tmp):
    os.makedirs(tmp)

if os.path.exists(zipped_images):
    os.remove(zipped_images)

headings = {
    "Project level APR": [
        "General Information",
        "Context and problem analysis",
        "Project strategy",
        "Organisation",
        "Stakeholder management and coherence",
        "Planning and Budgeting",
        "Monitoring and evaluation",
        "Learning and steering",
    ],
    "Project support": [
        "General Information",
        "Project strategy",
        "Planning and budgeting",
        "Monitoring and evaluation",
        "Learning and steering",
    ]
}

general_information_fields = [
    "Project title:",
    "Project code:",
    "Flemish institution:",
    "Partner institution:",
    "Flemish Coordinator:",
    "Partner Coordinator:",
    "Project Manager:",
    "ICOS:",
    "Summary of progress made:",
    "Sub-project n°:",
]

context_and_problem_analysis_fields = [
    "Describe relevant developments / changes at the institutional level. Please explain what the consequences are for your project.\n",
    "Describe relevant changes in the context of the country or region. Please explain what the consequences are for your project.\n",
]

project_strategy_fields = [
    "What have you achieved? Explain how your project is on its way to reach the outcomes that were described in your proposal.\n\n\nFor final year reporting in IUC Phase 1 or Phase 2: Elaborate to what extent the outcomes have been achieved. Describe how the project has contributed to the LNOB (incl. gender equality), environmental sustainability, and interconnectedness principles.\n",
]

# empty to show that there are no fields in the organisation section
organisation_fields = []

stakeholder_management_and_coherence_fields = [
    "Describe how you have engaged with stakeholders to create the effective uptake of project results outside of the academic sector? Did the involvement of key stakeholders lead to any noticeable changes?\n",
    "Explain how you have collaborated with other organisations? Describe active linkages with other (Belgian) development actors, and other VLIR-UOS funded activities. What was the result?\n",
]

# empty to show that there are no fields in the planning and budgeting section
planning_and_budgeting_fields = []

# empty to show that there are no fields in the monitoring and evaluation section
monitoring_and_evaluation_fields = []

learning_and_steering_fields = [
    "Learning - Share the lessons learned that you have acquired in the past year. What advice would you give to researchers who propose a similar project?\n",
    "Steering – Please describe major reorientations of the project that have been implemented in the reporting year, or planned in the next year. Please do not mention minor changes.\n"
]

subproject_general_info_fields = [
    "Sub-project title:",
    "Sub-project type:",
    "Flemish institution:",
    "Partner institution:",
    "Flemish Team leader:",
    "Partner Team leader:",
    "Summary of progress made:",
]

subproject_project_strategy_fields = [
    "What have you achieved? Explain how your sub-project is on its way to reach the outcomes that were described in your proposal.\n\n\nFor final year reporting in IUC Phase 1 or Phase 2: Elaborate to what extent the outcomes have been achieved. Describe how the project has contributed to the LNOB (incl. gender equality), environmental sustainability, and interconnectedness principles.\n",
    "What have you done? Report on the progress you have made or setbacks you have encountered regarding the intermediate changes (and the anticipated results) set out in your proposal. Explain the key results of the past year, referring to the underlying deliverables.\n",
]

# empty to show that there are no fields in the subproject planning and budgeting section
subproject_planning_and_budgeting_fields = []

# empty to show that there are no fields in the subproject monitoring and evaluation section
subproject_monitoring_and_evaluation_fields = []

subproject_learning_and_steering_fields = [
    "Learning - Share the lessons learned that you have acquired in the past year. What advice would you give to researchers who propose a similar project?\n",
    "Steering – Please describe major reorientations of the sub-project that have been implemented in the reporting year, or planned in the next year. Please do not mention minor changes. Important reorientations of the sub-project: important changes in the management / organisation of the sub-project / changes in the sub-project team (e.g. new team members); Major changes in activities / planning that have important consequences for the sub-project design and progress.\n"
]


def clean_text(text: str) -> str:
    cleaned_text = re.sub(r'Indicative max\..*?\n', '', text)
    cleaned_text = re.sub(r'\n+\d\.\d$', '', cleaned_text)
    return cleaned_text.strip().strip('\n')


def extract_text(file_path):
    images_before = len(os.listdir(tmp))
    file = docx2txt.process(file_path, tmp)
    file_images = os.listdir(tmp)[images_before:] or []
    return clean_text(file), file_images


def extract_sections(text: str) -> dict:
    sections = {}
    project_level_apr, project_support = text.split(
        "Feel free to attach pictures, stories, blogs, articles in the written press, etc. of or about your project.")
    project_level_apr = project_level_apr[project_level_apr.find(
        "General Information"):]
    for section in range(len(headings["Project level APR"])-1):
        h = headings["Project level APR"]
        hs = h[section]
        if hs in project_level_apr:
            res = clean_text(project_level_apr[project_level_apr.find(
                hs)+len(hs):project_level_apr.find(h[section+1])])
            sections[hs] = res
    sections["Learning and steering"] = clean_text(
        project_level_apr[project_level_apr.find("Learning and steering"):])

    project_support = project_support[project_support.find(
        "General Information"):]
    try:
        for index, sp in enumerate(project_support.split("General Information")):
            if index == 0:
                continue
            res = {}
            res["Subproject " + str(index)] = clean_text(sp)
            sections["Subproject " + str(index)] = res
    except Exception as e:
        print(e)

    del sections["Organisation"]
    del sections["Planning and Budgeting"]
    del sections["Monitoring and evaluation"]

    return sections


def process_general_information(text: str, data: dict):
    for i in range(len(general_information_fields)-1):
        field = general_information_fields[i]
        next_field = general_information_fields[i+1]
        data[field] = clean_text(
            text[text.find(field)+len(field):text.find(next_field)])
        if general_information_fields[i] == "Project code:":
            data["Project code:"] = data["Project code:"].split(' ')[
                0]
    return data


def process_context_and_problem_analysis(text: str, data: dict):
    data["Developments at institutional level"] = clean_text(text[text.find(
        context_and_problem_analysis_fields[0])+len(context_and_problem_analysis_fields[0]):text.find(context_and_problem_analysis_fields[1])])
    data["Changes in context of country or region"] = clean_text(
        text[text.find(context_and_problem_analysis_fields[1])+len(context_and_problem_analysis_fields[1]):])
    return data


def process_project_strategy(text: str, data: dict):
    data["Achievements"] = clean_text(
        text[text.find(project_strategy_fields[0])+len(project_strategy_fields[0]):])
    return data


def process_stakeholder_management_and_coherence(text: str, data: dict):
    data["Stakeholder engagement"] = clean_text(text[text.find(stakeholder_management_and_coherence_fields[0])+len(
        stakeholder_management_and_coherence_fields[0]):text.find(stakeholder_management_and_coherence_fields[1])])
    data["Collaboration with other organisations"] = clean_text(text[text.find(
        stakeholder_management_and_coherence_fields[1])+len(stakeholder_management_and_coherence_fields[1]):])
    return data


def process_learning_and_steering(text: str, data: dict):
    data["Learning"] = clean_text(text[text.find(
        learning_and_steering_fields[0])+len(learning_and_steering_fields[0]):text.find(learning_and_steering_fields[1])])
    data["Steering"] = clean_text(text[text.find(
        learning_and_steering_fields[1])+len(learning_and_steering_fields[1]):])
    return data


def process_subproject_general_info(text: str, data: dict):
    for i in range(len(subproject_general_info_fields)-1):
        field = subproject_general_info_fields[i]
        next_field = subproject_general_info_fields[i+1]
        data[field] = clean_text(
            text[text.find(field)+len(field):text.find(next_field)])
    return data


files_dir = r'C:\Users\DavyBellens\VLIR-UOS\Digital Factory - APR-indicatoren 2022-2023\IUC'

files = [f"{files_dir}\\{i}\\{j}" for i in os.listdir(files_dir) if os.path.isdir(files_dir+'\\'+i)
         for j in os.listdir(files_dir+'\\'+i) if j.endswith('.docx')]

data = {}

for index, file in enumerate(files):
    try:
        text, images = extract_text(file)
        sections = extract_sections(text)
        result = {}
        result = process_general_information(
            sections["General Information"], result)
        del sections["General Information"]
        result = process_context_and_problem_analysis(
            sections["Context and problem analysis"], result)
        del sections["Context and problem analysis"]
        result = process_project_strategy(
            sections["Project strategy"], result)
        del sections["Project strategy"]
        result = process_stakeholder_management_and_coherence(
            sections["Stakeholder management and coherence"], result)
        del sections["Stakeholder management and coherence"]
        result = process_learning_and_steering(
            sections["Learning and steering"], result)
        del sections["Learning and steering"]

        for sub in range(len(sections)):
            subproject = sections["Subproject " +
                                  str(sub+1)]["Subproject " + str(sub+1)]
            
            for i in range(len(subproject_general_info_fields)-1):
                field = subproject_general_info_fields[i]
                next_field = subproject_general_info_fields[i+1]
                result[f"SP-{sub+1} {field}"] = clean_text(
                    subproject[subproject.find(field)+len(field):subproject.find(next_field)])
            
            result[f"SP-{sub+1} Achievements"] = clean_text(
                subproject[subproject.find(subproject_project_strategy_fields[0])+len(subproject_project_strategy_fields[0]):subproject.find(subproject_project_strategy_fields[1])])
            
            result[f"SP-{sub+1} Progress"] = clean_text(
                subproject[subproject.find(subproject_project_strategy_fields[1])+len(subproject_project_strategy_fields[1]):subproject.find(subproject_learning_and_steering_fields[0])])
            
            result[f"SP-{sub+1} Learning"] = clean_text(
                subproject[subproject.find(subproject_learning_and_steering_fields[0])+len(subproject_learning_and_steering_fields[0]):subproject.find(subproject_learning_and_steering_fields[1])])
            
            result[f"SP-{sub+1} Steering"] = clean_text(
                subproject[subproject.find(subproject_learning_and_steering_fields[1])+len(subproject_learning_and_steering_fields[1]):])

        data[result["Project code:"]] = result

    except Exception as e:
        print(e)
        data[file.split('\\')[-2]] = {"Error": "Different format"}

with open(cd + '\\output\\data-2022-IUC.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, indent=4, ensure_ascii=False)

# zip the image directory


shutil.make_archive(cd + '\\output\\images', 'zip', tmp)

# remove the temporary image directory

shutil.rmtree(tmp)

# convert the json to a csv file

df = pd.read_json(cd + '\\output\\data-2022-IUC.json').T
df.to_excel(cd + '\\output\\data-2022-IUC.xlsx', index=False)
