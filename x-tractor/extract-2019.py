import docx2txt
import json

test_dir = r'C:\Users\DavyBellens\OneDrive - VLIR-UOS\Documenten\vliruos-scripts\x-tractor'
test_file = test_dir+r'\APR BO2019TEA480A103_Y4_V2 final.docx'

def extract_text(file_path):
    text = docx2txt.process(file_path)
    return text

headings = [
    'General Information\n',
    'Project progress (max. 5 pages)\n',
    'Risk management\n',
    'Synergy and complementarity\n',
    'Lessons Learned\n',
    'Annexes\n'
]

general_information_headings = [
    'Type:\n',
    'TEAM / SI / JOINT\n',
    'Country and region (within the country) of the project :\n',
    'Summary of progress made: Summarize the most important achievements of project during the reporting year (Max. 10 lines)\n',
    'Project start:\n',
    'End of project:\n',
    'Reporting year:\n',
    'Project team:\n',
    'Local promoter (Name, institution)\n',
    'Local co-promoter(s) (Name, institution)\n',
    'Flemish promoter (Name, institution)\n',
    'Flemish co-promoter(s) (Name, institution)\n',
    '… add rows for other team members/personnel involved in the project\n',
]

project_progress_headings = {
    'Progress made towards the intermediate results\n': [
        "Progress of indicators \n",
        "Progress made towards intermediate results\n"
    ],
    'Progress made towards the specific objective(s)\n': [
        "Progress of indicators \n",
        "Progress made towards specific objective(s)\n"
    ],
}



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

def get_substring(text: str, start: int, end: int = None):
    start_index = text.find(general_information_headings[start])+len(general_information_headings[start])
    result = text[start_index:]
    if end:
        end_index = text.find(general_information_headings[end])
        result = text[start_index:end_index]
    result = result.replace('\n\n', '\n')
    return result.strip().strip('\n')
    
def process_general_information(text: str):
    general_information = {}

    project_title = get_substring(text, 0, 1)
    project_type = get_substring(text, 1, 2)
    country = get_substring(text, 2, 3)
    summary = get_substring(text, 3, 4)
    project_dates = get_substring(text, 6, 7)
    local_promoter = get_substring(text, 8, 9)
    local_co_promoter = get_substring(text, 9, 10)
    flemish_promoter = get_substring(text, 10, 11)
    flemish_co_promoter = get_substring(text, 11, 12)
    other_team_members = get_substring(text, 12)

    general_information['Title'] = project_title
    general_information['Type'] = project_type
    general_information['Country and region (within the country) of the project :'] = country
    general_information['Summary of progress made: Summarize the most important achievements of project during the reporting year (Max. 10 lines)'] = summary
    general_information['Project start:'] = project_dates.split('\n')[0]
    general_information['End of project:'] = project_dates.split('\n')[1]
    general_information['Reporting year:'] = project_dates.split('\n')[2]
    general_information['Project team:'] = {
        'Local promoter (Name, institution)': local_promoter,
        'Local co-promoter(s) (Name, institution)': local_co_promoter,
        'Flemish promoter (Name, institution)': flemish_promoter,
        'Flemish co-promoter(s) (Name, institution)': flemish_co_promoter,
        '… add rows for other team members/personnel involved in the project': other_team_members
    }
    return general_information
    

def process_project_progress(text: str):
    project_progress = {}

    for heading in project_progress_headings:
        project_progress[heading.removesuffix('\n')] = {}
        for subheading in project_progress_headings[heading]:
            start_index = text.find(subheading)+len(subheading)
            end_index = text.find(subheading, start_index+1)
            if end_index == -1:
                end_index = len(text)
            project_progress[heading.removesuffix('\n')][subheading.removesuffix('\n')] = text[start_index:end_index].replace('\n\n', '\n').strip().strip('\n')
    return project_progress

t = extract_text(test_file)
sections = extract_sections(t)

def process(index):
    if index == 0:
        return process_general_information(sections[headings[index]])
    elif index == 1:
        return process_project_progress(sections[headings[index]])

def main():
    apr = {}
    
    for i in range(len(headings)):
        apr[headings[i].removesuffix('\n')] = process(i)

    with open('data-2019.json', 'w', encoding='utf-8') as f:
        json.dump(apr, f, ensure_ascii=False, indent=4)

main()