from openpyxl import load_workbook

wb = load_workbook('static/excel/spreadsheet_template.xlsx')


results = [
    { 'courseCode':'ges_100_1', 'session': 2017, 'score': 59 },
    { 'courseCode':'chm_130_1', 'session': 2016, 'score': 66 },
    { 'courseCode':'chm_131_2', 'session': 2016, 'score': 57 },
    { 'courseCode':'phy_216_1', 'session': 2017, 'score': 71 },
    { 'courseCode':'ges_100_1', 'session': 2016, 'score': 34 },
    { 'courseCode':'eng_201_1', 'session': 2017, 'score': 43 },
    { 'courseCode':'eng_103_2', 'session': 2016, 'score': 49 },
]

results.sort(key = lambda i: (i['session']))

user = {
    'first_name': 'John',
    'last_name': 'Doe',
    'other_name': '',
    'soo': 'Rivers',
    'mat_no': 'U2015/3025102',
    'sex': 'M',
    'marital': 'Single',
    'department': 'MEG'
}

def get_spread_sheet():
    user['name'] = (user['last_name'].upper() + ', ' + 
        user['first_name'].capitalize() + ' ' + user['other_name'].capitalize())
    for key in user.keys():
        if sheetMap.get(key) != None:
            wb['L100'][sheetMap[key]] = user[key]
    for result in results:
        courseMap = courseMapMEG
        if department == 'MCT':
            courseMap = courseMapMCT
        map = courseMap[result['courseCode']]
        wb[map['level']][map['cell']] = result['score']
        if wb[map['level']][sheetMap['session']].value == None:
            wb[map['level']][sheetMap['session']] = str(result['session'] - 1) + '/' + str(result['session'])
    wb.save('output/testing.xlsx')
    return

department = 'MEG'

sheetMap = {
    'name': 'C6',
    'soo': 'C7',
    'mat_no': 'F6',
    'sex': 'I7',
    'marital': 'G7',
    'session': 'E9'
}

courseMapMEG = {
    'chm_130_1': {
        'level': '100',
        'cell': 'F11'
    },
    'ges_100_1': {
        'level': '100',
        'cell': 'F12'
    },
    'ges_102_1': {
        'level': '100',
        'cell': 'F13'
    },
    'chm_131_2': {
        'level': '100',
        'cell': 'F18'
    },
    'eng_102_2': {
        'level': '100',
        'cell': 'F19'
    },
    'eng_103_2': {
        'level': '100',
        'cell': 'F20'
    },


    'phy_216_1': {
        'level': '200',
        'cell': 'F11'
    },
    'eng_201_1': {
        'level': '200',
        'cell': 'F12'
    },
    'eng_202_1': {
        'level': '200',
        'cell': 'F13'
    },
    'chm_240_2': {
        'level': '200',
        'cell': 'F18'
    },
    'eng_206_2': {
        'level': '200',
        'cell': 'F19'
    },
    'eng_207_2': {
        'level': '200',
        'cell': 'F20'
    }
}

courseMapMCT = {
    'chm_130_1': {
        'level': '100',
        'cell': 'F11'
    }
}

get_spread_sheet()
