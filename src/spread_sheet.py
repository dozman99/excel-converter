from openpyxl import load_workbook
from openpyxl.comments import Comment
from sample_result import user, results
import courses

wb = load_workbook('static/excel/spreadsheet_template.xlsx')

# sort only by session
results.sort(key = lambda i: (i['session']))

def get_spread_sheet():
    user['name'] = (user['last_name'].upper() + ', ' + 
        user['first_name'].capitalize() + ' ' + user['other_name'].capitalize())
    for key in user.keys():
        if sheetMap.get(key) != None:
            wb['L100'][sheetMap[key]] = user[key]

    level_status = { 'sessions': [0], 'last_sem': 101 }
    result_map = {}
    courseMap = courses.MEG

    if user['department'] == 'MCT':
        courseMap = courses.MCT
    # step 1: remove carry-overs
    for result in results:
        result.update(courseMap[result['courseCode']])
        result.update({'_session': result['session'], 'comment': ''})
        map = result_map.get(result['courseCode'])
        if  level_status['sessions'][-1] != result['session']:
            level_status['sessions'].append(result['session'])
        if result['level'] + result['sem'] > level_status['last_sem']:
            level_status['last_sem'] = result['level'] + result['sem']
        if result['score'] < 40:
            result['cu'] = 0
        if map == None or (map['score'] < 40 and result['session'] > map['session']):
            if map != None:
                result['comment'] = (map['comment'] + '[ session: ' + str(map['session'] - 1) + '/' 
                    + str(map['session']) + ', score: ' + str(map['score']) + ']\n')
                result['_session'] = result['session'] + 0.4
            result_map[result['courseCode']] = result
        else:
            result_map[result['courseCode']]['comment'] = (map['comment'] + 'flag* [ session: ' + str(result['session'] - 1) + '/' 
                + str(result['session']) + ', score: ' + str(result['score']) + ']\n')

    # step 2: add missing courses till last semester    
    for key in courseMap.keys():
        if result_map.get(key) == None and courseMap[key]['level'] + courseMap[key]['sem'] <= level_status['last_sem']:
            result_map[key] = courseMap[key]
            session = level_status['sessions'][int(courseMap[key]['level']/100)]
            result_map[key].update({'courseCode': key, 'cu': None, '_session': session + 0.5, 'session': session })

    # step 3: write the session to sheet
    for i in range(1, len(level_status['sessions'])):
        wb['L' + str(i * 100)][sheetMap['session']] = str(level_status['sessions'][i] - 1) + '/' + str(level_status['sessions'][i])

    final_results = []
    final_results.extend(result_map.values())
    final_results.sort(key = lambda i: (i['_session'], i['code']))

    # step 4: write reuluts to sheet
    write_results(final_results, level_status)
    
    # step 5: remove unused sheets
    for sheet in wb.worksheets:
        if sheet[sheetMap['session']].value == None:
            wb.remove(sheet)
    wb.save('output/testing.xlsx')
    return

def write_results(results, status):
    sems = {}
    for result in results:
        level = status['sessions'].index(result['session']) * 100
        sem_id = str(level + result['sem'])
        if sems.get(sem_id) == None:
            sems[sem_id] = 0
        i = sems[sem_id]
        if i >= 14:
            continue
        ws = wb['L' + str(level)]
        ws['A' + str(i + [11, 28][result['sem'] - 1])] = result['code']
        ws['B' + str(i + [11, 28][result['sem'] - 1])] = result['title']
        ws['C' + str(i + [11, 28][result['sem'] - 1])] = result['cu']
        ws['D' + str(i + [11, 28][result['sem'] - 1])] = result.get('score')
        if result.get('comment') != '' and result.get('comment') != None:
            ws.cell(i + [11, 28][result['sem'] - 1], 4).comment = Comment('Revisions:\n' + result.get('comment'), 'Auto', width=300)
        sems[sem_id] += 1
    return

sheetMap = {
    'name': 'B6',
    'soo': 'B7',
    'mat_no': 'D6',
    'sex': 'G7',
    'marital': 'E7',
    'session': 'C9',
    'hod': 'B46',
    'dept': 'A3',
    'faculty': 'A2'
}

get_spread_sheet()
