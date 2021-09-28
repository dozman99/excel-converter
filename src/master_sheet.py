from openpyxl import load_workbook
import uuid, re

wb = load_workbook('static/excel/ENG301.1.xlsx', data_only=True)

ws = wb.active

anchor = None
topLeft = [0, 0]
bottomRight = [0, 0]

headerMap = []
data_rows = []

def is_Anchor(value):
    return re.search('matric', str(value).lower()) != None

def map_Header(value):
    if re.search('matric', str(value).lower()) != None:
        return 'mat_no'
    elif re.search('(total|score)', str(value).lower()) != None:
        return 'score'
    return 'annotation'

def getBatchId():
    return rnd_id()

def rnd_id():
    return uuid.uuid4().hex

def getResultId(mat_no):
    return str(session) + '-' + courseId + '-' + mat_no.replace('/', '-')

def hasRows():
    return bottomRight[0] > topLeft[0]

def parse_header():
    global anchor, session, courseCode
    for r in range(1, 21):
        for c in range(1, 21):
            cell = ws.cell(r, c).value
            if is_Anchor(cell):
                anchor = [r, c]
                break
            elif re.search('session', str(cell).lower()) != None:
                s = peek_right('^(_){0,1}((\d){2}|(\d){4})/((\d){2}|(\d){4})(_){0,1}$', r, c)
                y2 = s.strip().replace('_','').split('/')[1]
                if len(y2) == 2:
                    y2 = '20' + y2
                session = int(y2)
            elif re.search('course (code|no)', str(cell).lower()) != None:
                code = peek_right('^(_){0,1}[A-z]{3}(\s|_){0,1}\d{3}\.\d(_){0,1}$', r, c)
                c1 = code.replace('_', '').replace(' ', '').lower()
                c2 = re.split('([a-z]{3})(\d{3})\.(\d)', c1)
                courseCode = c2[1] + '_' + c2[2] + '_' + c2[3]
            # TODO parse other headers
        if anchor != None:
            break

def peek_right(pattern, row, column):
    for c in range(column + 1, column + 4):
        cell = ws.cell(row, c).value
        if re.search(pattern, str(cell).strip()) != None:
            return cell

def parse_sheet():
    parse_header()
    if anchor != None:
        go_left()
        go_right()
        go_down()
        map_Headers()
        if hasRows():
            getRows()

def getRows():
    for ur in range(topLeft[0] + 1, bottomRight[0] + 1):
        row = {
            'batchId': batchId, 'session': session,
            'courseId': courseId, 'courseCode': courseCode,
            'mat_no': '', 'score': ''
            }
        annotation = []
        for uc in range(topLeft[1], bottomRight[1] + 1):
            if headerMap[uc - topLeft[1]] != 'annotation':
                row[headerMap[uc - topLeft[1]]] = sanitize(ws.cell(ur, uc).value)
            else:
                annotation.append({
                    'key': sanitize(ws.cell(anchor[0], uc).value),
                    'value': sanitize(ws.cell(ur, uc).value)
                })
        row['annotation'] = str(annotation)
        row['mat_no'] = row['mat_no'].upper()
        # TODO tests and sanitize (mat number, score), yada yada yada!

        row['resultId'] = getResultId(row['mat_no'])
        data_rows.append(row)

def sanitize(value):
    if value == None:
        return ''
    elif type(value) == str:
        return value.strip()
    else:
        return value

def map_Headers():
    for h in range(topLeft[1], bottomRight[1] + 1):
        headerMap.append(map_Header(ws.cell(anchor[0], h).value))
        
def go_left():
    topLeft[1] = anchor[1]
    for l in range(anchor[1] -1, 0, -1):
        if ws.cell(anchor[0], l).value == None:
            break
        else:
            topLeft[1] = l
def go_right():
    bottomRight[1] = anchor[1]
    while True:
        bottomRight[1] = bottomRight[1] + 1
        if ws.cell(anchor[0], bottomRight[1]).value == None:
            bottomRight[1] = bottomRight[1] - 1
            break
def go_down():
    topLeft[0] = anchor[0]
    bottomRight[0] = anchor[0]
    while True:
        bottomRight[0] = bottomRight[0] + 1
        if ws.cell(bottomRight[0], anchor[1]).value == None:
            bottomRight[0] = bottomRight[0] - 1
            break

def write_to_database():
    # TODO initiate database transaction
    for data in data_rows:
        print(data)

def parse_file():
    global ws, anchor
    for sheet in wb.worksheets:
        ws = sheet
        anchor = None
        parse_sheet()
    if len(data_rows) > 0:
        write_to_database()

batchId = getBatchId()
session = 2019
courseCode = 'chm_130_1'
courseId = rnd_id()

parse_file()
