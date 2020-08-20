from googleapiclient.discovery import build # Builds the call
from googleapiclient.errors import HttpError # Catches HTTP errors
from google.oauth2 import service_account # Google Authentication
import os # OS-independent way of getting directory paths
import csv
from collections import OrderedDict


def main(spreadsheetId, tour_date, last_tour_date):
    spreadsheet_id = '1Msq8pgWFj83DwLumVdgk84fSmBG_Uq3UcaJ7Ro_mk6Q'

    creds = service_account.Credentials.from_service_account_file(os.path.join(os.path.dirname(__file__)) + '/credentials.json', scopes=['https://www.googleapis.com/auth/spreadsheets'])
    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    result = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    properties = [sheet.get('properties') for sheet in result.get('sheets', [])]
    ids = [property.get('sheetId', 0) for property in properties]
    titles = [property.get('title', '') for property in properties]
    test_sheet = titles.index(tour_date) if tour_date in titles else 0

    if test_sheet:
        titles.pop(test_sheet)
        body = {
            'requests': [{
                'deleteSheet': {
                    "sheetId": ids.pop(test_sheet)
                }
            }]
        }
        result = service .spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range="'Tournament Turnout'!A1:Z", valueRenderOption='FORMULA').execute()
    values = result.get('values', [])
    last_turnout_index = len(values)
    columnA = tour_date
    columnB = values[-1][1][:2] + tour_date + values[-1][1][7:]
    columnC = values[-1][2][:10] + tour_date + values[-1][2][15:]

    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range="'" + last_tour_date + "'!A1:Z", valueRenderOption='FORMULA').execute()
    values = result.get('values', [])
    columnD = values[2][3].replace('""', '"')
    new_tour_columnG = values[2][6][:23] + last_tour_date+ values[2][6][28:]
    columnJ = values[2][9]
    columnK = values[2][9]
    columnL = values[2][9]
    columnX = values[2][-3][:32] + last_tour_date + values[2][-3][37:-27] + titles[len(ids)-7] + values[2][-3][-22:]
    columnY = values[2][-2][:23] + last_tour_date + values[2][-2][28:]
    columnZ = values[2][-1][:41] + last_tour_date + values[2][-1][46:]

    for value in values:
        if value[0]:
            last_index = values.index(value)+1

    body = {
        'destination_spreadsheet_id': spreadsheet_id,
    }
    result = service.spreadsheets().sheets().copyTo(spreadsheetId=spreadsheet_id, sheetId=ids[len(ids)-7], body=body).execute()
    new_sheet_id = result.get('sheetId')

    tour_csv = ''
    new_last_index = 0
    with open(os.path.join(os.path.dirname(__file__)) + '/tournamentData.csv', 'r', newline='') as csv_file:
        tour_reader = csv.reader(csv_file, delimiter='@', quotechar='|')
        for row in tour_reader:
            tour_csv += '@'.join(row) + '\n'
            new_last_index += 1
    new_last_index += last_index

    body = {
        'requests': [
        {
            'updateSheetProperties': {
                "properties": {
                    "sheetId": new_sheet_id,
                    "index": len(ids)-6,
                    "title": tour_date
                },
                "fields": 'index,title'
            }
        },
        {
            'copyPaste': {
                "source": {
                    "sheetId": ids[len(ids)-5],
                    "startRowIndex": last_turnout_index-1,
                    "endRowIndex": last_turnout_index
                },
                "destination": {
                    "sheetId": ids[len(ids)-5],
                    "startRowIndex": last_turnout_index,
                    "endRowIndex": last_turnout_index+1
                },
                "pasteType": 'PASTE_NORMAL',
                "pasteOrientation": 'NORMAL'
            }
        },
        {
            'updateCells': {
                "rows": [{
                    "values":[{
                        "userEnteredValue": {
                            "stringValue": columnA
                        }
                    },
                    {
                        "userEnteredValue": {
                            "formulaValue": columnB
                        }
                    },
                    {
                        "userEnteredValue": {
                            "formulaValue": columnC
                        }
                    }]
                }],
                "fields": 'userEnteredValue.stringValue,userEnteredValue.formulaValue',
                "range": {
                    "sheetId": ids[len(ids)-5],
                    "startRowIndex": last_turnout_index,
                    "endRowIndex": last_turnout_index+1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 3
                }
            }
        },
        {
            'updateCells': {
                "rows": [{
                    "values": [{
                        "userEnteredValue": {
                            "formulaValue": columnD
                        }
                    }]
                }],
                "fields": 'userEnteredValue.formulaValue',
                "start": {
                    "sheetId": new_sheet_id,
                    "rowIndex": 2,
                    "columnIndex": 3
                }
            }
        },
        {
            'updateCells': {
                "rows": [{
                    "values": [{
                        "userEnteredValue": {
                            "formulaValue": new_tour_columnG
                        }
                    }]
                }],
                "fields": 'userEnteredValue.formulaValue',
                "start": {
                    "sheetId": new_sheet_id,
                    "rowIndex": 2,
                    "columnIndex": 6
                }
            }
        },
        {
            'updateCells': {
                "rows": [{
                    "values": [{
                        "userEnteredValue": {
                            "formulaValue": columnX
                        }
                    },{
                        "userEnteredValue": {
                            "formulaValue": columnY
                        }
                    },{
                        "userEnteredValue": {
                            "formulaValue": columnZ
                        }
                    }]
                }],
                "fields": 'userEnteredValue.formulaValue',
                "range": {
                    "sheetId": new_sheet_id,
                    "startRowIndex": 2,
                    "endRowIndex": 3,
                    "startColumnIndex": 23,
                    "endColumnIndex": 26
                }
            }
        },
        {
            'updateCells': {
                "rows": [{
                    "values": [{
                        "userEnteredValue": {
                            "stringValue": ''
                        }
                    }]
                }],
                "fields": 'userEnteredValue.stringValue',
                "range": {
                    "sheetId": new_sheet_id,
                    "startRowIndex": 2,
                    "endRowIndex": last_index,
                    "startColumnIndex": 2,
                    "endColumnIndex": 3
                }
            }
        },
        {
            'updateCells': {
                "rows": [{
                    "values": [{
                        "userEnteredValue": {
                            "stringValue": ''
                        }
                    }]
                }],
                "fields": 'userEnteredValue.stringValue',
                "range": {
                    "sheetId": new_sheet_id,
                    "startRowIndex": 2,
                    "endRowIndex": last_index,
                    "startColumnIndex": 12,
                    "endColumnIndex": 14
                }
            }
        },
        {
            'pasteData': {
                "coordinate": {
                    "sheetId": new_sheet_id,
                    "rowIndex": last_index,
                    "columnIndex": 0
                },
                "data": tour_csv,
                "type": 'PASTE_NORMAL',
                "delimiter": '@'
            }
        },
        {
            'copyPaste': {
                "source": {
                    "sheetId": new_sheet_id,
                    "startRowIndex": 2,
                    "endRowIndex": 3,
                    "startColumnIndex": 3,
                    "endColumnIndex": 12
                },
                "destination": {
                    "sheetId": new_sheet_id,
                    "startRowIndex": 3,
                    "startColumnIndex": 3,
                    "endColumnIndex": 12
                },
                "pasteType": 'PASTE_NORMAL',
                "pasteOrientation": 'NORMAL'
            }
        },
        {
            'copyPaste': {
                "source": {
                    "sheetId": new_sheet_id,
                    "startRowIndex": 2,
                    "endRowIndex": 3,
                    "startColumnIndex": 6,
                    "endColumnIndex": 7
                },
                "destination": {
                    "sheetId": new_sheet_id,
                    "startRowIndex": 3,
                    "startColumnIndex": 6,
                    "endColumnIndex": 7
                },
                "pasteType": 'PASTE_NORMAL',
                "pasteOrientation": 'NORMAL'
            }
        },
        {
            'copyPaste': {
                "source": {
                    "sheetId": new_sheet_id,
                    "startRowIndex": 2,
                    "endRowIndex": 3,
                    "startColumnIndex": 23,
                    "endColumnIndex": 26
                },
                "destination": {
                    "sheetId": new_sheet_id,
                    "startRowIndex": 3,
                    "startColumnIndex": 23,
                    "endColumnIndex": 26
                },
                "pasteType": 'PASTE_NORMAL',
                "pasteOrientation": 'NORMAL'
            }
        },
        {
            'updateCells': {
                "rows": [{
                    "values": [{
                        "userEnteredValue": {
                            "formulaValue": columnJ
                        }
                    },
                    {
                        "userEnteredValue": {
                            "formulaValue": columnK
                        }
                    },
                    {
                        "userEnteredValue": {
                            "formulaValue": columnL
                        }
                    }]
                }],
                "fields": 'userEnteredValue.formulaValue',
                "range": {
                    "sheetId": new_sheet_id,
                    "startRowIndex": 2,
                    "endRowIndex": 3,
                    "startColumnIndex": 9,
                    "endColumnIndex": 12
                    
                }
            }
        },
        {
            'updateCells': {
                "rows": [{
                    "values": [{
                        "userEnteredValue": {
                            "stringValue": ''
                        }
                    }]
                }],
                "fields": 'userEnteredValue.stringValue',
                "range": {
                    "sheetId": new_sheet_id,
                    "startRowIndex": new_last_index
                }
            }
        }]
    }
    result = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range="'" + tour_date + "'!A3:A", valueRenderOption='FORMULA').execute()
    names = [name[0] for name in result.get('values', [])]
    uniq_names = list(OrderedDict.fromkeys(names))
    dup_names = set()
    dup_names2 = set()
    for name in names:
        if name in dup_names:
            dup_names2.add(name)
        else:
            dup_names.add(name)

    name_indices = [names.index(name)+2 for name in dup_names2]

    deleteReqs = []
    prev_indices = []
    rev_i = len(name_indices)
    for i in range(len(name_indices)):
        index = name_indices[i]
        if i:
            for prev_index in prev_indices:
                if index > prev_index:
                    index -= 1
        deleteReqs.append(
            {
                'deleteRange': {
                    "range": {
                        "sheetId": new_sheet_id,
                        "startRowIndex": index,
                        "endRowIndex": index+1
                    },
                    'shiftDimension': 'ROWS'
                }
            }
        )
        prev_indices.append(index)
        if len(deleteReqs) >= 20:
            body = {
                'requests': deleteReqs
            }
            result = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
            deleteReqs = []
            print("20 Delete Requests processed")
        elif len(deleteReqs) >= 10 and len(deleteReqs) + rev_i <= 20:
            body = {
                'requests': deleteReqs
            }
            result = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
            deleteReqs = []
            print("10 Delete Requests processed")
        elif len(deleteReqs) + rev_i <= 10: 
            body = {
                'requests': deleteReqs
            }
            result = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
            deleteReqs = []
            print("Delete Request processed")
        rev_i -= 1

    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range="'Win Rates'!A1:Z", valueRenderOption='FORMULA').execute()
    values = result.get('values', [])
    columnD = values[1][3][:-1] + values[1][3][-51:-27] + tour_date + values[1][3][-22:]
    columnE = values[1][4][:-1] + values[1][4][-51:-27] + tour_date + values[1][4][-22:]
    columnG = values[1][6][:23] + tour_date + values[1][6][28:]

    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range="'Attendance'!A1:Z", valueRenderOption='FORMULA').execute()
    values = result.get('values', [])
    attend_column = values[1][3][:23] + '{}' + values[1][3][28:34] + '{}' + values[1][3][35:37] + '{}' + values[1][3][39:]

    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range="'Rankings'!A1:Z", valueRenderOption='FORMULA').execute()
    values = result.get('values', [])
    rank_columnJ = values[1][9] + values[1][9][-64:-36] + tour_date + values[1][9][-31:]
    rank_columnK = values[1][10] + values[1][10][-123:-91] + tour_date + values[1][10][-86:-37] + tour_date + values[1][10][-32:]

    update_names = []
    for name in uniq_names:
        update_names.append({
            "values":[{
                "userEnteredValue": {
                    "stringValue": str(name)
                }
            }]
        })

    body = {
        'requests': [{
            'updateCells': {
                "rows": [{
                    "values":[{
                        "userEnteredValue": {
                            "formulaValue": attend_column.format(tour_date, 'X', 24)
                        }
                    },
                    {
                        "userEnteredValue": {
                            "formulaValue": attend_column.format(tour_date, 'Y', 25)
                        }
                    }]
                }],
                "fields": 'userEnteredValue.formulaValue',
                "range": {
                    "sheetId": ids[len(ids)-4],
                    "startRowIndex": 1,
                    "startColumnIndex": 3,
                    "endColumnIndex": 5
                }
            }
        },
        {
            'updateCells': {
                "rows": update_names,
                "fields": 'userEnteredValue.stringValue',
                "range": {
                    "sheetId": ids[0],
                    "startRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1
                }
            }
        },
        {
            'updateCells': {
                "rows": [{
                    "values":[{
                        "userEnteredValue": {
                            "formulaValue": new_tour_columnG[:19] + '2' + new_tour_columnG[20:23] + tour_date + values[2][6][28:]
                        }
                    }]
                }],
                "fields": 'userEnteredValue.formulaValue',
                "range": {
                    "sheetId": ids[0],
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 3,
                    "endColumnIndex": 4
                }
            }
        },
        {
            'updateCells': {
                "rows": [{
                    "values":[{
                        "userEnteredValue": {
                            "formulaValue": rank_columnJ
                        }
                    },
                    {
                        "userEnteredValue": {
                            "formulaValue": rank_columnK
                        }
                    }]
                }],
                "fields": 'userEnteredValue.formulaValue',
                "range": {
                    "sheetId": ids[0],
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 9,
                    "endColumnIndex": 11
                }
            }
        },
        {
            'updateCells': {
                "rows": update_names,
                "fields": 'userEnteredValue.stringValue',
                "range": {
                    "sheetId": ids[len(ids)-4],
                    "startRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1
                }
            }
        },
        {
            'updateCells': {
                "rows": [{
                    "values":[{
                        "userEnteredValue": {
                            "formulaValue": columnD
                        }
                    },
                    {
                        "userEnteredValue": {
                            "formulaValue": columnE
                        }
                    }]
                }],
                "fields": 'userEnteredValue.formulaValue',
                "range": {
                    "sheetId": ids[len(ids)-6],
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 3,
                    "endColumnIndex": 5
                }
            }
        },
        {
            'updateCells': {
                "rows": [{
                    "values":[{
                        "userEnteredValue": {
                            "formulaValue": columnG
                        }
                    }]
                }],
                "fields": 'userEnteredValue.formulaValue',
                "range": {
                    "sheetId": ids[len(ids)-6],
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 6,
                    "endColumnIndex": 7
                }
            }
        },
        {
            'copyPaste': {
                "source": {
                    "sheetId": ids[len(ids)-4],
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 1,
                    "endColumnIndex": 8
                },
                "destination": {
                    "sheetId": ids[len(ids)-4],
                    "startRowIndex": 2,
                    "endRowIndex": len(uniq_names)+1,
                    "startColumnIndex": 1,
                    "endColumnIndex": 8
                },
                "pasteType": 'PASTE_NORMAL',
                "pasteOrientation": 'NORMAL'
            }
        },
        {
            'copyPaste': {
                "source": {
                    "sheetId": ids[len(ids)-6],
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 1
                },
                "destination": {
                    "sheetId": ids[len(ids)-6],
                    "startRowIndex": 2,
                    "endRowIndex": len(uniq_names)+1,
                    "startColumnIndex": 1
                },
                "pasteType": 'PASTE_NORMAL',
                "pasteOrientation": 'NORMAL'
            }
        },
        {
            'copyPaste': {
                "source": {
                    "sheetId": ids[0],
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 1
                },
                "destination": {
                    "sheetId": ids[0],
                    "startRowIndex": 2,
                    "endRowIndex": len(uniq_names)+1,
                    "startColumnIndex": 1
                },
                "pasteType": 'PASTE_NORMAL',
                "pasteOrientation": 'NORMAL'
            }
        }]
    }
    result = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

    return "Hi"

print(main('1Msq8pgWFj83DwLumVdgk84fSmBG_Uq3UcaJ7Ro_mk6Q', '08-17', '08-10'))