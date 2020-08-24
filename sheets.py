from googleapiclient.discovery import build # Builds the call
from googleapiclient.errors import HttpError # Catches HTTP errors
from google.oauth2 import service_account # Google Authentication
import os # OS-independent way of getting directory paths
import csv # File processing
from collections import OrderedDict # Most efficient way to get unique list values
import datetime

def updateSeeding(spreadsheet_id, tour_date, last_tour_date, new_month=False):
    last_tour_datetime = datetime.datetime.strptime(last_tour_date + str(datetime.datetime.today().year), '%m-%d%Y')
    new_tour_datetime = datetime.datetime.strptime(tour_date + str(datetime.datetime.today().year), '%m-%d%Y')
    if last_tour_datetime.month - new_tour_datetime.month != 0:
        new_month = True
    tour_delta = ((new_tour_datetime - last_tour_datetime).days)

    # Pull Service account credentials from local file
    creds = service_account.Credentials.from_service_account_file(os.path.join(os.path.dirname(__file__)) + '/credentials.json', scopes=['https://www.googleapis.com/auth/spreadsheets'])
    service = build('sheets', 'v4', credentials=creds) # Builds API service for Sheets

    # Get values from all the sheets
    result = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    properties = [sheet.get('properties') for sheet in result.get('sheets', [])]
    ids = [property.get('sheetId', 0) for property in properties]
    titles = [property.get('title', '') for property in properties]

    # Copy last tournament sheet into a new one
    new_sheet_id = copyTo(service, spreadsheet_id, ids[len(ids)-7])

    # Get values from the Tournament Turnout sheet to add another row
    turnout_values = getValues(service, spreadsheet_id, 'Tournament Turnout')
    last_turnout_index = len(turnout_values)
    turnout_columnA = turnout_values[-1][0] + tour_delta
    turnout_columnB = turnout_values[-1][1][:2] + tour_date + turnout_values[-1][1][7:]
    turnout_columnC = turnout_values[-1][2][:10] + tour_date + turnout_values[-1][2][15:]

    # Get values from last tournament and update them for the new tournament
    last_tour_values = getValues(service, spreadsheet_id, last_tour_date)
    tour_columnD = last_tour_values[2][3].replace('""', '"')
    tour_columnG = last_tour_values[2][6][:19] + '{}' + last_tour_values[2][6][20:23] + '{}' + last_tour_values[2][6][28:]
    tour_columnX = last_tour_values[2][-3][:32] + last_tour_date + last_tour_values[2][-3][37:-27] + last_tour_date + last_tour_values[2][-3][-22:]
    tour_columnY = last_tour_values[2][-2][:23] + last_tour_date + last_tour_values[2][-2][28:]
    tour_columnZ = last_tour_values[2][-1][:41] + last_tour_date + last_tour_values[2][-1][46:]

    # Get the index of the last row in the tournament sheet
    last_tour_index = 1
    for value in last_tour_values:
        if value[0]:
            last_tour_index = last_tour_values.index(value)+1

    # Get new tournament data from the CSV
    tour_csv, new_tour_index = getCSV()

    # Figure out the new index 
    new_tour_index += last_tour_index

    # The first batchUpdate that takes place
    # Updates the new tournament sheet, and Tournament Turnout
    body = {
        'requests': [
        {
            # Update new tournament sheet's name and place it after the last tournament
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
            # Add a row to Tournament Turnout
            'copyPaste': {
                "source": {
                    "sheetId": ids[len(ids)-5],
                    "startRowIndex": last_turnout_index-1,
                    "endRowIndex": last_turnout_index,
                    "startColumnIndex": 0,
                    "endColumnIndex": 5
                },
                "destination": {
                    "sheetId": ids[len(ids)-5],
                    "startRowIndex": last_turnout_index,
                    "endRowIndex": last_turnout_index+1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 5
                },
                "pasteType": 'PASTE_NORMAL',
                "pasteOrientation": 'NORMAL'
            }
        },
        {
            # Update the new Tournament Turnout row to pull from new sheet
            'updateCells': {
                "rows": [{
                    "values":[{
                        "userEnteredValue": {
                            "numberValue": turnout_columnA
                        }
                    },
                    {
                        "userEnteredValue": {
                            "formulaValue": turnout_columnB
                        }
                    },
                    {
                        "userEnteredValue": {
                            "formulaValue": turnout_columnC
                        }
                    }]
                }],
                "fields": 'userEnteredValue.numberValue,userEnteredValue.formulaValue',
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
            # Update first row, column B with new date for new tournament sheet
            'updateCells': {
                "rows": [{
                    "values": [{
                        "userEnteredValue": {
                            "stringValue": tour_date + '-' + str(datetime.datetime.today().year)
                        }
                    }]
                }],
                "fields": 'userEnteredValue.stringValue',
                "start": {
                    "sheetId": new_sheet_id,
                    "rowIndex": 0,
                    "columnIndex": 1
                }
            }
        },
        {
            # Update first data row, column D for new tournament sheet
            'updateCells': {
                "rows": [{
                    "values": [{
                        "userEnteredValue": {
                            "formulaValue": tour_columnD
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
            # Update first data row, column G for new tournament sheet
            'updateCells': {
                "rows": [{
                    "values": [{
                        "userEnteredValue": {
                            "formulaValue": tour_columnG.format('3', last_tour_date)
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
            # Update first data row, columns XYZ for new tournament sheet
            'updateCells': {
                "rows": [{
                    "values": [{
                        "userEnteredValue": {
                            "formulaValue": tour_columnX
                        }
                    },{
                        "userEnteredValue": {
                            "formulaValue": tour_columnY
                        }
                    },{
                        "userEnteredValue": {
                            "formulaValue": tour_columnZ
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
            # Empty the Placement column so data doesn't conflict for new tournament sheet
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
                    "endRowIndex": last_tour_index,
                    "startColumnIndex": 2,
                    "endColumnIndex": 3
                }
            }
        },
        {
            # Empty Dropped and Eliminated columns so the data doesn't skew the metrics for new tournament sheet
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
                    "endRowIndex": last_tour_index,
                    "startColumnIndex": 12,
                    "endColumnIndex": 14
                }
            }
        },
        {
            # Append csv data into the new tournament sheet
            'pasteData': {
                "coordinate": {
                    "sheetId": new_sheet_id,
                    "rowIndex": last_tour_index,
                    "columnIndex": 0
                },
                "data": tour_csv,
                "type": 'PASTE_NORMAL',
                "delimiter": '@'
            }
        },
        {
            # Copy first data row, columns D-L for metrics on every participant for new tournament sheet
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
            # Copy first data row, columns XYZ for metrics on every participant for new tournament sheet
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
            # Clear out rows below the last player for new tournament sheet
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
                    "startRowIndex": new_tour_index
                }
            }
        }]
    }
    # Execute first batchUpdate
    result = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

    # Get all the names from the new tournament sheet
    tour_values = getValues(service, spreadsheet_id, tour_date)[2:]
    names = [name[0] for name in tour_values]

    # Get all the unique names
    uniq_names = list(OrderedDict.fromkeys(names))

    # Find all the names that have duplicates
    dup_names = set()
    dup_names2 = set()
    for name in names:
        if name in dup_names:
            dup_names2.add(name)
        else:
            dup_names.add(name)

    # Delete the first instance of each name (This will be the row with the placement column NOT filled in)
    deleteRequests(service, spreadsheet_id, new_sheet_id, [names.index(name)+2 for name in dup_names2])

    # Get the values to update columns D and E in Attendance
    attend_values = getValues(service, spreadsheet_id, 'Attendance')
    attend_column = attend_values[1][3][:23] + '{}' + attend_values[1][3][28:34] + '{}' + attend_values[1][3][35:37] + '{}' + attend_values[1][3][39:]

    # Get the values to update columns J and K in Rankings
    rank_values = getValues(service, spreadsheet_id, 'Rankings')
    rank_columnJ = rank_values[1][9] + rank_values[1][9][-64:-36] + tour_date + rank_values[1][9][-31:]
    rank_columnK = rank_values[1][10] + rank_values[1][10][-123:-91] + tour_date + rank_values[1][10][-86:-37] + tour_date + rank_values[1][10][-32:]

    # Get the values to update columns D, E, and G in Win Rates
    win_rate_values = getValues(service, spreadsheet_id, 'Win Rates')
    win_rate_D_add = win_rate_values[1][3][-51:-27] + tour_date + win_rate_values[1][3][-22:]
    win_rate_E_add = win_rate_values[1][4][-51:-27] + tour_date + win_rate_values[1][4][-22:]
    win_rate_columnD = win_rate_values[1][3][:-1] + win_rate_D_add
    win_rate_columnE = win_rate_values[1][4][:-1] + win_rate_E_add
    win_rate_columnG = win_rate_values[1][6][:23] + tour_date + win_rate_values[1][6][28:]

    update_names = []
    # Add each unique name as its own row
    for name in uniq_names:
        update_names.append({
            "values":[{
                "userEnteredValue": {
                    "stringValue": str(name)
                }
            }]
        })

    requests = [{
            # Update Seeding with unique names
            'updateCells': {
                "rows": update_names,
                "fields": 'userEnteredValue.stringValue',
                "range": {
                    "sheetId": ids[len(ids)-2],
                    "startRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1
                }
            }
        },
        {
            # Copy first row to all rows for Seeding
            'copyPaste': {
                "source": {
                    "sheetId": ids[len(ids)-2],
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 1,
                    "endColumnIndex": 3
                },
                "destination": {
                    "sheetId": ids[len(ids)-2],
                    "startRowIndex": 2,
                    "endRowIndex": len(uniq_names)+1,
                    "startColumnIndex": 1,
                    "endColumnIndex": 3
                },
                "pasteType": 'PASTE_NORMAL',
                "pasteOrientation": 'NORMAL'
            }
        },
        {
            # Update the first data row, columns D-E for Attendance to the new date
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
                    "endRowIndex": 2,
                    "startColumnIndex": 3,
                    "endColumnIndex": 5
                }
            }
        },
        {
            # Update Attendance with unique names
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
            # Copy first row to all rows for Attendance
            'copyPaste': {
                "source": {
                    "sheetId": ids[len(ids)-4],
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 3,
                    "endColumnIndex": 5
                },
                "destination": {
                    "sheetId": ids[len(ids)-4],
                    "startRowIndex": 2,
                    "endRowIndex": len(uniq_names)+1,
                    "startColumnIndex": 3,
                    "endColumnIndex": 5
                },
                "pasteType": 'PASTE_NORMAL',
                "pasteOrientation": 'NORMAL'
            }
        },
        {
            # Update the Rankings sheet with all the unique names
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
            # Update first data row, column D with new date for Rankings
            'updateCells': {
                "rows": [{
                    "values":[{
                        "userEnteredValue": {
                            "formulaValue": tour_columnG.format('2', tour_date)
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
            # Update first data row, columns J-K to append the new tournament sheet for Rankings
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
            # Copy first row to all rows for Rankings
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
        },
        # Update first data row, columns D-E appending the new tournament sheet for Win Rates
        {
            'updateCells': {
                "rows": [{
                    "values":[{
                        "userEnteredValue": {
                            "formulaValue": win_rate_columnD
                        }
                    },
                    {
                        "userEnteredValue": {
                            "formulaValue": win_rate_columnE
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
            # Update first data row, column G with new tournament sheet for Win Rates
            'updateCells': {
                "rows": [{
                    "values":[{
                        "userEnteredValue": {
                            "formulaValue": win_rate_columnG
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
            # Copy first row to all rows for Win Rates
            'copyPaste': {
                "source": {
                    "sheetId": ids[len(ids)-6],
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 3,
                    "endColumnIndex": 7
                },
                "destination": {
                    "sheetId": ids[len(ids)-6],
                    "startRowIndex": 2,
                    "endRowIndex": len(uniq_names)+1,
                    "startColumnIndex": 3,
                    "endColumnIndex": 7
                },
                "pasteType": 'PASTE_NORMAL',
                "pasteOrientation": 'NORMAL'
            }
        }]

    if new_month:
        turnout_months = []
        for x in range(len(turnout_values)):
            if x > 0:
                turnout_months.append((datetime.datetime.strptime("12-30-1899", '%m-%d-%Y') + datetime.timedelta(days=turnout_values[x][0])).month)
        tours_in_last_month = turnout_months.count(last_tour_datetime.month)

        win_rate_columnJ = win_rate_values[1][9][:-1] + win_rate_D_add
        win_rate_columnK = win_rate_values[1][10][:-1] + win_rate_E_add

        requests.extend([{
                # Update the first data row, columns D-E for Attendance to the new date
                'updateCells': {
                    "rows": [{
                        "values":[{
                            "userEnteredValue": {
                                "formulaValue": attend_column.format(last_tour_date, 'X', 24)
                            }
                        },
                        {
                            "userEnteredValue": {
                                "formulaValue": attend_column.format(last_tour_date, 'Y', 25)
                            }
                        }]
                    }],
                    "fields": 'userEnteredValue.formulaValue',
                    "range": {
                        "sheetId": ids[len(ids)-4],
                        "startRowIndex": 1,
                        "endRowIndex": 2,
                        "startColumnIndex": 6,
                        "endColumnIndex": 8
                    }
                }
            },
            {
                # Copy first row to all rows for Attendance
                'copyPaste': {
                    "source": {
                        "sheetId": ids[len(ids)-4],
                        "startRowIndex": 1,
                        "endRowIndex": 2,
                        "startColumnIndex": 6,
                        "endColumnIndex": 8
                    },
                    "destination": {
                        "sheetId": ids[len(ids)-4],
                        "startRowIndex": 2,
                        "endRowIndex": len(uniq_names)+1,
                        "startColumnIndex": 6,
                        "endColumnIndex": 8
                    },
                    "pasteType": 'PASTE_NORMAL',
                    "pasteOrientation": 'NORMAL'
                }
            },
            {
                # Add monthly information into Tournament Turnout
                'copyPaste': {
                    "source": {
                        "sheetId": ids[len(ids)-5],
                        "startRowIndex": last_turnout_index-tours_in_last_month,
                        "endRowIndex": last_turnout_index-tours_in_last_month+2,
                        "startColumnIndex": 7,
                        "endColumnIndex": 9
                    },
                    "destination": {
                        "sheetId": ids[len(ids)-5],
                        "startRowIndex": last_turnout_index,
                        "endRowIndex": last_turnout_index+2,
                        "startColumnIndex": 7,
                        "endColumnIndex": 9
                    },
                    "pasteType": 'PASTE_NORMAL',
                    "pasteOrientation": 'NORMAL'
                }
            },
            {
                # Update the Tournament Turnout with the previous monthly information
                'updateCells': {
                    "rows": [{
                        "values":[{
                            "userEnteredValue": {
                                "formulaValue": "=ceiling(average($B" + str(last_turnout_index+1-tours_in_last_month) + ":$B" + str(last_turnout_index) + "))"
                            }
                        }]
                    }],
                    "fields": 'userEnteredValue.formulaValue',
                    "range": {
                        "sheetId": ids[len(ids)-5],
                        "startRowIndex": last_turnout_index - tours_in_last_month,
                        "endRowIndex": last_turnout_index - tours_in_last_month + 1,
                        "startColumnIndex": 7,
                        "endColumnIndex": 8
                    }
                }
            },
            {
                # Update the Tournament Turnout with the previous monthly information
                'updateCells': {
                    "rows": [{
                        "values":[{
                            "userEnteredValue": {
                                "formulaValue": "=SUM($C" + str(last_turnout_index+1-tours_in_last_month) + ":$C" + str(last_turnout_index) + ")"
                            }
                        }]
                    }],
                    "fields": 'userEnteredValue.formulaValue',
                    "range": {
                        "sheetId": ids[len(ids)-5],
                        "startRowIndex": last_turnout_index - tours_in_last_month + 1,
                        "endRowIndex": last_turnout_index - tours_in_last_month + 2,
                        "startColumnIndex": 7,
                        "endColumnIndex": 8
                    }
                }
            },
            {
                # Update first row, column H with last tournament of last month for Rankings
                'updateCells': {
                    "rows": [{
                        "values":[{
                            "userEnteredValue": {
                                "stringValue": "*Prev = " + str(last_tour_date)
                            }
                        }]
                    }],
                    "fields": 'userEnteredValue.stringValue',
                    "range": {
                        "sheetId": ids[0],
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 7,
                        "endColumnIndex": 8
                    }
                }
            },
            {
                # Update first data row, column G with new date for Rankings
                'updateCells': {
                    "rows": [{
                        "values":[{
                            "userEnteredValue": {
                                "formulaValue": tour_columnG.format('2', last_tour_date)
                            }
                        }]
                    }],
                    "fields": 'userEnteredValue.formulaValue',
                    "range": {
                        "sheetId": ids[0],
                        "startRowIndex": 1,
                        "endRowIndex": 2,
                        "startColumnIndex": 6,
                        "endColumnIndex": 7
                    }
                }
            },
            {
                # Copy first data row, column G to all rows for Rankings
                'copyPaste': {
                    "source": {
                        "sheetId": ids[0],
                        "startRowIndex": 1,
                        "endRowIndex": 2,
                        "startColumnIndex": 6,
                        "endColumnIndex": 7
                    },
                    "destination": {
                        "sheetId": ids[0],
                        "startRowIndex": 2,
                        "endRowIndex": len(uniq_names)+1,
                        "startColumnIndex": 6,
                        "endColumnIndex": 7
                    },
                    "pasteType": 'PASTE_NORMAL',
                    "pasteOrientation": 'NORMAL'
                }
            },
            {
                # Update first data row, columns J and K for monthly update for Win Rates
                'updateCells': {
                    "rows": [{
                        "values":[{
                            "userEnteredValue": {
                                "formulaValue": win_rate_columnJ
                            }
                        },
                        {
                            "userEnteredValue": {
                                "formulaValue": win_rate_columnK
                            }
                        }]
                    }],
                    "fields": 'userEnteredValue.formulaValue',
                    "range": {
                        "sheetId": ids[len(ids)-6],
                        "startRowIndex": 1,
                        "endRowIndex": 2,
                        "startColumnIndex": 9,
                        "endColumnIndex": 11
                    }
                }
            },
            {
                # Update first data row, column M for monthly update for Win Rates
                'updateCells': {
                    "rows": [{
                        "values":[{
                            "userEnteredValue": {
                                "formulaValue": win_rate_columnG
                            }
                        }]
                    }],
                    "fields": 'userEnteredValue.formulaValue',
                    "range": {
                        "sheetId": ids[len(ids)-6],
                        "startRowIndex": 1,
                        "endRowIndex": 2,
                        "startColumnIndex": 12,
                        "endColumnIndex": 13
                    }
                }
            },
            {
                # Copy first data row, columns J, K and M to all rows for Win Rates
                'copyPaste': {
                    "source": {
                        "sheetId": ids[len(ids)-6],
                        "startRowIndex": 1,
                        "endRowIndex": 2,
                        "startColumnIndex": 9,
                        "endColumnIndex": 13
                    },
                    "destination": {
                        "sheetId": ids[len(ids)-6],
                        "startRowIndex": 2,
                        "endRowIndex": len(uniq_names)+1,
                        "startColumnIndex": 9,
                        "endColumnIndex": 13
                    },
                    "pasteType": 'PASTE_NORMAL',
                    "pasteOrientation": 'NORMAL'
                }
            }
        ])

    body = {}
    body.update({"requests": requests})
    result = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

    ids.insert(-6, new_sheet_id)
    sortRanges(service, spreadsheet_id, ids)

    return "Hi"

def getValues(service, spreadsheet_id, sheet_name):
    return service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range="'" + sheet_name + "'!A1:Z", valueRenderOption='FORMULA').execute().get('values', [])

def copyTo(service, spreadsheet_id, sheet_id):
    body = {
        'destination_spreadsheet_id': spreadsheet_id,
    }
    return service.spreadsheets().sheets().copyTo(spreadsheetId=spreadsheet_id, sheetId=sheet_id, body=body).execute().get('sheetId')

def getCSV():
    index = 0
    tour_csv = ''
    with open(os.path.join(os.path.dirname(__file__)) + '/tournamentData.csv', 'r', newline='') as csv_file:
        tour_reader = csv.reader(csv_file, delimiter='@', quotechar='|')
        for row in tour_reader:
            tour_csv += '@'.join(row) + '\n'
            index += 1
    return tour_csv, index

def deleteRequests(service, spreadsheet_id, sheet_id, delete_indices):
    deleteReqs = []
    prev_indices = []
    rev_i = len(delete_indices)

    for i in range(len(delete_indices)):
        index = delete_indices[i]
        if i:
            for prev_index in prev_indices:
                if index > prev_index:
                    index -= 1
        deleteReqs.append(
            {
                'deleteRange': {
                    "range": {
                        "sheetId": sheet_id,
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

def sortRanges(service, spreadsheet_id, ids):
    body = {
        "requests": [{
            'sortRange': {
                'range': {
                    'sheetId': ids[0],
                    'startRowIndex': 1,
                    'startColumnIndex': 0
                },
                'sortSpecs': [{
                    'sortOrder': 'ASCENDING',
                    'dimensionIndex': 1
                }]
            }
        },
        {
            'sortRange': {
                'range': {
                    'sheetId': ids[-7],
                    'startRowIndex': 2,
                    'startColumnIndex': 0
                },
                'sortSpecs': [{
                    'sortOrder': 'ASCENDING',
                    'dimensionIndex': 2
                }]
            }
        },
        {
            'sortRange': {
                'range': {
                    'sheetId': ids[-6],
                    'startRowIndex': 1,
                    'startColumnIndex': 0
                },
                'sortSpecs': [{
                    'sortOrder': 'ASCENDING',
                    'dimensionIndex': 1
                }]
            }
        },
        {
            'sortRange': {
                'range': {
                    'sheetId': ids[-4],
                    'startRowIndex': 1,
                    'startColumnIndex': 0
                },
                'sortSpecs': [{
                    'sortOrder': 'ASCENDING',
                    'dimensionIndex': 1
                }]
            }
        },
        {
            'sortRange': {
                'range': {
                    'sheetId': ids[-2],
                    'startRowIndex': 1,
                    'startColumnIndex': 0
                },
                'sortSpecs': [{
                    'sortOrder': 'ASCENDING',
                    'dimensionIndex': 1
                }]
            }
        }]
    }
    result = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

print(updateSeeding('1Msq8pgWFj83DwLumVdgk84fSmBG_Uq3UcaJ7Ro_mk6Q', '08-17', '08-10', new_month=True))