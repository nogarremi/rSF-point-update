from tasks import sheets, tournament
from tasks.secret import api_user, api_key

new_tour_date = ''
print("Welcome to Nogarremi's r/StreetFighter Sheet Updater!")
while True:
    print("1) Pull Tournament Data\n2) Update Spreadsheet\n3) Quit Program")
    u_input = int(input("Choose Your Option: "))
    
    if u_input == 3:
        print("Goodbye!")
        break
    elif u_input == 1:
        subdomain = input("Challonge Subdomain(Blank if there is no Challonge community): ").lower()
        bracket_id = input("Challonge Bracket Id: ").lower()

        new_tour_date = tournament.getTournamentCSV(api_user, api_key, subdomain, bracket_id)
        print(new_tour_date + " has been processed")
    elif u_input == 2:
        spreadsheet_id = input("Google Sheets Spreadsheet Id: ")
        tour_date = new_tour_date if new_tour_date else input("Tournament Date(mm-dd): ")
        last_tour_date = input("Last Tournament Date(mm-dd): ")
        new_month = bool(int(input("New Month(1 for True/0 for False): ")))

        sheets.updateSeeding(spreadsheet_id, tour_date, last_tour_date, new_month)
    else:
        print("Invalid input. Try again")
        continue
    print()