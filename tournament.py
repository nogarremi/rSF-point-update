import challonge
import csv
from secret import api_user, api_key
    
class Player:
    def __init__(self, username, name, id, rank):
        self.username = username
        
        if name == "":
            self.name = username
        else:
            self.name = name
            
        self.id = id
        self.rank = rank
        self.dropped = ""
        self.eliminated = ""

challonge.set_credentials(api_user, api_key)

tournament = challonge.tournaments.show("redditfighting-s2u03n9v") # Tournament link
tour_date = tournament["name"].split('(')[1].replace("/", "-")[:4]

participants = challonge.participants.index(tournament["id"])
players = []

for participant in participants:
    username = participant["username"]
    name = participant["name"]
    id = participant["id"]
    rank = participant["final-rank"]

    players.append(Player(username, name, id, rank))

matches = challonge.matches.index(tournament["id"])

for match in matches:
    # Initialize loser and winner variables
    loser = players[0]
    winner = players[0]
    lScore = ""
    wScore = ""
    hyphenPos = 0

    # Check for DQs that aren't marked as such. (Scores of -1-0, 0--1, 0-0)
    # Only worry about score if match is not marked as a DQ
    if match["forfeited"] != True:

        # Get score string from match
        scoreString = match["scores-csv"]

        # Test to see if it has scores per set or games
        if scoreString.count(",") == 0:

            # Split the scores into separate variables
            scores = scoreString.split('-')

            # Change forfeited value if score has a '-1' or scores == 0
            if len(scores) == 3:
                match["forfeited"] = True
            else:
                if int(scores[0]) + int(scores[1]) == 0:
                    match["forfeited"] = True

    for player in players:
        if match["loser-id"] == player.id:
            loser = player
            
        if match["winner-id"] == player.id:
            winner = player

    if match["round"] > 0:
        if match != matches[-1]:
            loser.dropped = winner.username
            
            #If the match was forfeited
            if match["forfeited"] == True:
                loser.dropped = "DQ"
        else:
            loser.eliminated = winner.username

            #If the match was forfeited
            if  match["forfeited"] == True:
                loser.eliminated = "DQ"
    else:
        loser.eliminated = winner.username

        #If the match was forfeited
        if match["forfeited"] == True:
            loser.eliminated = "DQ"

# Export data into a .csv file
with open('tournamentData.csv', 'w', newline='') as file:
    writer = csv.writer(file, delimiter='@', quotechar='|')

    # Write player data to csv file
    full_count = list(range(3, len(players)+3))
    for player in players:
        count = full_count.pop(0)
        if not player.rank == None:
            # Add data to dataList

            dataList = [player.username, player.name, player.rank]

            dataList.extend(['']*9)

            # Add rest of data to dataList
            dataList.append(player.dropped)
            dataList.append(player.eliminated)

            # Write dataList to csv
            writer.writerows([dataList])
            # print("Player ", x, " data written.")