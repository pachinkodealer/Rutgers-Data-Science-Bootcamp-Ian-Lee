import os
import csv
#reading csv path
election_csv = os.path.join("RBCDS", "Starter_Code","Week 3","Challenge", "Images", "PyPoll", "Resources", "election_data.csv") # use commas, folders
#initilize variables
total_votes = 0
stockham_votes = 0 
degette_votes = 0
doane_votes = 0
#initialize winner
winner = None

#reads the csv
with open(election_csv) as csvfile:
    election_data_csv = csv.reader(csvfile, delimiter=",")
    
    #skipping the header row
    next(election_data_csv, None)

    #Getting total number of months in dataset
    for row in election_data_csv:
        total_votes = total_votes + 1
        candidate_name = row[2]

        #If candidate's name is "Charles.."
        if candidate_name == "Charles Casper Stockham":
            stockham_votes += 1
        elif candidate_name == "Diana DeGette":
            degette_votes += 1
        elif candidate_name == "Raymon Anthony Doane":
            doane_votes += 1

#Determining the winner
if stockham_votes > degette_votes and stockham_votes > doane_votes:
    winner = "Charles Casper Stockham"
elif degette_votes > stockham_votes and degette_votes > doane_votes:
    winner = "Diana DeGette"
elif doane_votes > stockham_votes and doane_votes > degette_votes:
    winner = "Raymon Anthony Doane"

#Calculating vote percentage
stockham_percentage = round((stockham_votes / total_votes) * 100, 3)
degette_percentage = round((degette_votes / total_votes) * 100, 3)
doane_percentage = round((doane_votes / total_votes) * 100, 3)

#printing election results
print("Election Results")
print("--------------------------")
#printing total months
print(f"Total Votes: {total_votes}")
print("--------------------------")
#print total votes for candidate and percentage of votes for candidate
print(f"Charles Casper Stockham: {stockham_percentage}% ({stockham_votes})")
print(f"Diana Degette: {degette_percentage}% ({degette_votes})")
print(f"Raymon Anthony Doane: {doane_percentage}% ({doane_votes})")
print("--------------------------")
print(f"Winner: {winner}")
print("--------------------------")
#file path to export
output_txt = os.path.join("RBCDS", "Starter_Code", "Week 3", "Challenge", "Images", "PyPoll", "Resources", "output.txt")

# Write the results to the output text file
with open(output_txt, "w") as outputfile:
    outputfile.write("Election Results\n")
    outputfile.write("--------------------------\n")
    outputfile.write(f"Total Votes: {total_votes}\n")
    outputfile.write("--------------------------\n")
    outputfile.write(f"Charles Casper Stockham: {stockham_percentage:.3f}% ({stockham_votes})\n")
    outputfile.write(f"Diana Degette: {degette_percentage:.3f}% ({degette_votes})\n")
    outputfile.write(f"Raymon Anthony Doane: {doane_percentage:.3f}% ({doane_votes})\n")
    outputfile.write("--------------------------\n")
    outputfile.write(f"Winner: {winner}\n")
    outputfile.write("--------------------------\n")
#printing out txt file exported
print("Election results have been exported to output.txt")