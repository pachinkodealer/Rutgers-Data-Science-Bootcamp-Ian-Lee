import os
import csv
#reading csv path
budget_csv = os.path.join("RBCDS", "Starter_Code","Week 3","Challenge", "Images", "PyBank", "Resources", "budget_data.csv") # use commas, folders
#Finding relative path, RBCDS\Starter_Code\Week 3\Challenge\Images\PyBank\Resources\budget_data.csv

#Creating list of variable
budget_data = []
total_month = 0
total_p_L = 0
avg_change = 0
previous_value = None
sum_change = 0
greatest_increase = None
greatest_decrease = None
# Create a list to store the data that will be written to the CSV file
output_data = []

#reads the csv
with open(budget_csv) as csvfile:
    budget_data_csv = csv.reader(csvfile, delimiter=",")
    
    #skipping the header row
    next(budget_data_csv, None)

    #Getting total number of months in dataset
    for row in budget_data_csv:
        total_month = total_month + 1
    #getting net total of profit/losses
        budget_data = int(row[1])
        total_p_L = budget_data + total_p_L

    #current value of row
        current_value = int(row[1])
        #calculating total change
        if previous_value is not None:
            #calculate change from previous row
            change = current_value - previous_value
            #print(change)
            sum_change = sum_change + change

        #calculating whether current change is greater than greatest increase
            if greatest_increase is None or change > greatest_increase:
                greatest_increase = change
                greatest_increase_month = row[0]  # Store the month
        
        #calculating whether current change is less than greatest decrease
            if greatest_decrease is None or change < greatest_decrease:
                greatest_decrease = change
                greatest_decrease_month = row[0]  # Store the month
        
        #making previous value current value from start of data
        previous_value = current_value

#getting the average from sum_change
avg_change = round(sum_change/(total_month-1), 2)

#printing financial analysis
print("Financial Analysis")
#print dashes
print("--------------------------")
#printing total months
print(f"Total Months: {total_month}")
#printing total profit/loss
print(f"Total Profit/Loss: ${total_p_L}")
#printing average change
print(f"Average Change: ${avg_change}")
# Printing greatest increase and the corresponding month
print(f"Greatest Increase: {greatest_increase_month} ${greatest_increase}")
# Printing greatest decrease and the corresponding month
print(f"Greatest Increase: {greatest_decrease_month} ${greatest_decrease}")

# Define the output txt file path
output_txt = os.path.join("RBCDS", "Starter_Code","Week 3","Challenge", "Images", "PyBank", "Resources", "output.txt")
#writing to the txt file
with open(output_txt, "w") as outputfile:
    outputfile.write("Financial Analysis\n")
    outputfile.write("--------------------------\n")
    outputfile.write(f"Total Months: {total_month}\n")
    outputfile.write(f"Total Profit/Loss: {total_p_L}\n")
    outputfile.write(f"Average Change: {avg_change}\n")
    outputfile.write(f"Greatest Increase in Profits: {greatest_increase_month} (${greatest_increase})\n")
    outputfile.write(f"Greatest Decrease in Profits: {greatest_decrease} (${greatest_decrease_month})\n")
#printing txt file exported
print("Financial analysis data has been exported to output.txt")