import json
import utils
from math import e
import sys
        
def main():
    file_path = 'budget_data.json'
    balance, expense_list = utils.read_json(file_path)
    if balance == 0:
        balance = float(input("Please enter your initial budget: "))
    total = balance
    
    while True:
        print("\nWhat would you like to do?")
        print("1. Add an expense")
        print("2. Show budget details")
        print("3. Print my budget detail in Excel")
        print("4. Exit")
        
        choose = 0
        
        while choose == 0:
            try:
                choose = int(input("Enter your choice: "))
            except:
                print("Invalid characters. Please choose a number")
                continue
            
        if choose == 1:
            total -= utils.add_expense(balance, expense_list)
        elif choose == 2:
            total_spend = utils.budget_detail(balance, expense_list)
        elif choose == 3:
            utils.budget_sheet(expense_list, balance)
        elif choose == 4:
            utils.save_budget_detail(file_path, balance, expense_list)
            sys.exit()
        else:
            print("Invalid number. Please try again.")





if __name__ in "__main__":
    print("#############################")
    print("# Welcome to Budget Tracker #")
    print("#############################")
    main()