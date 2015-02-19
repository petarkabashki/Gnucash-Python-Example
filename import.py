#! /usr/bin/env python

import sys
import os, glob
import shutil
import csv
import time

from csv import DictReader

# sys.path.append("/home/grenada/unstable/gnucash/lib/python2.7/site-packages")
import gnucash

orig_file = "gnucash-data/gnucash-data.template"
new_file = "gnucash-data/ntags-ARD-2013-12-31.xml"

for f in glob.glob( "{0}/{1}*".format( os.getcwd(), new_file ) ) :
    os.remove(f)

shutil.copy(orig_file, new_file)

start_date = time.strptime( "01/01/2013", "%d/%m/%Y") 
end_date = time.strptime( "31/12/2013", "%d/%m/%Y")

months_hash = {"JAN": 1, "FEB":2, "MAR":3, "APR":4, "MAY":5, "JUN":6, "JUL":7, "AUG":8,"SEP":9,"OCT":10,"NOV":11,"DEC":12}


session = gnucash.Session( "xml://{0}/{1}".format( os.getcwd(), new_file ), is_new=False)

book = session.book
root_account = book.get_root_account()

comm_table = book.get_table()
gbp = comm_table.lookup("CURRENCY", "GBP")

def account_from_path(top_account, account_path, original_path=None):
		if original_path==None: original_path = account_path
		account, account_path = account_path[0], account_path[1:]
		account = top_account.lookup_by_name(account)
		if account.get_instance() == None:
				raise Exception(
				    "path " + ''.join(original_path) + " could not be found")
		if len(account_path) > 0 :
				return account_from_path(account, account_path, original_path)
		else:
				return account
				
acc_accounts_payable = root_account.lookup_by_name(  "Accounts Payable"  )
acc_payroll_expenses = root_account.lookup_by_name(  "Payroll Expenses"  )
acc_corporation_tax = root_account.lookup_by_name(  "Corporation Tax Payable"  )
acc_employer_nics = root_account.lookup_by_name(  "Employer NICs"  )
acc_interest_income = root_account.lookup_by_name(  "Interest Income"  )
acc_sales = root_account.lookup_by_name(  "Sales"  )
acc_accounts_receivable = root_account.lookup_by_name(  "Accounts Receivable"  )
acc_checking_account = root_account.lookup_by_name(  "Checking Account"  )
acc_savings_account = root_account.lookup_by_name(  "Savings Account"  )
acc_cash = root_account.lookup_by_name(  "Petty Cash"  )
acc_administration_expenses = account_from_path( root_account , [ "Expenses",  "Administration" ] )
acc_opening_balances = root_account.lookup_by_name(  "Opening Balances"  )
acc_temp = root_account.lookup_by_name(  "temp"  )
acc_miscellaneous = root_account.lookup_by_name(  "Miscellaneous"  )
acc_professional_fees = root_account.lookup_by_name( "Professional Fees"  )
       
def initialize_transaction(description, date):
    dd,mm,yy = date
    trans = gnucash.Transaction(book)
    trans.BeginEdit()
    trans.SetCurrency(gbp)
    trans.SetDescription( description )
    trans.SetDate( dd,mm,yy)
    return trans
        
def initialize_split(value, account, trans, memo):
    split = gnucash.Split(book)
    split.SetValue(value)
    split.SetAccount(account)
    split.SetParent(trans)
    split.SetMemo(memo)
    return split

def create_gnc_from_string(str_value, sign):
    return gnucash.GncNumeric( sign * 10000 * float(str_value ), 10000)

def import_payroll_csv(file_name, employee_name):

    csv_file = file(file_name)
    
    reader = DictReader(csv_file)
    for data in reader:
        transaction_date = time.strptime( data["Acc Date"], "%d/%m/%Y")
        
#        if transaction_date < start_date or transaction_date > end_date:
#            continue
        
        dd,mm,yy = map( int, data["Acc Date"].split("/") )
        
        trans1 = initialize_transaction("{0} Salary ({1})".format(employee_name,data["Acc Date"]), (dd,mm,yy) ) 
        
        initialize_split(gnucash.GncNumeric(   100 * float(data["Gross Payment"] ), 100), acc_payroll_expenses, trans1, "{0} Gross salary - {1}".format(employee_name,data["Acc Date"]) )
        initialize_split(gnucash.GncNumeric( - 100 * float(data["Net Payment"] ), 100), acc_accounts_payable, trans1, "{0} NET salary - {1}".format(employee_name,data["Acc Date"]) )

        if float(data["Income Tax"] ):
            initialize_split(gnucash.GncNumeric( - 100 * float(data["Income Tax"] ), 100), acc_accounts_payable, trans1, "{0} Income Tax - {1}".format(employee_name,data["Acc Date"]) )
        
        if float(data["NICs PR"] ):
            initialize_split(gnucash.GncNumeric( - 100 * float(data["NICs PR"] ), 100), acc_accounts_payable, trans1, "{0} NICs - {1}".format(employee_name,data["Acc Date"]) )

        trans1.CommitEdit()
        
        if float(data["NICs SEC"] ):
            trans2 = initialize_transaction("{0} Employer NICs ({1})".format(employee_name,data["Acc Date"]), (dd,mm,yy) ) 
            initialize_split(gnucash.GncNumeric(   100 * float(data["NICs SEC"] ), 100), acc_employer_nics, trans2, "{0} NICs SEC".format(employee_name) )
            initialize_split(gnucash.GncNumeric( - 100 * float(data["NICs SEC"] ), 100), acc_accounts_payable, trans2, "{0} NICs SEC".format(employee_name) )
            trans2.CommitEdit()
            
    csv_file.close()

def import_checking_account(csv):
    csv_file = file(csv)
       
    reader = DictReader(csv_file)
    for data in reader:
        dd,mm,yy = data["Date"].split() 
        dd,mm,yy = int(dd), months_hash[mm.upper()], int(yy)
        transaction_date = time.strptime( "{0}/{1}/{2}".format(dd,mm,yy), "%d/%m/%Y")
        
        if transaction_date < start_date or transaction_date > end_date:
            continue
         
        account = acc_temp
        
        amount = 0.0
        
        if data["Paid in"].strip() != '':
        		amount = float(data["Paid in"])
        		account = acc_accounts_receivable
        elif data["Paid out"].strip() != '' and data["Description"].find("HMRC CORP TAX CUMB") > -1 :
        		amount = -float(data["Paid out"])
        		account = acc_corporation_tax
        elif data["Paid out"].strip() != '':
        		amount = -float(data["Paid out"])
        		account = acc_accounts_payable
        else:
        		continue
                
        if data["Type"] == "TFR":
        		continue
        elif data["Type"] == " ": 
            print "empty type"
            continue
            
#        elif data["Type"] == "VIS" and ( data["Description"].find("COMPANIES HSE FILE INTERNET") > -1 \
#			 or data["Description"].find("MASTER COVER INSUR NEW BARNET") > -1 \
#			 or data["Description"].find("P C G LIMITED WEST DRAYTON") > -1 \
#			 or data["Description"].find("WWW.INSUREDRISK.CO INTERNET") > -1):
#            account = acc_professional_fees
            
        trans1 = initialize_transaction(data["Description"], (dd,mm,yy) ) 
        initialize_split(gnucash.GncNumeric(   10000 * amount, 10000), acc_checking_account, trans1, "" )
        initialize_split(gnucash.GncNumeric( - 10000 * amount, 10000), account, trans1, "" )
        trans1.CommitEdit()
        
    csv_file.close()

def import_savings_account(csv):
    
    #Checking account
    csv_file = file(csv)
       
    reader = DictReader(csv_file)
    for data in reader:
        print data
        print "\n"
        dd,mm,yy = data["Date"].split() 
        dd,mm,yy = int(dd), months_hash[mm.upper()], int(yy)
        transaction_date = time.strptime( "{0}/{1}/{2}".format(dd,mm,yy), "%d/%m/%Y")
        
        if transaction_date < start_date or transaction_date > end_date:
            continue
         
        account = acc_temp
        
        amount = 0
        
        if data["Paid in"].strip() != '':
        		amount = float(data["Paid in"])
        elif data["Paid out"].strip() != '':
        		amount = -float(data["Paid out"])
        else:
        		continue
                
        if data["Type"] == " ": 
            print "empty type"
            continue
        elif data["Type"] == "TFR":
            account = acc_checking_account
        elif data["Type"] == "CR" and data["Description"].find("GROSS INTEREST") != None :
            account = acc_interest_income
        
        trans1 = initialize_transaction(data["Description"], (dd,mm,yy) ) 
        initialize_split(gnucash.GncNumeric( 10000 * amount , 10000), acc_savings_account, trans1, "" )
        initialize_split(gnucash.GncNumeric( -10000 * amount , 10000), account, trans1, "" )
        trans1.CommitEdit()
        
    csv_file.close()

#
#def import_additional():
#    
#    #Checking account
#    csv_file = file("src-data/Additional-Transactions.csv")
#       
#    reader = DictReader(csv_file)
#    for data in reader:
#        
#        transaction_date = time.strptime( data["Date"], "%d/%m/%Y")
#        
#        if transaction_date < start_date or transaction_date > end_date:
#            continue
#        
#        dd,mm,yy = map( int, data["Date"].split("/") )
#                
#        if transaction_date < start_date or transaction_date > end_date:
#            continue
#         
#        account_from    = root_account.lookup_by_name(  data["From Account"]  )
#        account_to      = root_account.lookup_by_name(  data["To Account"]  )
#
#        
#        amount = float( data["Amount"] )
#        
#        trans1 = initialize_transaction(data["Description"], (dd,mm,yy) ) 
#        initialize_split(gnucash.GncNumeric(   100 * amount , 100), account_from, trans1, "" )
#        initialize_split(gnucash.GncNumeric( - 100 * amount, 100), account_to, trans1, "" )
#        trans1.CommitEdit()
#        
#    csv_file.close()

def import_sales_csv(file_name):

    csv_file = file(file_name)
    
    reader = DictReader(csv_file)
    for data in reader:
        if(not data["Date"]): continue
        
        transaction_date = time.strptime( data["Date"], "%d/%m/%Y")
        
        if transaction_date < start_date or transaction_date > end_date:
            continue
        
        dd,mm,yy = map( int, data["Date"].split("/") )
        
        trans1 = initialize_transaction("Invoice - {0} - {1}".format(data["Invoice #"],data["Customer"]), (dd,mm,yy) ) 
        
        initialize_split(gnucash.GncNumeric(   10000 * float( data["Net Total"] ), 10000), acc_accounts_receivable,     trans1, "Invoice - {0} - {1}".format(data["Invoice #"],data["Customer"]) )
        initialize_split(gnucash.GncNumeric( - 10000 * float( data["Net Total"] ), 10000), acc_sales,  trans1, "Invoice - {0} - {1}".format(data["Invoice #"],data["Customer"]) )

        trans1.CommitEdit()
            
    csv_file.close()


def import_csv_transactions(csv):
    
    #Checking account
    #csv_file = file("src-data/Finalization-Transactions.csv")
    csv_file = file(csv)
    reader = DictReader(csv_file)
    for data in reader:
        
        transaction_date = time.strptime( data["Date"], "%d/%m/%Y")
        
        if transaction_date < start_date or transaction_date > end_date:
            continue
        
        dd,mm,yy = map( int, data["Date"].split("/") )
                
        if transaction_date < start_date or transaction_date > end_date:
            continue
         
        account_from    = root_account.lookup_by_name(  data["From Account"]  )
        account_to      = root_account.lookup_by_name(  data["To Account"]  )

        
        amount = float( data["Amount"] )
        
        trans1 = initialize_transaction(data["Description"], (dd,mm,yy) ) 
        initialize_split(gnucash.GncNumeric(   10000 * amount , 10000), account_from, trans1, "" )
        initialize_split(gnucash.GncNumeric( - 10000 * amount , 10000), account_to, trans1, "" )
        trans1.CommitEdit()
        
    csv_file.close()

def import_balances():
    map = { "PP Petar": "Accounts Payable", "PP Aneta": "Accounts Payable", "NP Petar": "Accounts Payable", "NP Aneta": "Accounts Payable", 
           "ITP PETAR": "Accounts Payable", "Corporation Tax Payable": "Corporation Tax Payable", "Accounts Receivable": "Accounts Receivable" ,
           "Overpaid Nics & Tax": "Accounts Receivable"}
    
    csv_file = file("src-data/balances.csv")
       
    reader = DictReader(csv_file)
    for data in reader:        
        account = data["Account"]
        if(data["Account"] in map): account = map[data["Account"]]
        account = root_account.lookup_by_name(account)
        if(not account.get_instance()): account = acc_temp
        
        trans = initialize_transaction("{0} - initial balance".format(data["Account"]), (1,1,2013) )        
        initialize_split(gnucash.GncNumeric(   100 * float(data["Balance"] ), 100) , account, trans,
                         "{0} - initial balance".format(data["Account"]) )
        initialize_split(gnucash.GncNumeric( - 100 * float(data["Balance"] ), 100), acc_opening_balances, trans,
                         "{0} - initial balance".format(data["Account"]) ) 
        
        trans.CommitEdit()
        
    csv_file.close()


import_balances()
import_payroll_csv("./src-data/petar-payroll--2013.csv", "Petar")

import_checking_account("./src-data/2013-bank-checking.csv")
import_savings_account("./src-data/2013-bank-saving.csv")
import_csv_transactions("src-data/2013-Additional-Transactions.csv")
import_csv_transactions("src-data/2013-Finalization-Transactions.csv")
#import_sales_csv("./src-data/SalesLedger-NTAGS--ARD2012.csv")

session.save()
session.end()

for f in glob.glob( "{0}/{1}.*.*".format( os.getcwd(), new_file ) ) :
    os.remove(f)
