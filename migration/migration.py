import pandas as pd 
import numpy as np 
import pickle 
import csv
import xml.etree.ElementTree as ET
import openpyxl as pxl
import string 
import time
import datetime


def BOM_to_csv(data_root): 
    print('start parsing....')
    data = {
        'BOM code': [],
        'Stock Code': [], 
        'BOM description': [], 
        'Stock description': [], 
        'Sequence': [], 
        'Quantity': [], 
        'Unit of Measure': [], 
        'Scrap Percentage': [], 
        'Fixed Quantity': []
    }

    xml_data = open(data_root, 'r').read()  # data root was BOM.xlsx 
    root = ET.XML(xml_data)  # Parse XML
    records = root.findall('./BOMRecord')
    for record in records: 
        parent_code = record.find('./Reference').text
        parent_description = record.find('./Description').text
        components = record.findall('./BuildPackage/ComponentLine')
        for component in components: 
            child_code = component.find('./StockItemCode').text 
            child_unit = component.find('./UnitOfMeasure').text
            child_description = component.find('./Description').text
            child_quantity = float(component.find('./Quantity').text)
            scrap_percentage = component.find('./ScrapPercentage').text
            sequence_number = component.find('./SequenceNumber').text
            fixed_quantity = component.find('./FixedQuantity').text

            if child_unit == 'gm': 
                child_quantity /= 1000
                child_unit = 'Kg'

            if child_unit == 'ml': 
                child_quantity /= 1000
                child_unit = 'Ltr'
            
            data['BOM code'].append(parent_code)
            data['Stock Code'].append(child_code)
            data['BOM description'].append(parent_description)
            data['Stock description'].append(child_description)
            data['Sequence'].append(sequence_number)
            data['Quantity'].append(child_quantity)
            data['Unit of Measure'].append(child_unit)
            data['Scrap Percentage'].append(scrap_percentage)
            data['Fixed Quantity'].append(fixed_quantity)

    df = pd.DataFrame(data)
    df.to_csv('./migrated data/BOM/BOM.csv', index = False)
    print('finished parsing')


def nominal_accounts_clean(read_only = False): 
	nom_acc_file = './pre-migration/nominal accounts.xlsx'
	nom_acc = pd.read_excel(nom_acc_file, sheet_name = 'SageReportData1')
	nom_acc.columns = nom_acc.iloc[0]
	nom_acc = nom_acc[['NominalRecord.AccountReference', 'NominalRecord.AccountName']].iloc[1:]

	###
	account_ref_desc = pd.concat([pd.read_csv('./migrated data/Nominal/Predefined Report Categories.csv'),
	                               pd.read_csv('./migrated data/Nominal/ReportCategories CSV.csv')])
	###
	account_ref_index = dict()
	group_ref =[ 
	    [(1100, 1290), 110],[(2000, 2350), 200],[(3000,3190), 300],[(4000,4199), 400],[ (4900, 4906), 490],[(5000, 5110), 1001],
	    [(5120, 5222), 1002],[(6000, 6050), 1003],[(6051, 6099), 1004],[(6100, 6250), 1005],[(6251, 6260), 1004],
	    [(6310, 6900), 1006],[(7000, 9999), 1007], [(1001,1005), 100]
	]
	ind_ref = [(3200, 320),( 4200, 420),( 4400, 440 ),(10, 1), (11,1), (12,1), (20,2), (21,3), (30, 4), (31, 5), (40, 6), (41,7),
	          (50,8), (51,9)]

	for batch_ref in group_ref:
	    for i in range(batch_ref[0][0], batch_ref[0][1]+1): 
	        account_ref_index[i] = batch_ref[1]

	for ref in ind_ref: 
	    account_ref_index[ref[0]] = ref[1]
	    
	acc_des_index = dict()
	for j in range(account_ref_desc.shape[0]) : 
	    Code = account_ref_desc['Code'].iloc[j]
	    Description = account_ref_desc['Description'].iloc[j]
	    Type = account_ref_desc['Type'].iloc[j]
	    CategoryType = account_ref_desc['CategoryType'].iloc[j]
	    
	    acc_des_index[Code] = {'Description':Description, 
	                          'Type': Type,
	                          'CategoryType': CategoryType}
	    

	if read_only == False: 
		nominal_accounts = {
		'AccountNumber': [],
		'AccountCostCentre': [], 
		'AccountDepartment': [], 
		'AccountName': [],
		'AccountReportCategory': [],
		'AllowJournalsToBePosted': [], 
		'Memo': [], 
		'ActiveStatus': []
		}

		for i in range(nom_acc.shape[0]):
			code = nom_acc['NominalRecord.AccountReference'].iloc[i] 
			name = nom_acc['NominalRecord.AccountName'].iloc[i]
			nominal_accounts['AccountNumber'].append(code)
			acc_code = account_ref_index[int(code)]
			nominal_accounts['AccountName'].append(name)
			nominal_accounts['AccountReportCategory'].append(acc_code)
		for category in ['AccountCostCentre', 'AccountDepartment', 'AllowJournalsToBePosted', 'Memo','ActiveStatus' ]:
			[nominal_accounts[category].append(None) for i in range(nom_acc.shape[0])]

		pd.DataFrame(nominal_accounts).to_csv('./migrated data/Nominal/nominal codes.csv', index = False)


	    
	else: 
	    nominal_accounts = {
	        'Nominal Code': [],
	        'Nominal Name':[],
	        'Account Report Code': [], 
	        'Account Report Name': [], 
	        'Account Report Type': [], 
	        'Account Report Category Type' : []
	    }
	    
	    for i in range(nom_acc.shape[0]):
	        code = nom_acc['NominalRecord.AccountReference'].iloc[i] 
	        name = nom_acc['NominalRecord.AccountName'].iloc[i]
	        nominal_accounts['Nominal Code'].append(code)
	        acc_code = account_ref_index[int(code)]
	        nominal_accounts['Nominal Name'].append(name)
	        nominal_accounts['Account Report Code'].append(acc_code)
	        nominal_accounts['Account Report Name'].append(acc_des_index[acc_code]['Description'])
	        nominal_accounts['Account Report Type'].append(acc_des_index[acc_code]['Type'])
	        nominal_accounts['Account Report Category Type'].append(acc_des_index[acc_code]['CategoryType'])
	    pd.DataFrame(nominal_accounts).to_csv('./migrated data/Nominal/nominal codes (read only).csv', index = False)


def supplier_accounts_clean(): 
	sup_acc = pd.read_excel('./pre-migration/supplier accounts.xlsx', 'SageReportData1')
	sup_acc.columns = sup_acc.iloc[0]
	NecessaryColumns =['SupplierRecord.AccountReference', 'SupplierRecord.AccountName', 'SupplierRecord.CreditLimit', 'SupplierRecord.VATRegistrationNumber',
	 'SupplierRecord.DefaultNominal','SupplierRecord.Telephone','SupplierRecord.ContactName','SupplierRecord.Email1','SupplierRecord.BankSortCode', 
	 'SupplierRecord.BankAccountNumber'  ]
	sup_acc = sup_acc[NecessaryColumns].iloc[1:]
	sup_acc[['SupplierRecord.BankSortCode']] = sup_acc[['SupplierRecord.BankSortCode']].astype(str)

	acc_columns = ['AccountNumber', 'AccountName', 'ShortName', 'AccountBalance', 'CreditLimit', 'CurrencyISOCode', 'SYSExchangeRateType', 'CountryCode',
	 'TaxCode', 'TaxRegistrationNumber', 'DefaultOrderPriority', 'DefaultNominalAccountNumber', 'DefaultNominalCostCentre', 'DefaultNominalDepartment', 
	 'PLPaymentGroup', 'EarlySettlementDiscountPercent', 'DaysEarlySettlementDiscountApplies', 'PaymentTermsInDays', 'SYSPaymentTermsBasis', 'AccountIsOnHold', 
	 'ValueOfCurrentOrdersInPOP', 'DateOfLastTransaction', 'EuroAccountNumberCopiedFromTo', 'DateEuroAccountCopied', 'MainTelephoneCountryCode', 
	 'MainTelephoneAreaCode', 'MainTelephoneSubscriberNumber', 'MainFaxCountryCode', 'MainFaxAreaCode', 'MainFaxSubscriberNumber', 'MainWebsite', 
	 'AddressLine1', 'AddressLine2', 'AddressLine3', 'AddressLine4', 'City', 'County', 'Country', 'PostCode', 'ContactSalutation', 'ContactFirstName', 
	 'ContactMiddleName', 'ContactLastName', 'ContactTelephoneCountryCode', 'ContactTelephoneAreaCode', 'ContactTelephoneSubscriberNumber', 'ContactFaxCountryCode',
	  'ContactFaxAreaCode', 'ContactFaxSubscriberNumber', 'ContactMobileCountryCode', 'ContactMobileAreaCode', 'ContactMobileSubscriberNumber', 'ContactWebsite', 'ContactEmail', 
	  'TradingTerms', 'CreditReference', 'CreditBureau', 'CreditPosition', 'TermsAgreed', 'AccountOpened', 'LastCreditReview', 'NextCreditReview', 'ApplicationDate', 'DateReceived', 
	  'BankSortCode', 'BankAccountNumber', 'BankAccountName', 'BankPaymentReference', 'BankIBANNumber', 'BankBICNumber', 'BankRollNumber', 'BankBACSReference',
	   'BankAdditionalReference', 'BankNonUKSortCode', 'Memo', 'ActiveStatus', ]


	## Make supp balance keys 

	sup_bal = pd.read_excel('./pre-migration/supplier balance.xlsx', 'Sheet1')
	sup_bal = sup_bal[sup_bal['A/C'].notna()]
	balance_index = dict()
	for i in range(sup_bal.shape[0]): 
		balance_index[sup_bal['A/C'].iloc[i]] = sup_bal['Balance'].iloc[i] 

	## Start adding each of the fields to the right column for each supplier 
	supplier_accounts = {category: [] for category in acc_columns}
	for i in range(sup_acc.shape[0]):
		code = sup_acc['SupplierRecord.AccountReference'].iloc[i]
		supplier_accounts['AccountNumber'].append(code)
		supplier_accounts['AccountName'].append(sup_acc['SupplierRecord.AccountName'].iloc[i])
		if code in balance_index.keys(): 
			supplier_accounts['AccountBalance'].append(balance_index[code])
		else: 
			supplier_accounts['AccountBalance'].append(0)
		supplier_accounts['CreditLimit'].append(sup_acc['SupplierRecord.CreditLimit'].iloc[i])
		supplier_accounts['CurrencyISOCode'].append('GBP')
		supplier_accounts['TaxRegistrationNumber'].append(sup_acc['SupplierRecord.VATRegistrationNumber'].iloc[i])
		supplier_accounts['DefaultNominalAccountNumber'].append(sup_acc['SupplierRecord.DefaultNominal'].iloc[i])
		supplier_accounts[ 'MainTelephoneSubscriberNumber'].append(sup_acc['SupplierRecord.Telephone'].iloc[i])
		supplier_accounts['ContactFirstName'].append(sup_acc['SupplierRecord.ContactName'].iloc[i])
		supplier_accounts[ 'ContactEmail'].append(sup_acc['SupplierRecord.Email1'].iloc[i])
		if sup_acc['SupplierRecord.BankSortCode'].iloc[i] == 'nan': 
			sortcode = None
		else:
			sortcode = sup_acc['SupplierRecord.BankSortCode'].iloc[i].replace('-','').replace(' ','')
		supplier_accounts['BankSortCode' ].append(sortcode)
		account_number = str(sup_acc['SupplierRecord.BankAccountNumber'].iloc[i]).replace(' ','')
		if account_number == 'nan':
			supplier_accounts['BankAccountNumber'].append(None)
		else:
			supplier_accounts['BankAccountNumber'].append(account_number)

	## Rest of the columns that haven't been added to you fill with NA values 

	for key in supplier_accounts.keys(): 
	    if len(supplier_accounts[key]) == 0: 
	        [supplier_accounts[key].append(None) for i in range(sup_acc.shape[0])]

	## Create a csv output 
	pd.DataFrame(supplier_accounts).to_csv('./migrated data/Supplier/supplier accounts.csv', index = False)



def customer_accounts_clean(): 
	cus_acc = pd.read_excel('./pre-migration/customer accounts.xlsx', 'SageReportData1')
	cus_acc.columns = cus_acc.iloc[0]
	NecessaryColumns =  ['CustomerRecord.AccountReference', 'CustomerRecord.AccountName','CustomerRecord.CreditLimit','CustomerRecord.AddressLine1', 'CustomerRecord.AddressLine2',
	'CustomerRecord.DefaultNominal','CustomerRecord.AddressLine3', 'CustomerRecord.AddressLine4','CustomerRecord.AddressLine5','CustomerRecord.Telephone' ]
	cus_acc = cus_acc[NecessaryColumns].iloc[1:]

	acc_columns = ['AccountNumber', 'AccountName', 'ShortName', 'AccountBalance', 'CreditLimit', 'CurrencyISOCode', 'SYSExchangeRateType', 'CountryCode', 'TaxCode',
	 'TaxRegistrationNumber', 'DefaultOrderPriority', 'DefaultNominalAccountNumber', 'DefaultNominalCostCentre', 'DefaultNominalDepartment', 'EarlySettlementDiscountPercent', 
	 'DaysEarlySettlementDiscountApplies', 'PaymentTermsInDays', 'SYSPaymentTermsBasis', 'InvoiceLineDiscountPercent', 'InvoiceDiscountPercent', 'AccountIsOnHold', 
	 'ValueOfCurrentOrdersInSOP', 'DateOfLastTransaction', 'EuroAccountNumberCopiedFromTo', 'DateEuroAccountCopied', 'MainTelephoneCountryCode', 'MainTelephoneAreaCode', 
	 'MainTelephoneSubscriberNumber', 'MainFaxCountryCode', 'MainFaxAreaCode', 'MainFaxSubscriberNumber', 'MainWebsite', 'AddressLine1', 'AddressLine2', 'AddressLine3', 'AddressLine4',
	  'City', 'County', 'Country', 'PostCode', 'ContactSalutation', 'ContactFirstName', 'ContactMiddleName', 'ContactLastName', 'ContactTelephoneCountryCode', 'ContactTelephoneAreaCode',
	   'ContactTelephoneSubscriberNumber', 'ContactFaxCountryCode', 'ContactFaxAreaCode', 'ContactFaxSubscriberNumber', 'ContactMobileCountryCode', 'ContactMobileAreaCode', 
	   'ContactMobileSubscriberNumber', 'ContactWebsite', 'ContactEmail', 'TradingTerms', 'CreditReference', 'CreditBureau', 'CreditPosition', 'TermsAgreed', 'AccountOpened', 
	   'LastCreditReview', 'NextCreditReview', 'ApplicationDate', 'DateReceived', 'Memo', 'ActiveStatus']


	## Make supp balance keys 

	cus_bal = pd.read_excel('./pre-migration/customer balance.xlsx', 'Sheet1')
	cus_bal = cus_bal[cus_bal['A/C'].notna()]
	balance_index = dict()
	for i in range(cus_bal.shape[0]): 
		balance_index[cus_bal['A/C'].iloc[i]] = cus_bal['Balance'].iloc[i] 

	## Start adding each of the fields to the right column for each supplier 
	customer_accounts = {category: [] for category in acc_columns}
	for i in range(cus_acc.shape[0]):
		code = cus_acc['CustomerRecord.AccountReference'].iloc[i]
		customer_accounts['AccountNumber'].append(code)
		customer_accounts['AccountName'].append(cus_acc['CustomerRecord.AccountName'].iloc[i])
		if code in balance_index.keys(): 
			customer_accounts['AccountBalance'].append(balance_index[code])
		else: 
			customer_accounts['AccountBalance'].append(0)
		customer_accounts['CreditLimit'].append(cus_acc['CustomerRecord.CreditLimit'].iloc[i])
		customer_accounts['DefaultNominalAccountNumber'].append(cus_acc['CustomerRecord.DefaultNominal'].iloc[i])
		customer_accounts['MainTelephoneSubscriberNumber'].append(cus_acc['CustomerRecord.Telephone'].iloc[i])
		customer_accounts['AddressLine1'].append(cus_acc['CustomerRecord.AddressLine1'].iloc[i])
		customer_accounts['AddressLine2'].append(cus_acc['CustomerRecord.AddressLine2'].iloc[i])
		customer_accounts['AddressLine3'].append(cus_acc['CustomerRecord.AddressLine3'].iloc[i])
		customer_accounts['AddressLine4'].append(cus_acc['CustomerRecord.AddressLine4'].iloc[i])
		customer_accounts['PostCode'].append(cus_acc['CustomerRecord.AddressLine5'].iloc[i])
		customer_accounts['CurrencyISOCode'].append('GBP')
		
	## Rest of the columns that haven't been added to you fill with NA values 

	for key in customer_accounts.keys(): 
	    if len(customer_accounts[key]) == 0: 
	        [customer_accounts[key].append(None) for i in range(cus_acc.shape[0])]

	## Create a csv output 
	pd.DataFrame(customer_accounts).to_csv('./migrated data/Customer/customer accounts.csv', index = False)
