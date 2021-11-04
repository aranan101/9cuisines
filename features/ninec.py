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
    df.to_csv('./results/BOM.csv', index = False)
    print('finished parsing')



def initialise_BOM(data_root):
    print('intialising.....')
    # initialise dataset 
    primary_data = dict()
    description_data = dict()

    xml_data = open(data_root, 'r').read()  # data root was BOM.xlsx 
    root = ET.XML(xml_data)  # Parse XML
    records = root.findall('./BOMRecord')
    for record in records: 
        parent_code = record.find('./Reference').text
        components = record.findall('./BuildPackage/ComponentLine')
        for component in components: 
            child_code = component.find('./StockItemCode').text 
            child_unit = component.find('./UnitOfMeasure').text
            child_description = component.find('./Description').text
            child_quantity = float(component.find('./Quantity').text)

            if child_unit == 'gm': 
                child_quantity /= 1000
                child_unit = 'Kg'

            if child_unit == 'ml': 
                child_quantity /= 1000
                child_unit = 'Ltr'

            if parent_code in primary_data.keys(): 
                if child_code in primary_data[parent_code].keys(): 
                    primary_data[parent_code][child_code]['quantity'] += child_quantity
                else:
                    primary_data[parent_code][child_code] = {'unit': child_unit, 'quantity': child_quantity}
            else: 
                primary_data[parent_code] = {child_code: {'unit': child_unit, 'quantity': child_quantity} }
            
            if child_code in description_data.keys(): 
                pass 
            else: 
                description_data[child_code] =  child_description


    # Add Suppliers data 

    supplier_list = pd.read_csv('./database/suppliers list.csv')
    code = supplier_list.columns[0]
    food_type = supplier_list.columns[2]
    name = supplier_list.columns[3]
    supplier_list = supplier_list[[code, food_type, name]]
    supplier_list.columns = ['code', 'food type', 'name']
    supplier_data = dict()
    for i in range(supplier_list.shape[0]): 
        code = supplier_list['code'].iloc[i]
        food_type = supplier_list['food type'].iloc[i]
        name = supplier_list['name'].iloc[i]
        if code in supplier_data.keys(): 
            pass 
        else: 
            supplier_data[code] = {'food type': food_type, 'name': name }
            


    data = {'primary': primary_data, 'description': description_data, 'supplier': supplier_data}

        

    pickle_out = open("./database/database.pickle", "wb")
    pickle.dump(data, pickle_out)
    pickle_out.close()
    print('finished intialisation')



def trial_kit(input_data, df): 
	primary_data = input_data['primary']
	description_data = input_data['description']
	supplier_data = input_data['supplier']

	def query(data, primary_code, primary_quantity, primary_unit): 
		if primary_code[:3].upper() == 'ING' or primary_code[:3].upper() == 'PAC' : 
			data['Code'].append(primary_code)
			data['Quantity'].append(primary_quantity)
			data['Unit'].append(primary_unit)


		else:
			for code in primary_data[primary_code].keys(): 
				if code[:3].upper() == 'ING' or code[:3].upper() == 'PAC':
					data['Code'].append(code)
					data['Unit'].append( primary_data[primary_code][code]['unit'])
					data['Quantity'].append(primary_quantity * primary_data[primary_code][code]['quantity'] )

				else:
					query(data,code,primary_quantity * primary_data[primary_code][code]['quantity'], primary_data[primary_code][code]['unit'] )
					


   	# Query all the batches and make some output data 
	output_data = dict()
	output_data['Code'] = []
	output_data['Quantity'] = []
	output_data['Unit'] = []
	for i in range(df.shape[0]): 
	    row = df.iloc[i]
	    query(output_data, row[0].strip(), row[1], 'Each')
    # Consolidate 
	output_data = pd.DataFrame(output_data)
	consolidate = output_data.groupby(['Code','Unit']).sum().reset_index()


    # Add description 
	descriptors = dict()
	descriptors['Description'] = []
	descriptors['Food Type'] = []
	descriptors['Supplier'] = []
	descriptors['Stock'] = []
	descriptors['Order'] = []


	for i in range(consolidate.shape[0]): 
		code = consolidate['Code'].iloc[i]
		descriptors['Description'].append(description_data[code])
		try: 
			descriptors['Food Type'].append(supplier_data[code]['food type'])
			descriptors['Supplier'].append(supplier_data[code]['name'])
		except: 
			descriptors['Food Type'].append(None)
			descriptors['Supplier'].append(None)

	[descriptors['Stock'].append(None) for i in range(consolidate.shape[0])]
	[descriptors['Order'].append(None) for i in range(consolidate.shape[0])]
	descriptors = pd.DataFrame(descriptors)

    # return data 

	final_data = pd.concat([consolidate, descriptors], axis = 1)
	final_data = final_data[['Code', 'Description', 'Food Type', 'Quantity', 'Unit', 'Stock', 'Order', 'Supplier']]
	return final_data


def NAACs(filename, weekstart_raw,forecast,sheet1,sheet2,cycle_dsi,cycle_meal,cycle_quantity,cycle_spml_dsi,cycle_spml_meal,cycle_spml_quantity):

    def NAACs_parse(filename , sheet , quantity, dsi, meal , food_type ):
        data = pd.read_excel(filename, sheet_name = sheet)
        data.columns = [i for i in range(len(data.columns))]
        ## get all data with dsi codes 
        df = data.iloc[[i for i in range(len(data)) if 'DSI' in str(data[dsi].iloc[i])]]
        df = df.iloc[1: , :]
        df = df[[dsi,meal,quantity]]
        date_data = {'Cook Date': [], 'Menu Date': [], 'Chilled Raw Material Delivery Date': []}
        parse_date = dates.copy()
        ## add dates 
        for i in range(len(np.array(df.index))): 
            if i == 0: 
                date = parse_date.pop(0)
                date_data['Menu Date'].append(date)
                date_data['Chilled Raw Material Delivery Date'].append(date - datetime.timedelta(days=3))
                date_data['Cook Date'].append(date -  datetime.timedelta(days=2))

            else:
                if np.array(df.index)[i] - np.array(df.index)[i-1] > 1:
                    date = parse_date.pop(0)
                date_data['Menu Date'].append(date)
                date_data['Chilled Raw Material Delivery Date'].append(date - datetime.timedelta(days=3))
                date_data['Cook Date'].append(date -  datetime.timedelta(days=2))

        # Rename columns to correct one 
        df = df.reset_index(drop = True) 
        df.columns = ['DSI codes', 'Meal', 'Quantity']
        df = pd.concat([df, pd.DataFrame(date_data)], axis = 1)      
        df = df[df['Quantity']!= 0].reset_index(drop = True)
        type_data = pd.DataFrame({'Type':[food_type for i in range(len(df))]})
        df = pd.concat([df, type_data], axis = 1)  
        return df 

    # Consolidate 
    def Consolidate(Batch): 
        consolidate = Batch.groupby('DSI codes').sum()
        consolidate['DSI codes'] = consolidate.index
        return consolidate[['DSI codes', 'Quantity']].reset_index(drop = True)

    # Inititialise indexes
    date_index = {'Monday': 0,
     'Tuesday': 1,
     'Wednesday': 2,
     'Thursday': 3,
     'Friday': 4,
     'Saturday': 5,
     'Sunday': 6}

    

    # initialise dates and date seperators
    weekstart = datetime.datetime.strptime(weekstart_raw, '%d/%m/%Y')
    dates = [weekstart + datetime.timedelta(days=i) for i in range(0 - weekstart.weekday(), 7 - weekstart.weekday())]
    seperator = pd.to_datetime(weekstart + datetime.timedelta(days=4))

    # Create All Combined sheet 
    All_Combined = pd.concat([NAACs_parse(filename , sheet1 , cycle_quantity, cycle_dsi , cycle_meal , 'Chilled' ),
                            NAACs_parse(filename , sheet2 , cycle_spml_quantity, cycle_spml_dsi , cycle_spml_meal , 'Special')])
    All_Combined.set_index(All_Combined['Menu Date'])


    dfs = []
    df_names = []
    output_filename = './results/NAACs stock orders ' + str(int(time.time())) +'.xlsx'

    if forecast == 'all': 
        for df in [All_Combined]:
            for j in ['Chilled Raw Material Delivery Date', 'Cook Date', 'Menu Date']:
                df[j] = pd.to_datetime(df[j]).astype(str)

        # consolidate
        main_data = Consolidate(All_Combined)

        # trial kit 
        pickle_in = open(r"./database/database.pickle", "rb")
        database = pickle.load(pickle_in)
        main_tk = trial_kit(database, main_data)
       
        # Write into new xlsx file 
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            # Your loaded workbook is set as the "base of work"

            for df, name in zip([All_Combined, main_data, main_tk],
                                ['All Combined', 'Consolidate', 'trial kit']): 
                df.to_excel(writer, name, index = False)
            # Save the file
            writer.save()




    elif forecast == 'batches':
    
        # Seperate into batches 
        Batch_1 = All_Combined[All_Combined['Menu Date'] < seperator]
        Batch_2 = All_Combined[All_Combined['Menu Date'] >= seperator]


        for df in [All_Combined, Batch_1, Batch_2]:
            for j in ['Chilled Raw Material Delivery Date', 'Cook Date', 'Menu Date']:
                df[j] = pd.to_datetime(df[j]).astype(str)

        # Consolidate Batches 
        Batch1_concolidate = Consolidate(Batch_1)
        Batch2_concolidate = Consolidate(Batch_2)

        # Trial kit 
        pickle_in = open(r"./database/database.pickle", "rb")
        database = pickle.load(pickle_in)
        Batch1_tk = trial_kit(database, Batch1_concolidate)
        Batch2_tk = trial_kit(database, Batch2_concolidate)


        # Write into new xlsx file 
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            # Your loaded workbook is set as the "base of work"

            for df, name in zip([All_Combined,Batch_1, Batch1_concolidate, Batch_2, Batch2_concolidate,Batch1_tk, Batch2_tk ],
                                ['All Combined', 'Batch 1', 'Batch 1 Consolidate', 'Batch 2', 'Batch 2 Consolidate', 'Batch 1 Trial Kit', 'Batch 2 Trial Kit']): 
                df.to_excel(writer, name, index = False)
            # Save the file
            writer.save()


    else: 
        split = dates[date_index[forecast]]
        day_batch = All_Combined[All_Combined['Menu Date'] == split]

        for df in [All_Combined, day_batch]:
            for j in ['Chilled Raw Material Delivery Date', 'Cook Date', 'Menu Date']:
                df[j] = pd.to_datetime(df[j]).astype(str)

        # Consolidate Batches 
        db_consolidate = Consolidate(day_batch)

        # Trial kit 
        pickle_in = open(r"./database/database.pickle", "rb")
        database = pickle.load(pickle_in)
        db_tk = trial_kit(database, db_consolidate)


        # Write into new xlsx file 
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            # Your loaded workbook is set as the "base of work"

            # Loop through the existing worksheets in the workbook and map each title to\
            # the corresponding worksheet (that is, a dictionary where the keys are the\
            # existing worksheets' names and the values are the actual worksheets)
            
            for df, name in zip([All_Combined,day_batch,  db_consolidate, db_tk],
                                ['All Combined', forecast , forecast + ' Consolidate', forecast + ' Trial kit']): 
                df.to_excel(writer, name, index = False)

            # Save the file
            writer.save()
            
def AVANTI(filename,sheet, dsi,quantity, day): 
    output_filename = './results/AVANTI stock orders ' + str(int(time.time())) +'.xlsx'
    data = pd.read_excel(filename, sheet_name = sheet)
    data.columns = [i for i in range(len(data.columns))]
    data[day] = data[day].fillna(method="pad")
    data[day] = data[day].fillna(method="bfill")
    df = data.iloc[[i for i in range(len(data)) if 'DSI' in str(data[dsi].iloc[i])]]
    df = df[[ dsi,day,quantity]]
    df[dsi] = df[dsi].str.replace(" ","")
    df[quantity] =  df[quantity].astype(int)
    df.columns = ['DSI codes', 'day', 'Quantity']
    
    def Consolidate(Batch): 
        consolidate = Batch.groupby('DSI codes').sum()
        consolidate['DSI codes'] = consolidate.index
        return consolidate[['DSI codes', 'Quantity']].reset_index(drop = True)

    weekly_consolidate = Consolidate(df[['DSI codes', 'Quantity']])
    pickle_in = open(r"./database/database.pickle", "rb")
    database = pickle.load(pickle_in)
    weekly_tk = trial_kit(database, weekly_consolidate)

    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:

        weekly_tk.to_excel(writer, 'week trial kit', index = False)

        for day_index in df['day'].unique(): 
            day_df = Consolidate(df[['DSI codes', 'Quantity']][df['day'] == day_index])
            day_tk = trial_kit(database, day_df)
            day_tk.to_excel(writer, day_index, index = False)

        # Save the file
        writer.save()

