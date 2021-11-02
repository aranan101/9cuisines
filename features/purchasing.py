from ninec import *
from tkinter import * 
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import datetime

# Helper functions

def select_file(file_types):
    
    filename = fd.askopenfilename(
        title='upload file',
        initialdir='./',
        filetypes=file_types)

    return filename

def NAAC_stock_order():
	file_types = [('xlsx files', '.xlsx')]
	fname = select_file(file_types)
	alphabet = list(string.ascii_lowercase)
	alphabet_index = {letter:i for i, letter in enumerate(alphabet)}
	NAACs(
    	fname,
		wr.get().strip(), 
		clicked.get().strip(),
		s.get().strip(),
		s2.get().strip(),
		alphabet_index[cdsi.get().lower().strip()],
		alphabet_index[cmeal.get().lower().strip()],
		alphabet_index[cq.get().lower().strip()],
		alphabet_index[csdsi.get().lower().strip()],
		alphabet_index[csmeal.get().lower().strip()],
		alphabet_index[csq.get().lower().strip()])
	naacs_label = Label(naacs_frame, text = 'NAACs stock order list has been outputted successfully in the results folder')
	naacs_label.pack()

def do_trial_kit():
	file_types = [('xlsx files', '.xlsx')]
	fname = select_file(file_types)
	df = pd.read_excel(fname, sheet_name = 'Sheet1')
	pickle_in = open("./database/database.pickle", "rb")
	database = pickle.load(pickle_in)
	X = trial_kit(database, df)
	X.to_csv('./results/trial kit ' + str(int(time.time())) + '.csv', index = False)
	trial_kit_label = Label(tk_frame, text = 'Trial kitting has been outputted successfully in the results folder') 
	trial_kit_label.pack()

def do_BOM_initialise():
	file_types = [('XML files', '.xml')]
	fname = select_file(file_types)
	initialise_BOM(fname)
	finished_bom_label = Label(bom_frame, text = 'database has been outputted successfully in the database folder') 
	finished_bom_label.pack()

# Making the graphical interface


root = Tk()
root.title('9Cuisines Purchasing')

# Initialise database frame

bom_frame = LabelFrame(root, text = 'Initialise trial kitting database', padx = 5, pady = 5)
bom_frame.pack(padx = 10, pady = 10)

bom_txt = """import BOM materials as xml file
 make sure up to date suppliers list is saved as csv file in database folder with name "suppliers list"""
bom_button = Button(bom_frame, text = bom_txt, anchor = 'w',  command = do_BOM_initialise )
bom_button.pack()


# Trial kitting frame

tk_frame = LabelFrame(root, text = 'Trial kitting multiple dsi codes', padx = 5, pady = 5)
tk_frame.pack(padx = 10, pady = 10)

tk_button = Button(tk_frame, text = 'import file of DSI codes and quantites (layout of file is given in Layouts directory)', command = do_trial_kit )
tk_button.pack()



# NAACs frame 

naacs_frame = LabelFrame(root, text = 'Automate the NAACs dsi to stock order process', padx = 5, pady = 5)
naacs_frame.pack(padx = 10, pady = 10)

wr_label = Label(naacs_frame, text = 'What is the starting date of stock order in dd/mm/yyyy?')
wr_label.pack()
wr = Entry(naacs_frame)
wr.pack()

fc_label = Label(naacs_frame, text = 'What type of forecast do you want to do?')
fc_label.pack()
clicked = StringVar()
clicked.set('batches')
fc = OptionMenu(naacs_frame, clicked, "Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday", 'batches', 'all')
fc.pack()

s_label = Label(naacs_frame, text = 'What is the name of the sheet with the chilled meals? Make sure the spelling is accurate')
s_label.pack()
s = Entry(naacs_frame)
s.pack()

s2_label = Label(naacs_frame, text = 'What is the name of the sheet with the special meals? Make sure the spelling is accurate')
s2_label.pack()
s2 = Entry(naacs_frame)
s2.pack()

cdsi_label = Label(naacs_frame, text = 'What letter column for the chilled foods are the dsi codes in?')
cdsi_label.pack()
cdsi = Entry(naacs_frame)
cdsi.pack()

cmeal_label = Label(naacs_frame, text = 'What letter column for the chilled foods are the meal descriptions in?')
cmeal_label.pack()
cmeal = Entry(naacs_frame)
cmeal.pack()

cq_label = Label(naacs_frame, text = 'What letter column for the chilled foods are the quantity descriptions in?')
cq_label.pack()
cq = Entry(naacs_frame)
cq.pack()

csdsi_label = Label(naacs_frame, text = 'What letter column for the special foods are the dsi codes in?')
csdsi_label.pack()
csdsi = Entry(naacs_frame)
csdsi.pack()

csmeal_label = Label(naacs_frame, text = 'What letter column for the special foods are the meal descriptions in?')
csmeal_label.pack()
csmeal = Entry(naacs_frame)
csmeal.pack()

csq_label = Label(naacs_frame, text = 'What letter column for the special foods are the quantity descriptions in?')
csq_label.pack()
csq = Entry(naacs_frame)
csq.pack()

naacs_button = Button(naacs_frame, text = 'Finish selection and upload the NAACs file in an xlsx format', command = NAAC_stock_order)
naacs_button.pack()



root.mainloop()