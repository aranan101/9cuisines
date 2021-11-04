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

## Automation classes 
class NAACs_plane: 


	def __init__(self): 
		naacs_frame = LabelFrame(root, text = 'Automate the NAACs dsi to stock order process', padx = 5, pady = 5)
		naacs_frame.pack(padx = 10, pady = 10)

		naacs_upload = Button(naacs_frame, text = 'Upload the NAACs file in an xlsx format', command = self.upload_NAACs)
		naacs_upload.pack()
		self.naacs_frame = naacs_frame


	def upload_NAACs(self): 
		file_types = [('xlsx files', '.xlsx')]
		self.fname = select_file(file_types)
		naacs_frame = self.naacs_frame
		xl = pd.ExcelFile(self.fname)

		wr_label = Label(naacs_frame, text = 'What is the starting date of stock order in dd/mm/yyyy?')
		wr_label.pack()
		self.wr = Entry(naacs_frame)
		self.wr.pack()

		fc_label = Label(naacs_frame, text = 'What type of forecast do you want to do?')
		fc_label.pack()
		self.clicked = StringVar()
		self.clicked.set('batches')
		clicked = self.clicked
		fc = OptionMenu(naacs_frame, clicked, "Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday", 'batches', 'all')
		fc.pack()

		s_label = Label(naacs_frame, text = 'What is the name of the sheet with the chilled meals? Make sure the spelling is accurate')
		s_label.pack()
		self.s = StringVar()
		self.s.set(xl.sheet_names[0])
		s = self.s
		sh1 = OptionMenu(naacs_frame, s, *xl.sheet_names)
		sh1.pack()

		s2_label = Label(naacs_frame, text = 'What is the name of the sheet with the special meals? Make sure the spelling is accurate')
		s2_label.pack()
		self.s2 = StringVar()
		self.s2.set(xl.sheet_names[0])
		s2 = self.s2
		sh2 = OptionMenu(naacs_frame, s2, *xl.sheet_names)
		sh2.pack()

		cdsi_label = Label(naacs_frame, text = 'What letter column for the chilled foods are the dsi codes in?')
		cdsi_label.pack()
		self.cdsi = Entry(naacs_frame)
		self.cdsi.pack()

		cmeal_label = Label(naacs_frame, text = 'What letter column for the chilled foods are the meal descriptions in?')
		cmeal_label.pack()
		self.cmeal = Entry(naacs_frame)
		self.cmeal.pack()

		cq_label = Label(naacs_frame, text = 'What letter column for the chilled foods are the quantity descriptions in?')
		cq_label.pack()
		self.cq = Entry(naacs_frame)
		self.cq.pack()

		csdsi_label = Label(naacs_frame, text = 'What letter column for the special foods are the dsi codes in?')
		csdsi_label.pack()
		self.csdsi = Entry(naacs_frame)
		self.csdsi.pack()

		csmeal_label = Label(naacs_frame, text = 'What letter column for the special foods are the meal descriptions in?')
		csmeal_label.pack()
		self.csmeal = Entry(naacs_frame)
		self.csmeal.pack()

		csq_label = Label(naacs_frame, text = 'What letter column for the special foods are the quantity descriptions in?')
		csq_label.pack()
		self.csq = Entry(naacs_frame)
		self.csq.pack()

		naacs_button = Button(naacs_frame, text = 'Finish selection', command = self.NAAC_stock_order)
		naacs_button.pack()

	def NAAC_stock_order(self):
		alphabet = list(string.ascii_lowercase)
		alphabet_index = {letter:i for i, letter in enumerate(alphabet)}
		naacs_frame = self.naacs_frame 
		NAACs(
	    	self.fname,
			self.wr.get().strip(), 
			self.clicked.get().strip(),
			self.s.get().strip(),
			self.s2.get().strip(),
			alphabet_index[self.cdsi.get().lower().strip()],
			alphabet_index[self.cmeal.get().lower().strip()],
			alphabet_index[self.cq.get().lower().strip()],
			alphabet_index[self.csdsi.get().lower().strip()],
			alphabet_index[self.csmeal.get().lower().strip()],
			alphabet_index[self.csq.get().lower().strip()])
		naacs_label = Label(naacs_frame, text = 'NAACs stock order list has been outputted successfully in the results folder')
		naacs_label.pack()

class AVANTI_plane:

	def __init__(self): 
		avanti_frame = LabelFrame(root, text = 'Automate the AVANTI dsi to stock order process', padx = 5, pady = 5)
		avanti_frame.pack(padx = 10, pady = 10)

		avanti_upload = Button(avanti_frame, text = 'Upload the AVANTI file in an xlsx format', command = self.upload_avanti)
		avanti_upload.pack()
		self.avanti_frame = avanti_frame

	def upload_avanti(self): 
		file_types = [('xlsx files', '.xlsx')]
		self.fname = select_file(file_types)
		avanti_frame = self.avanti_frame
		xl = pd.ExcelFile(self.fname)

		s_label = Label(avanti_frame, text = 'What is the name of the sheet for forecasting?')
		s_label.pack()
		self.s = StringVar()
		self.s.set(xl.sheet_names[0])
		s = self.s
		sh1 = OptionMenu(avanti_frame, s, *xl.sheet_names)
		sh1.pack()


		cdsi_label = Label(avanti_frame, text = 'What letter column for the dsi codes in?')
		cdsi_label.pack()
		self.cdsi = Entry(avanti_frame)
		self.cdsi.pack()

		cday_label = Label(avanti_frame, text = 'What letter column does the day in "MON" format stay in?')
		cday_label.pack()
		self.cday = Entry(avanti_frame)
		self.cday.pack()

		cq_label = Label(avanti_frame, text = 'What letter column are the quantity descriptions in?')
		cq_label.pack()
		self.cq = Entry(avanti_frame)
		self.cq.pack()

		avanti_button = Button(avanti_frame, text = 'Finish selection', command = self.avanti_stock_order)
		avanti_button.pack()

	def avanti_stock_order(self): 
		alphabet = list(string.ascii_lowercase)
		alphabet_index = {letter:i for i, letter in enumerate(alphabet)}
		avanti_frame = self.avanti_frame 
		AVANTI(
			self.fname,
			self.s.get(), 
			alphabet_index[self.cdsi.get().lower().strip()],
			alphabet_index[self.cq.get().lower().strip()], 
			alphabet_index[self.cday.get().lower().strip()]
			)
		avanti_label = Label(avanti_frame, text = 'avanti stock order list has been outputted successfully in the results folder')
		avanti_label.pack()







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


# Automation frames 
NAACs_plane()
AVANTI_plane()

# Run GUI
root.mainloop()