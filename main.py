import openpyxl as xl

FILENAME = "testing.xlsx"

def create_excel(name):

	# NOTE: creates a new workbook with atleast one worksheet
	wb = xl.Workbook() 

	# Gives you the active worksheet... ?? (need to check what happens
	# when more than one worksheet is present in the workbook)
	# Apparently always gives the FIRST worksheet... set to 0 by default
	ws = wb.active

	# What happens if you don't save the workbook? It just 
	# automatically deletes or disappears? TBC...
	# Also note that you are saving the WORKBOOK, so it
	# makes a lot of sense to save from the workbook obj
	wb.save(FILENAME)

def load_file(name):

	test = xl.load_workbook(name)
	ws = test.active
	# I need a way to search for a cell... so, i can do that
	# using VBA, can I then access these buttons 


	test.save()





if __name__ == "__main__":
	#create_excel("testing")
	load_file(FILENAME)