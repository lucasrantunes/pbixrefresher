import time
import os
import sys
import argparse
import psutil
from pywinauto.application import Application
from pywinauto import timings
from time import sleep

def main():   
	# Parse arguments from cmd
	parser = argparse.ArgumentParser()
	parser.add_argument("workbook", help = "Path to .pbix file")
	parser.add_argument("--workspace", help = "name of online Power BI service work space to publish in", default = "My workspace")
	parser.add_argument("--refresh-timeout", help = "refresh timeout", default = 30000, type = int)
	parser.add_argument("--no-publish", dest='publish', help="don't publish, just save", default = True, action = 'store_false' )
	args = parser.parse_args()

	timings.after_clickinput_wait = 1
	WORKBOOK = args.workbook
	WORKSPACE = args.workspace.replace(" ", "{SPACE}") # Replace space by the equivalent key code
	WAIT_TIMEOUT = args.refresh_timeout

	# Kill running PBI
	PROCNAME = "PBIDesktop.exe"
	for proc in psutil.process_iter():
		# check whether the process name matches
		if proc.name() == PROCNAME:
			proc.kill()
	time.sleep(3)

	# Start PBI and open the workbook
	print("Starting Power BI")
	os.system('start "" "' + WORKBOOK + '"')
	sleep(1)

	# Connect pywinauto
	print("Identifying Power BI window")
	app = Application(backend = 'uia').connect(path = PROCNAME)
	win = app.window(title_re = '.*Power BI Desktop')
	timings.after_clickinput_wait = 1
	
	# Publish
	print("Waiting Power BI to load")
	publish = win.child_window(title="Publish", control_type='Button', found_index=0).wait('visible', timeout=WAIT_TIMEOUT)
	print("Publishing workbook")
	publish.click()

	publish_dialog = win.child_window(auto_id = "KoPublishToGroupDialog")
	publish_dialog.wait('visible', timeout = WAIT_TIMEOUT)

	print("Selecting workspace")
	publish_dialog.child_window(title = "Search", control_type="Edit").type_keys(WORKSPACE)
	publish_dialog.Select.click()

	print("Waiting replace screen")
	replace = win.child_window(title="Replace", control_type='Button').wait('visible', timeout=WAIT_TIMEOUT)
	replace.click()

	print("Waiting success screen")
	got_it = win.child_window(title="Got it", control_type='Button').wait('visible', timeout=WAIT_TIMEOUT)
	got_it.click()
		
	#Close
	print("Exiting")
	win.close()

	# Force close
	for proc in psutil.process_iter():
		if proc.name() == PROCNAME:
			proc.kill()

		
if __name__ == '__main__':
	try:
		main()
	except Exception as e:
		print(e)
		sys.exit(1)
