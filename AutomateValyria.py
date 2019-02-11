from pywinauto import application, timings, findwindows
from time import sleep
from os import system


class ValyriaAutomate:
	valyria_application = None
	valyria_handler = None
	valyria_dialog = None

	def __init__(self):
		self.valyria_application = None
		self.terminate_valyria()
		self.start_valyria()

	def __del__(self):
		# self.terminate_valyria()
		print '\n\ndestroy'

	def check_application_started(self):
		"""
		# wait a maximum of 10.5 seconds for the
		# window to be found in increments of .5 of a second.
		# P.int a message and re-raise the original exception if never found.
		return self.valyria_application.findwindows.find_windows(title=u'Valyria Simulator')[0]
		timings.wait_until_passes(20, 0.5, check)  # Important: 'check' without brackets
		"""
		# describe the window inside Notepad.exe process
		dlg_spec = self.valyria_application['Valyria Simulator']
		# wait till the window is really open
		self.valyria_dialog = dlg_spec.wait('visible')

	def start_valyria(self):
		try:

			# window = timings.wait_until_passes(10.5, .5, Exists, (ElementNotFoundError))
			# window = timings.WaitUntilPasses(10, 0.5, lambda: Application.window_(title=u'Valyria Simulator'))
			# Open "Control Panel"
			self.valyria_application = application.Application(backend="uia")
			self.valyria_application.start(r'C:\Valyria\ValyriaTray.exe')
			# valyria_app = Application(backend='uia').connect(path='ValyriaTray.exe', title='Valyria Simulator')
			got_it = self.check_application_started()
			self.valyria_handler = self.valyria_application[u'Valyria Simulator']
			self.valyria_handler.wait('ready')
		except timings.TimeoutError as exc:
			print("timed out")
			print exc

	def terminate_valyria(self):
		if self.valyria_application:
			self.valyria_application.kill()
		else:
			system("TASKKILL /F /IM ValyriaTray.exe")

	def print_identifiers(self):
		self.valyria_application.window(title=u'Valyria Simulator').print_control_identifiers()

	# self.valyria_dialog.window().print_control_identifiers()

	def click_is_door_open(self):
		try:
			sleep(3)
			self.valyria_dialog.set_focus()
			sleep(1)
			# This draw (default colour is green) outline at the control if found
			# self.valyria_application.Dialog.Custom3.Button12.draw_outline()
			self.valyria_application.Dialog.Custom3.Button12.click_input()
		except Exception as e:
			print e.message


if __name__ == '__main__':
	valyria = ValyriaAutomate()
	# valyria.print_identifiers()
	valyria.click_is_door_open()
