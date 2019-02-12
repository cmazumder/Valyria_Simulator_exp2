from pywinauto import application, timings
from time import sleep
from os import system
from setup_File import Env_variable as ConfigVariable
import pyodbc
import psutil


class Main:
	door_status = None
	door_count_from_database_start = None
	door_count_from_database_end = None

	def __int__(self):
		self.door_status = 0
		self.door_count_from_database_start = 0
		self.door_count_from_database_end = 0

	def check_if_process_is_running(self, processName):
		"""
		:param processName: Check the status of the process passed to this function
		:return: True of False
		"""
		try:
			process = psutil.win_service_get(processName)
			if process.status() == 'running':
				print '{} is {}'.format(processName, process.status())
				return True
			else:
				return False
		except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
			pass
			print 'check what?'
		return False

	def run_test_service_not_running(self, valyria_object):
		try:
			while True:
				valyria_object.toggle_valyria_door_switch()
		except KeyboardInterrupt:
			print "Door open count (Simulator) : {}".format(valyria_object.count_door_open)
			pass

	def run_test_service_running(self, valyria_object, database_object):
		self.door_count_from_database_start = database_object.get_LogicCageOpen_count()
		try:
			while True:
				valyria_object.open_valyria_door_switch()
				if database_object.get_LogicCageOpen_status() == 1:
					valyria_object.close_valyria_door_switch()
				elif database_object.get_LogicCageOpen_status() == 0:
					print 'Door status is already closed \nIssue with Valyria driver and database sync'
					exit(0)
				else:
					print 'Issue is database for door status'
					exit(0)

		except KeyboardInterrupt:
			self.door_count_from_database_end = database_object.get_LogicCageOpen_count()
			print "Door open count (Simulator) : {}".format(valyria_object.count_door_open)
			print "Door open value_start (database) : {}".format(self.door_count_from_database_start)
			print "Door open value_end (database) : {}".format(self.door_count_from_database_end)
			print "Door open count (database) : {}".format(self.door_count_from_database_end
			                                               - self.door_count_from_database_start)
			pass


class Database:
	database_con = None
	database_cursor = None
	db_name = None

	def __init__(self, db_name):
		""" constructor"""
		self.db_name = db_name
		self.database_connection()

	def database_connection(self):
		"""
		Connect to database via prameters in configuration from imported setup_File
		"""
		try:
			connection_string = r"DRIVER={{SQL Server}};SERVER={0}; database={1}; trusted_connection=yes;UID={2};PWD={3}".format(
				ConfigVariable.get("db_server_local"), self.db_name,
				ConfigVariable.get("db_username"),
				ConfigVariable.get("db_password"))

			self.database_con = pyodbc.connect(connection_string)
			self.database_cursor = self.database_con.cursor()
		except pyodbc.Error as E:
			print 'Cannot setup database connection to %s \n %s', self.db_name, E.message

	def __del__(self):
		""" destructor"""
		self.database_con.close()

	def execute_sql_query(self, sql_query):
		"""
		Query database and fetch data
		:param sql_query: Sql text query
		:return: data
		"""

		try:
			if self.database_con:
				self.database_cursor.execute(sql_query)
				data = self.database_cursor.fetchall()
				return data
		except pyodbc.Error as E:
			print 'Problem with objects\'s misc_cursor in execute query \n %s', E.args
		return None

	def get_LogicCageOpen_count(self):
		sql_query = r"SELECT * FROM Vertex.dbo.ControllerMeter where MeterName='LogicCageOpen'"
		result_from_query = self.execute_sql_query(sql_query=sql_query)
		if result_from_query:
			value_from_list = [x[4] for x in result_from_query]
			return value_from_list[0]
		else:
			print 'Logic cage open count not available in database'
		return None

	def get_LogicCageOpen_status(self):
		sql_query = r"SELECT * FROM Vertex.dbo.ControllerMeter where MeterName='LogicCageOpenState'"
		result_from_query = self.execute_sql_query(sql_query=sql_query)
		if result_from_query:
			value_from_list = [x[4] for x in result_from_query]
			return value_from_list[0]
		else:
			print 'Logic cage status not available in database'
		return None


class ValyriaAutomate:
	valyria_application = None
	valyria_handler = None
	valyria_dialog = None
	count_door_open = None

	def __init__(self):
		self.valyria_application = None
		self.terminate_valyria()
		self.start_valyria()
		self.count_door_open = 0

	def __del__(self):
		# self.terminate_valyria()
		print '\ndestroy'

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
			self.valyria_application.start(ConfigVariable.get("valyria_app_path"))
			# valyria_app = Application(backend='uia').connect(path='ValyriaTray.exe', title='Valyria Simulator')
			self.check_application_started()

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
			# This draw (default colour is green) outline at the control if found
			# self.valyria_application.Dialog.Custom3.Button12.draw_outline()
			self.valyria_application.Dialog.Custom3.Button12.click_input()
		except Exception as e:
			print e.message

	def close_valyria_door_switch(self):
		self.click_is_door_open()

	def open_valyria_door_switch(self):
		self.click_is_door_open()
		self.count_door_open += 1

	def toggle_valyria_door_switch(self):
		self.open_valyria_door_switch()
		self.closde_valyria_door_switch()


def check_if_process_is_running(processName):
	"""
	:param processName: Check the status of the process passed to this function
	:return: True of False
	"""
	try:
		process = psutil.win_service_get(processName)
		if process.status() == 'running':
			print '{} is {}'.format(processName, process.status())
			return True
		else:
			return False
	except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
		pass
		print 'check what?'
	return False


if __name__ == '__main__':
	valyria = ValyriaAutomate()
	# valyria.print_identifiers()
	test = Main()
	aristocrat_vertex_process = ConfigVariable.get("controller_process")
	dummy_process = r'lfsvc'
	if test.check_if_process_is_running(dummy_process):
		vertex_database = Database('Vertex')
		if vertex_database:
			print "Connected to database"
			test.run_test_service_running(valyria_object=valyria, database_object=vertex_database)
		else:
			print 'Could not connect to database'
	else:
		print '{} is stopped. Running test without ControllerMeter values'.format(aristocrat_vertex_process)
		test.run_test_service_not_running(valyria_object=valyria)
