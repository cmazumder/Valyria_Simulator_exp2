from pywinauto import application, timings
from time import sleep
from datetime import datetime
from os import system
from setup_File import Env_variable as ConfigVariable
import pyodbc
import psutil
import logging
from logging.handlers import RotatingFileHandler

LEVELS = {'debug': logging.DEBUG,
          'info': logging.INFO,
          'warning': logging.WARNING,
          'error': logging.ERROR,
          'critical': logging.CRITICAL}

LOGFILE = ConfigVariable.get("log_path")
log_handler = RotatingFileHandler(LOGFILE, maxBytes=1048576, backupCount=5)
log_handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s: %(message)s ' '[in %(pathname)s:%(lineno)d]'))
logging = logging.getLogger("Python_Valyria")

logging.addHandler(log_handler)

try:
	set_log_level = LEVELS.get(ConfigVariable.get("log_level"))
	logging.setLevel(set_log_level)
	logging.addHandler(log_handler)
except Exception as E:
	logging.setLevel(logging.debug)
	print 'Default log level set to debug\n' \
	      'Cannot set Level as {}\n' \
	      'Reason: '.format(ConfigVariable.get("AutomateValyria_log_level"), E.message)


# logger = logging.getLogger('Python_Valyria')
# hdlr = logging.FileHandler(ConfigVariable.get("AutomateValyria_log_path"))
# formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
# hdlr.setFormatter(formatter)
# logger.addHandler(hdlr)
# try:
# 	logger.setLevel(logging.ConfigVariable.get("AutomateValyria_log_level"))
# except Exception as E:
# 	logger.setLevel(logging.debug)
# 	print 'Default log level set to debug\n' \
# 	      'Cannot set Level as {}\n' \
# 	      'Reason: '.format(ConfigVariable.get("AutomateValyria_log_level"), E.message)


class Main:
	door_status = None
	door_count_from_database_start = None
	door_count_from_database_end = None

	def __int__(self):
		self.door_status = 0
		self.door_count_from_database_start = 0
		self.door_count_from_database_end = 0

	def __del__(self):
		logging.info('Test End: %s', datetime.now())
		self.write_result_to_file()
		logging.critical('Database destroyed')

	def check_if_process_is_running(self, process_name):
		"""
		:param process_name: Check the status of the process passed to this function
		:return: True of False
		"""
		try:
			process = psutil.win_service_get(process_name)
			if process.status() == 'running':
				logging.info('%s is %s', process_name, process.status())
				return True
			else:
				logging.info('%s is %s', process_name, process.status())
				return False
		except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as E:
			pass
			logging.error('%s is %s', process_name, process.status())
			logging.error('Error message: %s', E.message)
		return False

	def run_test_service_not_running(self, valyria_object):
		try:
			while True:
				valyria_object.toggle_valyria_door_switch()
				logging.debug('Toggle : %s', valyria_object.count_door_open + 1)
		except KeyboardInterrupt as E:
			logging.debug("Keyboard Interrup %s", E.message)
			logging.info("Door open count (Simulator) : %s", valyria_object.count_door_open)
			pass

	def run_test_service_running(self, valyria_object, database_object):
		self.door_count_from_database_start = database_object.get_LogicCageOpen_count()
		try:
			while True:
				valyria_object.open_valyria_door_switch()
				logging.debug('Door open')
				if database_object.get_LogicCageOpen_status() == 1:
					logging.debug('Door status database: %s', database_object.get_LogicCageOpen_status())
					valyria_object.close_valyria_door_switch()
					logging.debug('Door Closed')
				elif database_object.get_LogicCageOpen_status() == 0:
					logging.error('Door status is already closed \nIssue with Valyria driver and database sync')
					logging.error('Door status database: %s', database_object.get_LogicCageOpen_status())
					exit(0)
				else:
					logging.error('Issue is database for door status')
					exit(0)

		except KeyboardInterrupt:
			self.door_count_from_database_end = database_object.get_LogicCageOpen_count()
			logging.info('Door open count (Simulator) : %s', valyria_object.count_door_open)
			logging.info('Door open value_start (database) :  %s', self.door_count_from_database_start)
			logging.info('Door open value_end (database) : %s', self.door_count_from_database_end)
			logging.info('Door open count (database) : %s', self.door_count_from_database_end
			                                               - self.door_count_from_database_start)
			pass

	def write_result_to_file(self, simulator_count, db_door_count_start, db_door_count_end):
		with open(ConfigVariable.get("result_path"), 'w') as file_handler:
			output_string = "Door open count (Simulator) {}, simulator_count)" \
			                "Door open value_start (database) :  {}" \
			                "Door open value_end (database) : {}" \
			                "Door open count (database) : {}".format(simulator_count,
			                                                         db_door_count_start,
			                                                         db_door_count_end,
			                                                         db_door_count_end - db_door_count_start)
			print output_string
			file_handler.write("%s" % output_string)
			file_handler.close()

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
			logging.info('Got db connection to %s : %s', self.db_name, self.database_con)
		except pyodbc.Error as pyodbc_error:
			logging.error('Cannot setup database connection to %s', self.db_name)
			logging.error('Issue %s', pyodbc_error.message)

	def __del__(self):
		""" destructor"""
		logging.info('Closing database connection')
		self.database_con.close()
		logging.critical('Database destroyed')

	def execute_sql_query(self, sql_query):
		"""
		Query database and fetch data
		:param sql_query: Sql text query
		:return: data
		"""
		logging.debug('SQL query:  %s', sql_query)
		try:
			if self.database_con:
				self.database_cursor.execute(sql_query)
				fetch_db_result = self.database_cursor.fetchall()
				logging.debug('Result from query:  %s', fetch_db_result)
				return fetch_db_result
		except pyodbc.Error as pyodbc_error:
			logging.error('Problem with query\nargs: %s\nMessage: %s', pyodbc_error.args, pyodbc_error.message)
		return None

	def get_LogicCageOpen_count(self):
		sql_query = r"SELECT * FROM Vertex.dbo.ControllerMeter where MeterName='LogicCageOpen'"
		result_from_query = self.execute_sql_query(sql_query=sql_query)
		if result_from_query:
			value_from_list = [x[4] for x in result_from_query]
			logging.debug('Value get_logicCageOpenCount: %s', value_from_list)
			return value_from_list[0]
		else:
			logging.error('Logic cage open count not available in database')
		return None

	def get_LogicCageOpen_status(self):
		sql_query = r"SELECT * FROM Vertex.dbo.ControllerMeter where MeterName='LogicCageOpenState'"
		result_from_query = self.execute_sql_query(sql_query=sql_query)
		if result_from_query:
			value_from_list = [x[4] for x in result_from_query]
			logging.debug('Value get_LogicCageOpen_status: %s', value_from_list)
			return value_from_list[0]
		else:
			logging.error('Logic cage status not available in database')
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
		logging.critical('ValyriaAutomate destroyed')

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
			logging.error('timed out')
			logging.error(exc.message)
			logging.debug(exc)

	def terminate_valyria(self):
		if self.valyria_application:
			logging.debug('Kill valyria application')
			self.valyria_application.kill()
		else:
			logging.info('Valyria already running')
			logging.debug('Terminate valyria application from system command')
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
			logging.error(e.message)

	def close_valyria_door_switch(self):
		self.click_is_door_open()
		logging.debug('close_valyria_door_switch')

	def open_valyria_door_switch(self):
		self.click_is_door_open()
		self.count_door_open += 1
		logging.debug('open_valyria_door_switch: %s', self.count_door_open)

	def toggle_valyria_door_switch(self):
		logging.debug('toggle_valyria_door_switch')
		self.open_valyria_door_switch()
		self.close_valyria_door_switch()


if __name__ == '__main__':
	logging.info('Test start: %s', datetime.now())
	valyria = ValyriaAutomate()
	# valyria.print_identifiers()
	test = Main()
	aristocrat_vertex_process = ConfigVariable.get("controller_process")
	dummy_process = r'lfsvc'
	if test.check_if_process_is_running(dummy_process):
		vertex_database = Database('Vertex')
		if vertex_database:
			logging.info('Connected to database')
			test.run_test_service_running(valyria_object=valyria, database_object=vertex_database)
		else:
			logging.info('Could not connect to database')
	else:
		logging.info('%s is stopped. Running test without ControllerMeter values', aristocrat_vertex_process)
		test.run_test_service_not_running(valyria_object=valyria)
