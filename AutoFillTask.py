#coding=utf-8
#该脚本支持py2.7
import urllib2, urllib, re, codecs, os, os.path, ConfigFile, sys, re
from datetime import *
from ntlm import HTTPNtlmAuthHandler
from sgmllib import SGMLParser
from xml.dom import minidom

username = ''
password = ''
employeeName = ''

PATH_LOG = 'log\\'
PATH_CONFIG = 'config\\'
CONFIG_FILE_NAME = 'config.ini'
CONFIG_FIELD_URL = 'url'
CONFIG_FIELD_USER = 'user'
LOG_FILE_NAME = 'log.txt'
XML_FILE_NAME = 'CRM.xml'
PAGE_NODE_NAME = 'Page'

RECORD_COUNT_PATTERN = '>([\d]+?)</([\s\S]+?)>条记录'
WORKSHEET_LIST_PATTERN = '<td width="60">(.*?)</td>[\S\s]*?onclick="winopen\(\'([\S\s]*?)\'\)'

WORKSHEET_TYPE_MEETING = '会议'
WORKSHEET_TYPE_BUSINESS = '事物'
PAGE_ONGONGINGTASK = '工单列表'

def IsNum(str):	
	try:
		int(str)
		return True
	except ValueError as e:
		return False

def GetAbsPath():
	'''
	sys.argv为执行该python脚本时的命令行参数
	'''
	if len(os.path.dirname(sys.argv[0])) < 1:
		return ''
	else:
		return os.path.dirname(sys.argv[0]) + '\\'

def GetConfigFile():
	path = GetAbsPath() + PATH_CONFIG 
	if not os.path.exists(path):
		os.mkdir(path)
	return path + CONFIG_FILE_NAME

def GetLogFile():
	path = GetAbsPath() + PATH_LOG 
	if not os.path.exists(path):
		os.mkdir(path)
	return path + LOG_FILE_NAME

def Log(logStr):
	logFile = GetLogFile()
	fpLog = open(logFile, 'a')
	nowTime = datetime.strftime(datetime.now(), '%Y-%m-%d %H:%M:%S')
	fpLog.write('%s  %s\n' % (nowTime, logStr))
	fpLog.close()

class HtmlInputList(SGMLParser):
	def __init__(self, postFields):
		self.reset()
		self.attrsInit = True;
		for field in postFields:
			self.inputs[field.name] = ''
		
	def reset(self):		
		SGMLParser.reset(self)
		self.attrsInit = False;
		self.inputs = {}	

	def start_input(self, attrs):
		self.isInput = True
		count = 0
		value = ''
		name = ''
		for k,v in attrs:			
			if count >= 2:
				break
			if k == 'value':
				value = v.strip()
				count+= 1
			if k == 'name':				
				name = v.strip()
				if self.attrsInit and (not self.inputs.has_key(name)):
					continue
				count+= 1
		if count == 2:
			self.inputs[name] = value

class Field():
	def __init__(self, sourceField):
		if sourceField != None:
			self.name = sourceField.name
			self.formatStr = sourceField.formatStr
			self.defaultValue = sourceField.defaultValue
			self.fixedValue = sourceField.fixedValue

		def GetValue(self, name):
			if FieldDataType.dtInteger == self.dataType:
				return int(self.value)
			elif FieldDataType.dtBoolean == self.dataType:
				return bool(self.value)
			else:
				return self.value

class PostPage():
	def __init__(self, pageName, url, crm):
		self.pageName = pageName.decode('utf-8')
		if url.find('home:8000') == -1:
			self.url = url_base + url
		else:
			self.url = url
		self.crm = crm
		self.postData = {}
		self.postFields = []
		self.GetConfig()
		self.htmlInputs = HtmlInputList(self.postFields)		

	def GetConfig(self):		
		dom = minidom.parse(GetAbsPath() + PATH_CONFIG + XML_FILE_NAME)
		root = dom.documentElement
		pageNodeList = root.getElementsByTagName(PAGE_NODE_NAME)		
		for pageNode in pageNodeList:			
			if self.pageName == pageNode.getAttribute('Name'):				
				print ('Page = ' + self.pageName)
				postFieldsNode = pageNode.getElementsByTagName('PostFields')[0]
				fieldList = postFieldsNode.getElementsByTagName('Field')
				self.postFields = []
				for fieldNode in fieldList:
					field = Field(None)					
					field.name = fieldNode.getAttribute('Name')					
					field.formatStr = fieldNode.getAttribute('FormatStr')
					field.defaultValue = fieldNode.getAttribute('DefaultValue')
					field.fixedValue = fieldNode.getAttribute('FixedValue')
					#print ('%s=%s'%(field.name, field.fixedValue))
					self.postFields.append(field)

	def Commit(self):
		#fpw = open('postData.txt', 'w')
		self.GetPostData()	
		#for key in self.postData:
			#fpw.write(key + '=')
			#fpw.write('%s\n'%self.postData[key])
		#fpw.close()
		response = crm.Post(self.url, self.postData)

	def GetPostData(self):
		content = crm.GetPageContent(self.url)
		self.htmlInputs.feed(content)
		for field in self.postFields:
			if len(field.fixedValue) > 0:
				self.postData[field.name] = field.fixedValue
				continue
			if not self.htmlInputs.inputs.has_key(field.name):
				self.postData[field.name] = ''					
			elif len(self.htmlInputs.inputs[field.name]) < 1:				
				self.postData[field.name] = field.defaultValue
				#print ('no value %s'%field.name)
			else:
				self.postData[field.name] = self.htmlInputs.inputs[field.name]			
		return self.postData

	def GetPostDataFromResponse(self, response):
		self.htmlInputs.feed(response)
		for field in self.postFields:
			self.postData[field.name] = self.htmlInputs.inputs[field.name]
		return postData

class WorkSheet(PostPage):
	def __init__(self, wsType, url, crm):
		PostPage.__init__(self, wsType, url, crm)
		Log('Task Type: %s' % wsType)
		self.type = wsType
		self.pageName = wsType

class CRM():
	def __init__(self, username, password, employeeName):
		self.username = username
		self.password = password
		self.employeeName = employeeName
		print 'Login:%s %s'%(self.username, self.employeeName)		

	def GetPageContent(self, pageurl):
	    self.Login(pageurl)
	    response = urllib2.urlopen(pageurl)
	    return response.read()

	def Post(self, url , data):
		postStr = ''
		Log('PostTo: %s' % url)
		for key in data:
			postStr = postStr + key + '=' + data[key] + '\n'
		Log('postData:\n%s===============================================================' % postStr)
		req = urllib2.Request(url)
		data = urllib.urlencode(data)
		opener = self.Login(url)
		response = opener.open(req, data)
		return response.read()

	def GetFirstOngoingTask(self):
		url = url_ongoingTasks
		ongoingPage = PostPage(PAGE_ONGONGINGTASK, url, self)
		postData = ongoingPage.GetPostData()
		response = self.Post(url, postData)
		match = re.search(WORKSHEET_LIST_PATTERN, response)		
		if (match != None):
			task = WorkSheet(match.group(1), match.group(2), self)
			return task		
		return None

	def Login(self, url):
		passman = urllib2.HTTPPasswordMgrWithDefaultRealm()
		passman.add_password(None, url, self.username, self.password)
		# create the NTLM authentication handler
		authNTLM = HTTPNtlmAuthHandler.HTTPNtlmAuthHandler(passman)
		# create and install the opener
		opener = urllib2.build_opener(authNTLM)
		urllib2.install_opener(opener)
		return opener

	def IsWorkSheetFilled(self):
		url = url_dailyWorks
		content = self.GetPageContent(url)
		match = re.search(RECORD_COUNT_PATTERN, content)
		if (match != None) and IsNum(match.group(1)):
			count = int(match.group(1))
			if count > 0:
				return True
		if (match == None):
			Log('Search RECORD_COUNT_PATTERN Fail')	
		return False

	def FillSheetWork(self):
		if not self.IsWorkSheetFilled():
			Log('Been Filled Task')
			return
		task = self.GetFirstOngoingTask()
		if task != None:
			Log('Commit Task')	
			task.Commit()

reload(sys)
#print sys.getdefaultencoding()
sys.setdefaultencoding('utf8')
#print sys.getdefaultencoding()
Log('Application Start')
username = ConfigFile.Read(GetConfigFile(), CONFIG_FIELD_USER, 'username')	
password = ConfigFile.Read(GetConfigFile(), CONFIG_FIELD_USER, 'password')
employeeName = ConfigFile.Read(GetConfigFile(), CONFIG_FIELD_USER, 'employeeName')
url_base = ConfigFile.Read(GetConfigFile(), CONFIG_FIELD_URL, 'url_base')
url_ongoingTasks = ConfigFile.Read(GetConfigFile(), CONFIG_FIELD_URL, 'url_ongoingTasks')
url_dailyWorks = ConfigFile.Read(GetConfigFile(), CONFIG_FIELD_URL, 'url_dailyWorks')

crm = CRM(username, password, employeeName)
crm.FillSheetWork()
Log('Application Closed')
