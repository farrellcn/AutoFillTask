#coding=utf-8

import ConfigParser, sys

def Read(configFile, field, key):
	cf = ConfigParser.ConfigParser()
	try:	
		cf.read(configFile)
		result = cf.get(field, key)
	except:		
		sys.exit(1)
	return result

#写ini文件
def Write(configFile, field, key, value):
    cf = ConfigParser.ConfigParser()
    try:
        cf.read(configFile)
        cf.set(field, key, value)
        cf.write(open(configFile,'w'))
    except:
        sys.exit(1)
    return True
