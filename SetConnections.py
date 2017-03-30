import cloudshell.helpers.scripts.cloudshell_scripts_helpers as helpers
import cloudshell.api.common_cloudshell_api as quali_api
from cloudshell.api.cloudshell_api import CloudShellAPISession
from urllib2 import URLError
import xlrd
import sys
import logging


workbookname = 'C:\\CloudShell\\DataModel\\Inventory.xlsx'

try:
    WorkBook = xlrd.open_workbook(workbookname)
    print "Workbook open."
except:
    print 'Could not open file  %s' % workbookname
    sys.exit(-9)

sheet = WorkBook.sheet_by_name('Settings')
cs_host = sheet.cell(1,1).value
cs_username = sheet.cell(2,1).value
cs_password = sheet.cell(3,1).value
cs_domain = sheet.cell(4,1).value
logfilename = sheet.cell(5,1).value

logging.basicConfig(filename=logfilename, level=logging.DEBUG, format='%(asctime)s %(levelname)s: %(message)s')
logging.info('--------------- SetConnections Starting up! -------------------')

try:
    cs = helpers.CloudShellAPISession(host=cs_host,username=cs_username,password=cs_password,
                                      domain=cs_domain,timezone='UTC',port=8029)
    print "CloudShell session open."
except (quali_api.CloudShellAPIError, URLError) as e:
        if isinstance(e, quali_api.CloudShellAPIError):
            if 'user:' in e.message:
                print e.message
                logging.error(e.message)
            else:
                print 'Login failed for user: ' + cs_username + ' and domain: ' + cs_domain
                logging.error('Login failed for user: ' + cs_username + ' and domain: ' + cs_domain)

        elif isinstance(e, URLError):
            print "Connection error, please check server address: " + cs_host
            logging.error("Connection error, please check server address: " + cs_host)
        quit()

sheet = WorkBook.sheet_by_name('3-SetConnections')

num_cols = 4

for row in range(5, sheet.nrows):
    ignore = sheet.cell(row,0).value
    FromPath = sheet.cell(row,1).value
    ToPath = sheet.cell(row,2).value
    if ignore != 'Y':
        print '%s -to- %s.' % (FromPath,ToPath)
        logging.info('%s -to- %s.', FromPath, ToPath)
        try:
            cs.UpdatePhysicalConnection(FromPath, ToPath,overrideExistingConnections=True)
        except (quali_api.CloudShellAPIError, URLError) as e:
                if isinstance(e, quali_api.CloudShellAPIError):
                    print e.message
                    logging.error('Connecting %s -to- %s.... FAILED.  %s', FromPath, ToPath, repr(e.message))

    elif FromPath > ' ':
        print "Ignoring row %s " % row
        logging.info("Ignoring row %s ", row)

helpers.CloudShellAPISession.Logoff(cs)
print "Run complete."
logging.info('Run complete.')