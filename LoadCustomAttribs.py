import cloudshell.helpers.scripts.cloudshell_scripts_helpers as helpers
import xlrd
import sys
import logging

# Admin use only upon loading a new Shell (that creates new models not previously installed)
# This project is to add attributes to models created by shells.  Needs to be handled as a special task if new
# shells are loaded with models not yet in the system.
# expects file  C:\CloudShell\DataModel\Inventory.xlsx

workbookname = 'C:\\CloudShell\\DataModel\\Inventory.xlsx'

try:
    WorkBook = xlrd.open_workbook(workbookname)
    print "Workbook open."
except:
    print "Could not open file  " + workbookname
    sys.exit(-9)

sheet = WorkBook.sheet_by_name('Settings')
cs_host = sheet.cell(1,1).value
cs_username = sheet.cell(2,1).value
cs_password = sheet.cell(3,1).value
cs_domain = sheet.cell(4,1).value
logfilename = sheet.cell(5,1).value

logging.basicConfig(filename=logfilename, level=logging.DEBUG, format='%(asctime)s %(levelname)s: %(message)s')
logging.info('--------------- LoadCustomAttribs Starting up! -------------------')

try:
    cs = helpers.CloudShellAPISession(host='svl-dev-quali',username='admin',password='dev',
                                      domain='Global',timezone='UTC',port=8029)
    print "CloudShell session open."
    logging.info('CS Session opened.')
except:
    print "Could not access CloudShell."
    sys.exit(-8)

try:
    sheet = WorkBook.sheet_by_name('0-AddCustomAttribs')

    num_cols = 3

    for row in range(5, sheet.nrows):
        ignore = sheet.cell(row,0).value
        ModelName = sheet.cell(row, 1).value
        if ignore != 'Y' and ModelName > ' ':
            logging.info('Processing model ' + ModelName)
            AttributeName = sheet.cell(row, 2).value
            DefaultValue = sheet.cell(row,3).value
            cs.SetCustomShellAttribute(modelName=ModelName,attributeName=AttributeName,
                                       defaultValue=DefaultValue,restrictedValues='')
            print '   ModelName %s updated with %s.' % (ModelName, AttributeName)
            logging.info('   ModelName %s updated with %s.', ModelName, AttributeName)

except:
    print ('Failed at row [%s]' % row)
    logging.info('Failed at row [%s]', row)
    quit()

print "Run complete."
logging.info('Run complete!')

helpers.CloudShellAPISession.Logoff(cs)

