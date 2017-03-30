import cloudshell.helpers.scripts.cloudshell_scripts_helpers as helpers
import cloudshell.api.common_cloudshell_api as quali_api
from cloudshell.api.cloudshell_api import CloudShellAPISession
from urllib2 import URLError
import xlrd
from time import gmtime, strftime
import sys
import logging

#


workbookname = 'C:\\CloudShell\\DataModel\\Inventory.xlsx'

try:
    WorkBook = xlrd.open_workbook(workbookname)
    print "Workbook open."
    logging.info('Workbook opened.')
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
logging.info('--------------- CreateAndAutoLoad Starting up! -------------------')

try:
    cs = helpers.CloudShellAPISession(host=cs_host,username=cs_username,password=cs_password,
                                      domain=cs_domain,timezone='UTC',port=8029)
    print "CloudShell session open."
    logging.info('CS Session opened.')
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

sheet = WorkBook.sheet_by_name('1-CreateAndAutoLoad')

num_cols = 17

# diag:::  print sheet.cell(4,16).value  # should be "SNMP Read Community"

print ''

for row in range(5, sheet.nrows):
    ignore = sheet.cell(row,0).value
    autoload = sheet.cell(row,1).value
    Parent = sheet.cell(row,2).value
    ResourceName = sheet.cell(row, 3).value
    FamilyName = sheet.cell(row, 4).value
    ModelName = sheet.cell(row,5).value
    Domain = sheet.cell(row,6).value
    FullAddress = sheet.cell(row,7).value
    if ignore != 'Y' and ResourceName > ' ' and ModelName > ' ' and FullAddress > ' ':
        print "Processing ", row, Parent, ResourceName
        logging.info('Row %s, Parent %s, ResourceName %s', row, Parent, ResourceName)
        FolderPath = sheet.cell(row,8).value
        Description = sheet.cell(row,13).value
        if FamilyName != 'CS_Port':
            ConnectionType = sheet.cell(row,9).value
            Username = sheet.cell(row, 10).value
            Password = sheet.cell(row, 11).value
            EnablePassword = sheet.cell(row, 12).value
            DriverName = sheet.cell(row, 14).value
            SNMPVersion = sheet.cell(row, 15).value
            SNMPReadCommunity = sheet.cell(row, 16).value

        cs.CreateResource(FamilyName,ModelName,ResourceName,FullAddress,FolderPath,Parent,Description)
        print '%s Created.' % (ResourceName)
        logging.info('%s created.', ResourceName)
        if Parent <= ' ':
            if FolderPath != '':
                resource_path = FolderPath + '/' + ResourceName
            else:
                resource_path = ResourceName

            if ModelName == 'Cisco IOS Switch 2G':
                ATT_ConnectionType = 'Cisco IOS Switch 2G.CLI Connection Type'
                ATT_Username = 'Cisco IOS Switch 2G.User'
                ATT_Password = 'Cisco IOS Switch 2G.Password'
                ATT_EnablePassword = 'Cisco IOS Switch 2G.Enable Password'
                ATT_Description = 'Cisco IOS Switch 2G.Description'
                ATT_SNMPVersion = 'Cisco IOS Switch 2G.SNMP Version'
                ATT_SNMPReadCommunity = 'Cisco IOS Switch 2G.SNMP Read Community'
            elif ModelName == 'Cisco IOS Router 2G':
                ATT_ConnectionType = 'Cisco IOS Router 2G.CLI Connection Type'
                ATT_Username = 'Cisco IOS Router 2G.User'
                ATT_Password = 'Cisco IOS Router 2G.Password'
                ATT_EnablePassword = 'Cisco IOS Router 2G.Enable Password'
                ATT_Description = 'Cisco IOS Router 2G.Description'
                ATT_SNMPVersion = 'Cisco IOS Router 2G.SNMP Version'
                ATT_SNMPReadCommunity = 'Cisco IOS Router 2G.SNMP Read Community'
            elif ModelName == 'Cisco NXOS Switch 2G':
                ATT_ConnectionType = 'Cisco NXOS Switch 2G.CLI Connection Type'
                ATT_Username = 'Cisco NXOS Switch 2G.User'
                ATT_Password = 'Cisco NXOS Switch 2G.Password'
                ATT_EnablePassword = 'Cisco NXOS Switch 2G.Enable Password'
                ATT_Description = 'Cisco NXOS Switch 2G.Description'
                ATT_SNMPVersion = 'Cisco NXOS Switch 2G.SNMP Version'
                ATT_SNMPReadCommunity = 'Cisco NXOS Switch 2G.SNMP Read Community'
            else:
                ATT_ConnectionType = 'CLI Connection Type'
                ATT_Username = 'User'
                ATT_Password = 'Password'
                ATT_EnablePassword = 'EnablePassword'
                ATT_Description = 'Description'
                ATT_SNMPVersion = 'SNMP Version'
                ATT_SNMPReadCommunity = 'SNMP Read Communnity'

            if Username != '':
                logging.info('   Set attribute ' + ATT_Username + ' to ' + Username)
                cs.SetAttributeValue(resource_path,ATT_Username,Username)

            if Password != '':
                logging.info('   Set attribute ' + ATT_Password + ' to ' + Password)
                cs.SetAttributeValue(resource_path,ATT_Password,Password)

            if EnablePassword != '':
                logging.info('   Set attribute ' + ATT_EnablePassword + ' to ' + EnablePassword)
                cs.SetAttributeValue(resource_path,ATT_EnablePassword,EnablePassword)

            if ConnectionType != '':
                logging.info('   Set attribute ' + ATT_ConnectionType + '  to ' + ConnectionType)
                cs.SetAttributeValue(resource_path,ATT_ConnectionType,ConnectionType)

            if SNMPVersion != '':
                    logging.info(')  Set attribute ' + SNMPVersion + ' to ' + ATT_SNMPVersion)
                    cs.SetAttributeValue(resource_path,ATT_SNMPVersion,SNMPVersion)

            if SNMPReadCommunity != '':
                logging.info('   Set attribute ' + ATT_SNMPReadCommunity + ' to ' + SNMPReadCommunity)
                cs.SetAttributeValue(resource_path,ATT_SNMPReadCommunity,SNMPReadCommunity)

            if DriverName != '':
                print '   Setting driver to ' + DriverName
                logging.info('   Setting driver to ' + DriverName)
                cs.UpdateResourceDriver(resource_path, DriverName)

        else:
            resource_path = Parent + '/' + ResourceName   # for a sub resource

        print 'Create phase complete for ' + ResourceName
        logging.info('Create phase complete for ' + ResourceName)

        # AUTOLOAD
        """ Run AutoLoad"""
        if autoload == 'Y':
            print 'AutoLoading ' + ResourceName + ' at addr ' + FullAddress
            logging.info('Autoloading ' + ResourceName)
            cs.AutoLoad(resource_path)

        # ADD Resource to domains
        resnames = list()
        resnames.append(resource_path)
        domainlist = str(Domain.strip()).split(';')
        for domainname in domainlist:
            if domainname != '':
                print 'Adding ' + ResourceName + ' to domain ' + domainname
                logging.info('Adding ' + ResourceName + ' to domain ' + domainname)
                cs.AddResourcesToDomain(domainname, resnames, includeDecendants=True)


    else:
        if ignore != 'Y':
             print "Incomplete data for %s --skipping." % (ResourceName)
             logging.info('Incomplete data for %s --skipping.', ResourceName)
        else:
            print "Ignoring " + ResourceName
            logging.info('Ignoring ' + ResourceName)

logging.info('Run complete!')

helpers.CloudShellAPISession.Logoff(cs)