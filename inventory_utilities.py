import cloudshell.api.cloudshell_api as cs_api
from cloudshell.api.common_cloudshell_api import CloudShellAPIError
import logging
import xlrd
import csv
import os
import row_helpers


BAD_VALUE = ' 80y  | 53hak~ljbdpiSY* piwh4t[09AP s<cj'
LOG_DICT = {"DEBUG": 10, "INFO": 20, "WARNING": 30, "WARN": 30, "ERROR": 40, "CRITICAL": 50, "CRIT": 50}


class CloudShellInventoryUtilities:
    def __init__(self):
        self.filepath = ''
        self.workbook = None  # excel workbook to open
        self.workbookname = 'inventory.xlsx'
        self. workbook = self._load_workbook()
        self._load_configs()
        self.cs_session = self.open_cs_session()
        self.connection_list = []
        self.selection = row_helpers.SelectionHelper()

        # set logging
        logging.basicConfig(format='%(asctime)s:%(levelname)s:%(message)s',
                            filename=self.logfilename,
                            level=LOG_DICT[self.loglevel])

    def open_cs_session(self):
        # connect to CloudShell
        try:
            return cs_api.CloudShellAPISession(self.cs_host,
                                               username=self.cs_username,
                                               password=self.cs_password,
                                               domain=self.cs_domain,
                                               port=int(self.cs_port))
        except CloudShellAPIError as err:
            print err.message
            logging.critical('Unable to open CloudShell API sesssion: %s' % err.message)
            return None

    def _load_workbook(self):
        cwd = os.getcwd()
        self.filepath = '%s/%s' % (cwd, self.workbookname)
        try:
            return xlrd.open_workbook(filename=self.filepath)
        except StandardError as err:
            print 'Could not open %s' % self.filepath
            print 'Message: %s' % err.message

    def _load_configs(self):
        sheet = self.workbook.sheet_by_name('Settings')
        self.cs_host = sheet.cell(1, 1).value
        self.cs_username = sheet.cell(2, 1).value
        self.cs_password = sheet.cell(3, 1).value
        self.cs_domain = sheet.cell(4, 1).value
        self.cs_port = sheet.cell(5, 1).value
        self.logfilename = sheet.cell(6, 1).value
        self.loglevel = sheet.cell(7,1).value
        if self.loglevel.strip() == '':
            self.loglevel = 'INFO'

    def _make_connection(self, point_a='', point_b='', override=True):
        try:
            self.cs_session.UpdatePhysicalConnection(resourceAFullPath=point_a, resourceBFullPath=point_b,
                                                     overrideExistingConnections=override)
            logging.info('Mapped Physical Connection %s to %s' %(point_a, point_b))
        except CloudShellAPIError as err:
            print 'Error - Attempting to connect %s to %s' % (point_a, point_b)
            print '  > %s' % err.message
            logging.debug('Error mapping connection %s to %s' % (point_a, point_b))
            logging.error(err.message)

    def get_attribute_value(self, device_name, attribute_name):
        try:
            lkup = self.cs_session.GetAttributeValue(resourceFullPath=device_name,
                                                     attributeName=attribute_name).Value
            logging.debug('Look up of Attribute %s Value on %s: %s' %(attribute_name, device_name, lkup))
            return lkup
        except CloudShellAPIError as err:
            print 'Error - Getting Value of Attribute %s for Device %s' % (attribute_name, device_name)
            print '  > %s' % err.message
            logging.debug('Error looking up attribute %s value on %s' %(attribute_name, device_name))
            logging.error(err.message)

    def set_attribute_value(self, device_name, attribute_name, value, may_not_exist=False):
        try:
            self.cs_session.SetAttributeValue(resourceFullPath=device_name,
                                              attributeName=attribute_name,
                                              attributeValue=value)
            logging.debug('Set new Attribute value for %s on %s: %s' %(attribute_name, device_name, value))
        except CloudShellAPIError as err:
            if may_not_exist:
                pass
            else:
                print 'Error - setting value on Attribute %s for Device %s' % (attribute_name, device_name)
                print '  > %s' % err.message

    def has_attribute(self, attribute_name, attribute_list):
        tar = BAD_VALUE
        for sub in attribute_list:
            if attribute_name.split('.')[-1] in sub:  # in case the header is 'y.user' on 'x.*' device
                # check for an exact match: 'Password' in [x.Password, x.Enable Password]
                if attribute_name == sub:
                    tar = sub
                elif attribute_name == sub.split('.')[-1]:
                    tar = sub
                elif attribute_name.split('.')[-1] == sub.split('.')[-1]:
                    tar = sub

        return tar

    def _resource_exists(self, device_name):
        """
        Boolean query to see if a device exists in inventory
        :param string device_name: the name of the device to query 
        :return: boolean 
        """
        try:
            self.cs_session.GetResourceDetails(resourceFullPath=device_name)
            return True
        except CloudShellAPIError as err:
            print err.message
            return False

    def _inner_connections(self, dev_details):
        temp = [dev_details.Name, '']
        for each in dev_details.Connections:
            temp[1] = each.FullPath
        self.connection_list.append(temp)
        for child in dev_details.ChildResources:
            self._inner_connections(child)

    def attribute_names(self, device_name):
        lst = []
        try:
            ratt = self.cs_session.GetResourceDetails(resourceFullPath=device_name).ResourceAttributes
            for res in ratt:
                lst.append(res.Name)
        except CloudShellAPIError as err:
            print err.message

        return lst

    def create_n_autoload(self):
        logging.debug('Create and Autoload called')
        sheet = self.workbook.sheet_by_name('1-CreateAndAutoLoad')
        for ro in range(5, sheet.nrows):
            row = row_helpers.AutoloadRow(sheet.row(ro))  # builds the data into our object from the AutoLoad Tab
            if row.valid and not row.ignore:
                try:
                    # build new item
                    self.cs_session.CreateResource(resourceFamily=row.resource_family,
                                                   resourceModel=row.resource_model,
                                                   resourceName=row.name,
                                                   resourceAddress=row.address,
                                                   folderFullPath=row.folder_path,
                                                   parentResourceFullPath=row.parent,
                                                   resourceDescription=row.description)
                    logging.info('New Resource Created: %s (F: %s, M: %s, A: %s, P: %s)' %
                                 (row.name, row.resource_family, row.resource_model, row.address, row.folder_path))

                    # set the driver:
                    if row.driver_name.strip() != '':
                        self.cs_session.UpdateResourceDriver(resourceFullPath=row.fullname,
                                                             driverName=row.driver_name)
                        logging.debug('Driver %s added to %s' % (row.driver_name, row.fullname))

                    # set attributes
                    a_list = self.cs_session.GetResourceDetails(resourceFullPath=row.fullname).ResourceAttributes
                    for attribute in a_list:
                        if '.' in attribute.Name:
                            a_name = attribute.Name.split('.')[1]
                        else:
                            a_name = attribute.Name

                        if a_name == 'User':
                            self.set_attribute_value(device_name=row.fullname,
                                                     attribute_name=attribute.Name,
                                                     value=row.user)
                        elif a_name == 'Password':
                            self.set_attribute_value(device_name=row.fullname,
                                                     attribute_name=attribute.Name,
                                                     value=row.password)
                        elif a_name == 'Enable Password':
                            self.set_attribute_value(device_name=row.fullname,
                                                     attribute_name=attribute.Name,
                                                     value=row.enable_password)
                        elif a_name == 'SNMP Version':
                            self.set_attribute_value(device_name=row.fullname,
                                                     attribute_name=attribute.Name,
                                                     value=row.snmp_version)
                        elif a_name == 'SNMP Read Community':
                            self.set_attribute_value(device_name=row.fullname,
                                                     attribute_name=attribute.Name,
                                                     value=row.snmp_read_str)
                        elif a_name == 'CLI Connection Type':
                            self.set_attribute_value(device_name=row.fullname,
                                                     attribute_name=attribute.Name,
                                                     value=row.connection_type)
                        elif a_name == 'Location':
                            self.set_attribute_value(device_name=row.fullname,
                                                     attribute_name=attribute.Name,
                                                     value=row.location)

                    # preform autoload
                    if row.autoload:
                        self.cs_session.AutoLoad(resourceFullPath=row.fullname)
                        logging.info('Autoload ran on %s' %row.fullname)

                except CloudShellAPIError as err:
                    print 'Error - Loading Initial Attributes from the Create and Autoload'
                    print '  > %s' % err.message
                    logging.debug('Error in Create and Autoload - %s %s %s %s'
                                  % (row.name, row.address, row.resource_family, row.resource_model))
                    logging.error(err.message)
            elif not row.valid:
                print 'Invalid Row, missing info (row # %s)' % ro
                print 'Name: %s  Address: %s  Resource Family: %s  Model %s' \
                      % (row.name, row.address, row.resource_family, row.resource_model)
                logging.warning('Invalid Row Missing info: %s' % row.name)

    def set_attributes(self):
        logging.info('Set Attributes Called')
        sheet = self.workbook.sheet_by_name('2-SetAttributes')
        custom_attributes = []
        for col in range(2, sheet.ncols):
            custom_attributes.append(sheet.cell(4, col).value)  # builds the custom attribute list
        for ro in range(5, sheet.nrows):
            row = row_helpers.SetAttributesRow(sheet.row(ro), custom_attributes)  # builds the data object
            if not row.ignore:
                device_att_list = self.attribute_names(row.name)  # get a list of this devices attributes by name
                for att in custom_attributes:  # walk the headers and assign if they match (skip blanks)
                    if sheet.cell(ro, custom_attributes.index(att)).strip() != '':  # not empty
                        x_name = self.has_attribute(attribute_name=att, attribute_list=device_att_list)
                        if x_name != BAD_VALUE:
                            try:
                                self.set_attribute_value(device_name=row.name, attribute_name=x_name,
                                                         value=row.attributes[att], may_not_exist=True)
                                logging.info('Attribute %s set to %s on %s' % (x_name, row.attributes[att], row.name))
                            except CloudShellAPIError as err:
                                print 'Error - Trying to set Attribute %s on %s' % (att, row.name)
                                print '  > %s' % err.message
                                logging.debug('Error setting Attribute %s to value %s on %s'
                                              % (x_name, row.attributes[att], row.name))
                                logging.error(err.message)

    def list_connections(self):
        logging.info('List Connections Called')
        sheet = self.workbook.sheet_by_name('4-ListConnections')
        device_query_list = []
        for ro in range(5, sheet.nrows):
            a_device = sheet.cell(ro, 0).value
            if self._resource_exists(device_name=a_device):
                device_query_list.append(a_device)
            else:
                pass

        line = ['Source', 'Connected To']
        csv_filepath = '%s/current_connections.csv' % os.getcwd()
        try:
            with open(csv_filepath, 'wb') as f:  # open in overwrite binary
                csvout = csv.writer(f)
                csvout.writerow(line)
                f.close()

            for item in device_query_list:
                self.connection_list = []
                details = self.cs_session.GetResourceDetails(resourceFullPath=item)
                self._inner_connections(details)

                with open(csv_filepath, 'ab') as f:  # open in append binary
                    csvout = csv.writer(f)
                    for pairing in self.connection_list:
                        csvout.writerow(pairing)  # each connection is [point_a, point_b]
                    csvout.writerow(['', ''])
                    f.close()

            print '==> Connections List Created: %s' % csv_filepath
        except StandardError as err:
            print 'Error Creating Connections List'
            print ' > %s' % err.message
            logging.debug('Error Creating Connection List')
            logging.error(err.message)

    def set_connections(self):
        logging.info('Set Connections Called')
        sheet = self.workbook.sheet_by_name('3-SetConnections')
        for ro in range (5, sheet.nrows):
            row = row_helpers.SetConnectionsRow(sheet.row(ro))
            if not row.ignore:
                if self._resource_exists(row.point_a) and self._resource_exists(row.point_b):
                    self._make_connection(row.point_a, row.point_b)
                elif not row.point_a:
                    pass
                elif not row.point_b:
                    pass
                else:
                    pass

    def add_custom_attributes(self):
        logging.info('Add Custom Attributes Called')
        sheet = self.workbook.sheet_by_name('0-AddCustomAttributes')

        for ro in range(5, sheet.nrows):
            row = row_helpers.CustomAttributeRow(sheet.row(ro))
            if not row.ignore:
                try:
                    self.cs_session.SetCustomShellAttribute(modelName=row.model_name,
                                                            attributeName=row.attribute_name,
                                                            defaultValue=row.default_value, restrictedValues=[''])
                    logging.info('Custom Attribute %s added to %s Shell - Default Value %s' %
                                 (row.attribute_name, row.model_name, row.default_value))
                except CloudShellAPIError as err:
                    print err.message
                    logging.error(err.message)

##########################################
def main():
    skip = False
    input_loop = True
    print '\n\nCloudShell Inventory Bulk Upload Utility'
    local = CloudShellInventoryUtilities()

    print '\nUsing: %s' % local.filepath
    print '%s' % '-' * 40
    print 'Make your selection:'
    print ' 1) Create and AutoLoad'
    print ' 2) Set Attributes'
    print ' 3) Set Connections'
    print ' 4) List Connections'
    print ' 5) Bulk Load (1, 2, 3)'

    while input_loop:
        print "\n'0' or 'exit' to Exit"
        # main prompt
        user_input = raw_input('Selection: ')

        logging.debug('User Input from Main Prompt: %s' % user_input)
        if user_input == '0' or user_input.upper() == 'EXIT':
            skip = True
            input_loop = False
        else:
            if user_input == '1':
                local.selection.create_and_load = True
                input_loop = False
            elif user_input == '2':
                local.selection.set_attributes = True
                input_loop = False
            elif user_input == '3':
                local.selection.set_connections = True
                input_loop = False
            elif user_input == '4':
                local.selection.list_connections = True
                input_loop = False
            elif user_input == '5':
                local.selection.create_and_load = True
                local.selection.set_attributes = True
                local.selection.set_connections = True
                input_loop = False
        # end while loop for input

    # act on the input
    if not skip:
        if local.selection.list_connections:
            local.list_connections()
        if local.selection.create_and_load:
            local.create_n_autoload()
        if local.selection.set_attributes:
            local.set_attributes()
        if local.selection.set_connections:
            local.set_connections()

    print '\nComplete!'


if __name__ == '__main__':
    main()