import cloudshell.api.cloudshell_api as cs_api
from cloudshell.api.common_cloudshell_api import CloudShellAPIError
import xlrd
import csv
import os
import row_helpers
import time


BAD_VALUE = ' 80y  | 53hak~ljbdpiSY* piwh4t[09AP s<cj'


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

    def _make_connection(self, point_a='', point_b='', override=True):
        try:
            self.cs_session.UpdatePhysicalConnection(resourceAFullPath=point_a, resourceBFullPath=point_b,
                                                     overrideExistingConnections=override)
        except CloudShellAPIError as err:
            print 'Error - Attempting to connect %s to %s' % (point_a, point_b)
            print '  > %s' % err.message

    def get_attribute_value(self, device_name, attribute_name):
        try:
            return self.cs_session.GetAttributeValue(resourceFullPath=device_name,
                                                     attributeName=attribute_name).Value
        except CloudShellAPIError as err:
            print 'Error - Getting Value of Attribute %s for Device %s' % (attribute_name, device_name)
            print '  > %s' % err.message

    def set_attribute_value(self, device_name, attribute_name, value, may_not_exist=False):
        try:
            self.cs_session.SetAttributeValue(resourceFullPath=device_name,
                                              attributeName=attribute_name,
                                              attributeValue=value)
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

                    # set the driver:
                    self.cs_session.UpdateResourceDriver(resourceFullPath=row.fullname,
                                                         driverName=row.driver_name)

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

                except CloudShellAPIError as err:
                    print 'Error - Loading Initial Attributes from the Create and Autoload'
                    print '  > %s' % err.message
            elif not row.valid:
                print 'Invalid Row, missing info (row # %s)' % ro
                print 'Name: %s  Address: %s  Resource Family: %s  Model %s' % (row.name, row.address,
                                                                                row.resource_family, row.resource_model)

    def set_attributes(self):
        sheet = self.workbook.sheet_by_name('2-SetAttributes')
        custom_attributes = []
        for col in range(2, sheet.ncols):
            custom_attributes.append(sheet.cell(4, col).value)  # builds the custom attribute list
        for ro in range(5, sheet.nrows):
            row = row_helpers.SetAttributesRow(sheet.row(ro), custom_attributes)  # builds the data object
            if not row.ignore:
                device_att_list = self.attribute_names(row.name)  # get a list of this devices attributes by name
                for att in custom_attributes:  # walk the headers and assign if they match (skip blanks)
                    if sheet.cell(ro, custom_attributes.index(att)) not in ['', ' ']:  # not empty
                        x_name = self.has_attribute(attribute_name=att, attribute_list=device_att_list)
                        if x_name != BAD_VALUE:
                            try:
                                self.set_attribute_value(device_name=row.name, attribute_name=x_name,
                                                         value=row.attributes[att], may_not_exist=True)
                            except CloudShellAPIError as err:
                                print 'Error - Trying to set Attribute %s on %s' % (att, row.name)
                                print '  > %s' % err.message

    def list_connections(self):
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

    def set_connections(self):
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

    # def add_custom_attributes(self):
    #     sheet = self.workbook.sheet_by_name('0-AddCustomAttributes')
    #
    #     for ro in range(5, sheet.nrows):
    #         row = row_helpers.CustomAttributeRow(sheet.row(ro))
    #         if not row.ignore:
    #             try:
    #                 self.cs_session.SetCustomShellAttribute(modelName=row.model_name,
    #                                                         attributeName=row.attribute_name,
    #                                                         defaultValue=row.default_value, restrictedValues='')
    #             except CloudShellAPIError as err:
    #                 print err.message

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
        user_input = raw_input('Selection: ')

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