class AutoloadRow:
    def __init__(self, row):
        # set ignore
        if str(row[0].value).upper() == 'Y':
            self.ignore = True
        else:
            self.ignore = False

        # set autoload
        if str(row[1].value).upper() == 'Y':
            self.autoload = True
        else:
            self.autoload = False

        self.parent = row[2].value.split(' ')[0]
        self.name = row[3].value
        if self.parent:
            self.fullname = '%s/%s' % (self.parent, self.name)
        else:
            self.fullname = self.name

        self.resource_family = row[4].value
        self.resource_model = row[5].value
        self.domain = row[6].value
        self.address = row[7].value
        self.folder_path = row[8].value
        self.connection_type = row[9].value
        self.user = row[10].value
        self.password = row[11].value
        self.enable_password = row[12].value
        self.description = row[13].value
        self.driver_name = row[14].value
        self.snmp_version = row[15].value
        self.snmp_read_str = row[16].value
        self.location = row[17].value
        if self.name != '' and self.address != '' and self.resource_family != '' and self.resource_model != '':
            self.valid = True
        else:
            self.valid = False


class SetAttributesRow:
    def __init__(self, row, attribute_list):
        # set ignore
        if str(row[0].value).upper() == 'Y':
            self.ignore = True
        else:
            self.ignore = False

        self.name = row[1].value
        self.attributes = dict()
        n = 2
        for h in attribute_list:
            self.attributes[h] = row[n].value
            n += 1


class SetConnectionsRow:
    def __init__(self, row):
        if str(row[0].value).upper() == 'Y':
            self.ignore = True
        else:
            self.ignore = False

        self.point_a = row[1].value
        self.point_b = row[2].value