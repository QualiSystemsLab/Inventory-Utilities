class AutoloadRow:
    def __init__(self, row):
        # set ignore
        if str(row[0].value).upper() == 'Y':
            self.ignore = True
        else:
            self.ignore = False

        # check if it's update
        if str(row[1].value).upper() == 'Y':
            self.update = True
        else:
            self.update = False

        # set autoload
        if str(row[2].value).upper() == 'Y':
            self.autoload = True
        else:
            self.autoload = False

        self.parent = row[3].value.split(' ')[0]
        self.name = row[4].value
        if self.parent:
            self.fullname = '%s/%s' % (self.parent, self.name)
        else:
            self.fullname = self.name

        self.resource_family = row[5].value
        self.resource_model = row[6].value
        self.domain = []
        dom_list = row[7].value
        temp = dom_list.split(',')
        for each in temp:
            self.domain.append(each.strip())
        self.address = row[8].value
        self.folder_path = row[9].value
        self.connection_type = row[10].value
        self.user = row[11].value
        self.password = row[12].value
        self.enable_password = row[13].value
        self.description = row[14].value
        self.driver_name = row[15].value
        self.snmp_version = row[16].value
        self.snmp_read_str = row[17].value
        self.location = row[18].value
        if self.name.strip() != '' and \
                        self.address.strip() != '' and \
                        self.resource_family.strip() != '' and \
                        self.resource_model.strip() != '':
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
            # blank check handled in the set_attributes method
            self.attributes[h] = row[n].value
            n += 1


class SetConnectionsRow:
    def __init__(self, row):
        if str(row[0].value).upper() == 'Y':
            self.ignore = True
        else:
            self.ignore = False
            
        if row[1].value == '':
            self.point_a = None
        else:
            self.point_a = row[1].value

        if row[2].value == '':
            self.point_b = None
        else:
            self.point_b = row[2].value


class CustomAttributeRow:
    def __init__(self, row):
        if str(row[0].value).upper() == 'Y':
            self.ignore = True
        else:
            self.ignore = False

        self.model_name = row[1].value
        self.attribute_name = row[2].value
        self.default_value = row[3].value


class SelectionHelper:
    def __init__(self):
        self.create_and_load = False
        self.set_attributes = False
        self.set_connections = False
        self.list_connections = False
        self.add_custom_attributes = False
        self.inventory_report = False
        self.user_report = False
        self.update_users = False


class UserUpdateRow:
    def __init__(self, row):
        if str(row[0].value).upper() == 'Y':
            self.ignore = True
        else:
            self.ignore = False

        self.user = row[1].value
        self.email = row[2].value

        if str(row[3].value).upper() == 'N':
            self.active = False
        else:
            self.active = True

        if len(row[4].value) > 0:
            self.add_groups = str(row[4].value).split(',')
        else:
            self.add_groups = []

        if len(row[5].value) > 0:
            self.remove_groups = str(row[5].value).split(',')
        else:
            self.remove_groups = []

        self.max_reservation = str(row[6].value).strip()
        self.max_duration = str(int(row[7].value) * 60)
