from inventory_utilities import CloudShellInventoryUtilities


unit = CloudShellInventoryUtilities()

# check init
# print 'Host: %s' % unit.cs_host
# print 'User: %s' % unit.cs_username
# print 'Pwrd: %s' % unit.cs_password
# print 'Domain: %s' % unit.cs_domain

# unit.create_n_autoload()

unit.set_attributes()
unit._resource_exists()


# details = unit.cs_session.GetResourceDetails('mrv-116')
# print unit._resource_exists('mrv-116')

# unit.list_connections()

# unit.set_connections()

