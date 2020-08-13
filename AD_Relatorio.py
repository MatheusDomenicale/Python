import pyad
import pyad.adquery
import win32api
from pyad import *



cont = 0
valid = True

pyad.set_defaults(ldap_server="", username="", password="")
ou = pyad.adcontainer.ADContainer.from_dn("OU=02 - Notebooks,OU=17 - CAA,OU=Estrutura Organizacional CIBE,DC=latam,DC=corp,DC=net")
##for obj in ou.get_children():
##    print (obj)


for comp in ou.get_children():
    ntbesp = comp.get_attribute('cn')
    print (ntbesp)
    grupMemberOf = comp.get_attribute('MemberOf')
    valid = True
    while valid == True:
        if  cont < len(grupMemberOf):
            ouGroup = pyad.adcontainer.ADContainer.from_dn(grupMemberOf[cont])
            namGroup = ouGroup.get_attribute('cn')
            print (namGroup)
            cont = 1 + cont
        else:
            cont = 0
            valid = False
