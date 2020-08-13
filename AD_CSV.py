import pyad
import pyad.adquery
import win32api

import csv

from pyad import *



cont = 0
valid = True
cami = 'D:/dados_industriais.csv'

pyad.set_defaults(ldap_server="", username="", password="")
ou = pyad.adcontainer.ADContainer.from_dn("OU=02 - Notebooks,OU=17 - CAA,OU=Estrutura Organizacional CIBE,DC=latam,DC=corp,DC=net")


csv.register_dialect('myDialect', quoting=csv.QUOTE_ALL)
with open(cami, 'w', newline='') as file:
    writer = csv.writer(file)

    for comp in ou.get_children():
        ntbesp = comp.get_attribute('cn')
        print (ntbesp)
        writer.writerow(ntbesp)
        grupMemberOf = comp.get_attribute('MemberOf')
        valid = True
        while valid == True:
            if  cont < len(grupMemberOf):
                ouGroup = pyad.adcontainer.ADContainer.from_dn(grupMemberOf[cont])
                namGroup = ouGroup.get_attribute('cn')
                print (namGroup)
                writer.writerow(namGroup)
                cont = 1 + cont
            else:
                cont = 0
                valid = False
