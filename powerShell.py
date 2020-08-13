import os
import shutil

fileSource = "D:/VAN/FINNET/EXT_BRA/"
fileDestiny = 'D:/VAN/FINNET/SAP/BB/EXTRATO/'

#pega apenas arquivos .ret
extension = ".ret", ".RET"
try:
    for arquivo in os.listdir(fileSource):
        if arquivo.endswith(extension):
            print(arquivo)
            shutil.copy2(fileSource + arquivo, fileDestiny + arquivo)
            os.remove(fileSource + arquivo)
except:
    
