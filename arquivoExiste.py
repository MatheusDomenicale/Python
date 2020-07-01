import os.path
anexo = 'D:/Python/teste/ConectaE.xlsx'

def removeArquivo():
    if os.path.exists(anexo) ==  True:
        os.remove(anexo)
