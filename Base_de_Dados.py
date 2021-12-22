from genericpath import exists
import os
import shutil
from openpyxl import Workbook, load_workbook
import time

def main():

    usuario = str(input("Qual o seu usuário?\n"))

    Planilha_base_de_dados = load_workbook("C:\\Users\\{}\\Documents\\Colaboradores\\Colaboradores.xlsx".format(usuario))
    Aba = Planilha_base_de_dados.active

    funcionários = []
    matriculas = []
    funções = []
    equipes= []

    for celula_nome in Aba['B']:  
        linha_nome = celula_nome.row
        nome = str(Aba["B{}".format(linha_nome)].value)
        if nome == "Funcionário":
            time.sleep(0.00001)
        #elif nome =="":
        #    time.sleep(0.00001)
        else:
            funcionários.append(nome)

    for celula_matricula in Aba['A']:  
        linha_matricula = celula_matricula.row
        matricula = str(Aba["A{}".format(linha_matricula)].value)
        if matricula == "Matricula":
            time.sleep(0.00001)
        #elif matricula =="":
        #    time.sleep(0.00001)
        else:
            matriculas.append(matricula)

    for celula_função in Aba['C']:  
        linha_função= celula_função.row
        função = str(Aba["C{}".format(linha_função)].value)
        if função == "Função":
            time.sleep(0.00001)
        #elif função =="":
        #    time.sleep(0.00001)
        else:
            funções.append(função)

    for celula_equipe in Aba['D']:  
        linha_equipe = celula_equipe.row
        equipe = str(Aba["D{}".format(linha_equipe)].value)
        if equipe == "Equipe":
            time.sleep(0.00001)
        #elif equipe =="":
        #    time.sleep(0.00001)
        else:
          equipes.append(equipe)
    
    def buscar_informação():

        busca = str(input("\nQuer buscar por 'nome', 'matrícula', 'função' ou 'equipe'?\n")).lower()
        
        if busca == "nome":
            print("\nModo de Busca selecionado 'Nome'")
            a = "1"

            while a == "1":

                entrada = str(input("Digite o nome do funcionário: ")).lower()
                b=0
                c=0

                for i in funcionários:
                    if entrada in i:
                        c += 1

                print("\nExistem {} colaboradores com este nome.".format(c)) 

                for i in funcionários:
                    if entrada in i:
                        saida = b 
                        print("\nNome:{}".format(funcionários[saida]).upper())
                        print("Matricula:{}".format(matriculas[saida]).upper())
                        print("Função:{}".format(funções[saida]).upper())
                        print("Equipe:{}\n".format(equipes[saida]).upper())

                        caminho = "C:\\Users\\{}\\Documents\\Colaboradores\\{}".format(usuario, funcionários[saida])
                        if not os.path.exists(caminho):
                            os.makedirs(caminho) #Criar diretório
                            abrir_pasta = str(input("Deseja abrir o diretótio referênte ao colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                            if abrir_pasta == "sim":
                                caminho=os.path.realpath(caminho)
                                os.startfile(caminho)#Abrir diretório
                                copiar_arquivos = str(input("Deseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                                if copiar_arquivos == "sim":
                                    for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                        for arquivo in arquivos:
                                            if entrada in arquivo.lower():
                                                caminho_Completo = os.path.join(raiz, arquivo)
                                                nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                                local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                                shutil.copyfile(caminho_Completo, local_salvar)
                                
                            else:
                                copiar_arquivos = str(input("Deseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                                if copiar_arquivos == "sim":
                                    for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                        for arquivo in arquivos:
                                            if entrada in arquivo.lower:
                                                caminho_Completo = os.path.join(raiz, arquivo)
                                                nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                                local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                                shutil.copyfile(caminho_Completo, local_salvar)
                        
                        else:
                            abrir_pasta = str(input("Deseja abrir o diretótio referênte ao colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                            if abrir_pasta == "sim":
                                caminho=os.path.realpath(caminho)
                                os.startfile(caminho)#Abrir diretório
                                copiar_arquivos = str(input("Deseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                                if copiar_arquivos == "sim":
                                    for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                        for arquivo in arquivos:
                                            if entrada in arquivo.lower():
                                                caminho_Completo = os.path.join(raiz, arquivo)
                                                nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                                local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                                shutil.copyfile(caminho_Completo, local_salvar)
                                
                            else:
                                copiar_arquivos = str(input("Deseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                                if copiar_arquivos == "sim":
                                    for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                        for arquivo in arquivos:
                                            if entrada in arquivo.lower():
                                                caminho_Completo = os.path.join(raiz, arquivo)
                                                nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                                local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                                if arquivo in local_salvar:
                                                    time.sleep(0.0001)
                                                else:
                                                    shutil.copyfile(caminho_Completo, local_salvar)
                                
                        a = "0"
                        b+=1

                    else:
                        b+=1

        elif busca == "matrícula" or busca =="matricula":
            print("\nModo de Busca selecionado 'Matrícula'")
            a = "1"

            while a == "1":

                entrada = str(input("Digite a matrícula do funcionário: "))

                if entrada in matriculas:
                    saida = matriculas.index(entrada) 
                    print("\nNome:{}".format(funcionários[saida]).upper())
                    print("Matricula:{}".format(matriculas[saida]).upper())
                    print("Função:{}".format(funções[saida]).upper())
                    print("Equipe:{}\n".format(equipes[saida]).upper())
                    a = "0"

                    caminho = "C:\\Users\\{}\\Documents\\Colaboradores\\{}".format(usuario, funcionários[saida])
                    if not os.path.exists(caminho):
                        os.makedirs(caminho) #Criar diretório
                        abrir_pasta = str(input("Deseja abrir o diretótio referênte ao colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                        if abrir_pasta == "sim":
                            caminho=os.path.realpath(caminho)
                            os.startfile(caminho)#Abrir diretório
                            copiar_arquivos = str(input("Deseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                            if copiar_arquivos == "sim":
                                for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                    for arquivo in arquivos:
                                        if entrada in arquivo:
                                            caminho_Completo = os.path.join(raiz, arquivo)
                                            nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                            local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                            shutil.copyfile(caminho_Completo, local_salvar)
                                
                        else:
                            copiar_arquivos = str(input("Deseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                            if copiar_arquivos == "sim":
                                for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                    for arquivo in arquivos:
                                        if entrada in arquivo:
                                            caminho_Completo = os.path.join(raiz, arquivo)
                                            nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                            local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                            shutil.copyfile(caminho_Completo, local_salvar)
                        
                    else:
                            abrir_pasta = str(input("Deseja abrir o diretótio referênte ao colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                            if abrir_pasta == "sim":
                                caminho=os.path.realpath(caminho)
                                os.startfile(caminho)#Abrir diretório
                                copiar_arquivos = str(input("Deseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                                if copiar_arquivos == "sim":
                                    for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                        for arquivo in arquivos:
                                            if entrada in arquivo:
                                                caminho_Completo = os.path.join(raiz, arquivo)
                                                nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                                local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                                shutil.copyfile(caminho_Completo, local_salvar)
                                
                            else:
                                copiar_arquivos = str(input("Deseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                                if copiar_arquivos == "sim":
                                    for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                        for arquivo in arquivos:
                                            if entrada in arquivo:
                                                caminho_Completo = os.path.join(raiz, arquivo)
                                                nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                                local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                                if arquivo in local_salvar:
                                                    time.sleep(0.0001)
                                                else:
                                                    shutil.copyfile(caminho_Completo, local_salvar)
                                    
                else:
                    print("\nMatrícula incorreta... Digite novamente...")

        elif busca == "função" or busca == "funcao":
            print("\nModo de Busca selecionado 'Função'")
            a = "1"

            while a == "1":

                entrada = str(input("Digite a Função que deseja buscar: ")).lower()
                b = 0
                c = 0

                for i in funcionários:
                    if entrada in i:
                        c += 1

                print("\nExistem {} colaboradores com esta função.".format(c))

                for i in funções:
                    if entrada == i:
                        saida = b
                        print("\nNome:{}".format(funcionários[saida]).upper())
                        print("Matricula:{}".format(matriculas[saida]).upper())
                        print("Função:{}".format(funções[saida]).upper())
                        print("Equipe:{}\n".format(equipes[saida]).upper())
                        a = "0" 
                        b+=1 

                        caminho = "C:\\Users\\{}\\Documents\\Colaboradores\\{}".format(usuario, funcionários[saida])
                        if not os.path.exists(caminho):
                            os.makedirs(caminho) #Criar diretório
                            abrir_pasta = str(input("Deseja abrir o diretótio referênte ao colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                            if abrir_pasta == "sim":
                                caminho=os.path.realpath(caminho)
                                os.startfile(caminho)#Abrir diretório
                                copiar_arquivos = str(input("Deseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                                if copiar_arquivos == "sim":
                                    for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                        for arquivo in arquivos:
                                            if entrada in arquivo:
                                                caminho_Completo = os.path.join(raiz, arquivo)
                                                nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                                local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                                shutil.copyfile(caminho_Completo, local_salvar)
                                
                            else:
                                copiar_arquivos = str(input("Deseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                                if copiar_arquivos == "sim":
                                    for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                        for arquivo in arquivos:
                                            if entrada in arquivo:
                                                caminho_Completo = os.path.join(raiz, arquivo)
                                                nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                                local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                                shutil.copyfile(caminho_Completo, local_salvar)
                        
                        else:
                            abrir_pasta = str(input("Deseja abrir o diretótio referênte ao colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                            if abrir_pasta == "sim":
                                caminho=os.path.realpath(caminho)
                                os.startfile(caminho)#Abrir diretório
                                copiar_arquivos = str(input("Deseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                                if copiar_arquivos == "sim":
                                    for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                        for arquivo in arquivos:
                                            if entrada in arquivo:
                                                caminho_Completo = os.path.join(raiz, arquivo)
                                                nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                                local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                                shutil.copyfile(caminho_Completo, local_salvar)
                                
                            else:
                                copiar_arquivos = str(input("Deseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                                if copiar_arquivos == "sim":
                                    for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                        for arquivo in arquivos:
                                            if entrada in arquivo:
                                                caminho_Completo = os.path.join(raiz, arquivo)
                                                nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                                local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                                if arquivo in local_salvar:
                                                    time.sleep(0.0001)
                                                else:
                                                    shutil.copyfile(caminho_Completo, local_salvar)

                    else:
                        b+=1

        elif busca == "equipe":
            print("\nModo de Busca selecionado 'Equipe'")
            a = "1"

            while a == "1":

                entrada = str(input("Digite a Equipe que deseja buscar: ")).lower()
                b = 0
                c = 0

                for i in funcionários:
                    if entrada in i:
                        c += 1

                print("\nExistem {} colaboradores nesta equipe.".format(c))

                for i in equipes:
                    if entrada == i:
                        saida = b
                        print("\nNome:{}".format(funcionários[saida]).upper())
                        print("Matricula:{}".format(matriculas[saida]).upper())
                        print("Função:{}".format(funções[saida]).upper())
                        print("Equipe:{}\n".format(equipes[saida]).upper())
                        a = "0" 
                        b+=1 

                        caminho = "C:\\Users\\{}\\Documents\\Colaboradores\\{}".format(usuario, funcionários[saida])
                        if not os.path.exists(caminho):
                            os.makedirs(caminho) #Criar diretório
                            abrir_pasta = str(input("Deseja abrir o diretótio referênte ao colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                            if abrir_pasta == "sim":
                                caminho=os.path.realpath(caminho)
                                os.startfile(caminho)#Abrir diretório
                                copiar_arquivos = str(input("Deseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                                if copiar_arquivos == "sim":
                                    for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                        for arquivo in arquivos:
                                            if entrada in arquivo:
                                                caminho_Completo = os.path.join(raiz, arquivo)
                                                nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                                local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                                shutil.copyfile(caminho_Completo, local_salvar)
                                
                            else:
                                copiar_arquivos = str(input("Deseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                                if copiar_arquivos == "sim":
                                    for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                        for arquivo in arquivos:
                                            if entrada in arquivo:
                                                caminho_Completo = os.path.join(raiz, arquivo)
                                                nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                                local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                                shutil.copyfile(caminho_Completo, local_salvar)
                        
                        else:
                            abrir_pasta = str(input("Deseja abrir o diretótio referênte ao colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                            if abrir_pasta == "sim":
                                caminho=os.path.realpath(caminho)
                                os.startfile(caminho)#Abrir diretório
                                copiar_arquivos = str(input("Deseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                                if copiar_arquivos == "sim":
                                    for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                        for arquivo in arquivos:
                                            if entrada in arquivo:
                                                caminho_Completo = os.path.join(raiz, arquivo)
                                                nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                                local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                                shutil.copyfile(caminho_Completo, local_salvar)
                                
                            else:
                                copiar_arquivos = str(input("Deseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                                if copiar_arquivos == "sim":
                                    for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                        for arquivo in arquivos:
                                            if entrada in arquivo:
                                                caminho_Completo = os.path.join(raiz, arquivo)
                                                nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                                local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                                if arquivo in local_salvar:
                                                    time.sleep(0.0001)
                                                else:
                                                    shutil.copyfile(caminho_Completo, local_salvar)

                    else:
                        b+=1      

        else:
            print("Dados inválidos... Digite novamente")
            return buscar_informação()
        
        buscar_informação()
    
    buscar_informação()

main()
