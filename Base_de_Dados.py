from genericpath import exists
import os
import shutil
from openpyxl import Workbook, load_workbook
import time

def main():

    #Login de Usuário
    usuario = str(input("Qual o seu usuário?\n")).lower()

    #Coletar informações dos colaboradores na base de dados
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
        else:
            funcionários.append(nome)

    for celula_matricula in Aba['A']:  
        linha_matricula = celula_matricula.row
        matricula = str(Aba["A{}".format(linha_matricula)].value)
        if matricula == "Matricula":
            time.sleep(0.00001)
        else:
            matriculas.append(matricula)

    for celula_função in Aba['C']:  
        linha_função= celula_função.row
        função = str(Aba["C{}".format(linha_função)].value)
        if função == "Função":
            time.sleep(0.00001)
        else:
            funções.append(função)

    for celula_equipe in Aba['D']:  
        linha_equipe = celula_equipe.row
        equipe = str(Aba["D{}".format(linha_equipe)].value)
        if equipe == "Equipe":
            time.sleep(0.00001)
        else:
          equipes.append(equipe)

    #Criar Diretório para os funcionários que constam na base de dados
    cont = 0
    for i in funcionários:
        caminho = "C:\\Users\\{}\\Documents\\Colaboradores\\{}".format(usuario, funcionários[cont])
        if not os.path.exists(caminho):
            os.makedirs(caminho) #Criar diretório
        cont +=1

    #Apagar Pasta e Arquivos dos colaboradores que não constam mais na base de dados
    for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}\\Documents\\Colaboradores".format(usuario)):
        for diretorio in diretorios:
            if diretorio in funcionários:
                caminho_Completo = os.path.join(raiz, diretorio)
                caminho=os.path.realpath(caminho_Completo)
            else:
                caminho_Completo = os.path.join(raiz, diretorio)
                caminho=os.path.realpath(caminho_Completo)
                shutil.rmtree(caminho)

    #Copiar Arquivos de todos os colaboradores para suas pastas
    i = 0
    copiar_arquivos = str(input("\nDeseja copiar todos os arquivos existentes com o nome dos colaboradores em sua respectiva pasta?\n'sim' ou 'não'?\n")).lower()
    if copiar_arquivos == "sim":
        while i!= len(funcionários):
            for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                for arquivo in arquivos:
                    if funcionários[i] in arquivo: #Verifica se consta o nome do funcionário (Nome que completo que consta na base de dados) em algum arquivo
                        caminho_Completo = os.path.join(raiz, arquivo)
                        nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                        local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[i],nome_arquivo,ext_arquivo)
                        if os.path.exists(local_salvar):
                            print("O Arquivo {}{} já existe na pasta do Colaborador {}".format(nome_arquivo,ext_arquivo,funcionários[i]))
                        else:
                            shutil.copyfile(caminho_Completo, local_salvar) #Copia os Arquivos com o nome deste funcionário para a sua pasta
            i += 1

    #Buscar informações dos colaboradores que constam na base de dados
    def buscar_informação():

        busca = str(input("\nQuer buscar por 'Nome', 'Matrícula', 'Função' ou 'Equipe'?\n")).lower()
        
        if busca == "nome":
            print("\nModo de Busca selecionado 'Nome'")
            a = "1"

            while a == "1":

                entrada = str(input("\nDigite o nome do funcionário: ")).upper()
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

                        #Abrir diretório
                        abrir_pasta = str(input("\nDeseja abrir o diretótio referênte ao colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                        if abrir_pasta == "sim":
                            caminho = "C:\\Users\\{}\\Documents\\Colaboradores\\{}".format(usuario, funcionários[saida])
                            caminho=os.path.realpath(caminho)
                            os.startfile(caminho)
                            copiar_arquivos = str(input("\nDeseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                            if copiar_arquivos == "sim":

                                #Buscar Arquivos e Diretórios na raiz
                                for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                    for arquivo in arquivos:
                                        if funcionários[saida] in arquivo: #Verifica se consta o nome do funcionário (Nome que completo que consta na base de dados) em algum arquivo
                                            caminho_Completo = os.path.join(raiz, arquivo)
                                            nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                            local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                            if os.path.exists(local_salvar):
                                                print("O Arquivo {}{} já existe na pasta do Colaborador {}".format(nome_arquivo,ext_arquivo,funcionários[saida]))
                                            else:
                                                shutil.copyfile(caminho_Completo, local_salvar) #Copia os Arquivos com o nome deste funcionário para a sua pasta
                                
                        else:
                            copiar_arquivos = str(input("\nDeseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                            if copiar_arquivos == "sim":
                                for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                    for arquivo in arquivos:
                                        if funcionários[saida] in arquivo:
                                            caminho_Completo = os.path.join(raiz, arquivo)
                                            nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                            local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                            if os.path.exists(local_salvar):
                                                print("O Arquivo {}{} já existe na pasta do Colaborador {}".format(nome_arquivo,ext_arquivo,funcionários[saida]))
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

                    #Abrir diretório
                    abrir_pasta = str(input("\nDeseja abrir o diretótio referênte ao colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                    if abrir_pasta == "sim":
                        caminho = "C:\\Users\\{}\\Documents\\Colaboradores\\{}".format(usuario, funcionários[saida])
                        caminho=os.path.realpath(caminho)
                        os.startfile(caminho)
                        copiar_arquivos = str(input("\nDeseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                        if copiar_arquivos == "sim":

                            #Buscar Arquivos e Diretórios na raiz
                            for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                for arquivo in arquivos:
                                    if funcionários[saida] in arquivo: #Verifica se consta o nome do funcionário (Nome que completo que consta na base de dados) em algum arquivo
                                        caminho_Completo = os.path.join(raiz, arquivo)
                                        nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                        local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                        if os.path.exists(local_salvar):
                                            print("O Arquivo {}{} já existe na pasta do Colaborador {}".format(nome_arquivo,ext_arquivo,funcionários[saida]))
                                        else:
                                            shutil.copyfile(caminho_Completo, local_salvar) #Copia os Arquivos com o nome deste funcionário para a sua pasta
                                
                    else:
                        copiar_arquivos = str(input("\nDeseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                        if copiar_arquivos == "sim":
                            for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                for arquivo in arquivos:
                                    if funcionários[saida] in arquivo:
                                        caminho_Completo = os.path.join(raiz, arquivo)
                                        nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                        local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                        if os.path.exists(local_salvar):
                                            print("O Arquivo {}{} já existe na pasta do Colaborador {}".format(nome_arquivo,ext_arquivo,funcionários[saida]))
                                        else:
                                            shutil.copyfile(caminho_Completo, local_salvar)
                                    
                else:
                    print("\nMatrícula incorreta... Digite novamente...")

        elif busca == "função" or busca == "funcao":
            print("\nModo de Busca selecionado 'Função'")
            a = "1"

            while a == "1":

                entrada = str(input("Digite a Função que deseja buscar: ")).upper()
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

                        #Abrir diretório
                        abrir_pasta = str(input("\nDeseja abrir o diretótio referênte ao colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                        if abrir_pasta == "sim":
                            caminho = "C:\\Users\\{}\\Documents\\Colaboradores\\{}".format(usuario, funcionários[saida])
                            caminho=os.path.realpath(caminho)
                            os.startfile(caminho)
                            copiar_arquivos = str(input("\nDeseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                            if copiar_arquivos == "sim":

                                #Buscar Arquivos e Diretórios na raiz
                                for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                    for arquivo in arquivos:
                                        if funcionários[saida] in arquivo: #Verifica se consta o nome do funcionário (Nome que completo que consta na base de dados) em algum arquivo
                                            caminho_Completo = os.path.join(raiz, arquivo)
                                            nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                            local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                            if os.path.exists(local_salvar):
                                                print("O Arquivo {}{} já existe na pasta do Colaborador {}".format(nome_arquivo,ext_arquivo,funcionários[saida]))
                                            else:
                                                shutil.copyfile(caminho_Completo, local_salvar) #Copia os Arquivos com o nome deste funcionário para a sua pasta
                                
                        else:
                            copiar_arquivos = str(input("\nDeseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                            if copiar_arquivos == "sim":
                                for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                    for arquivo in arquivos:
                                        if funcionários[saida] in arquivo:
                                            caminho_Completo = os.path.join(raiz, arquivo)
                                            nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                            local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                            if os.path.exists(local_salvar):
                                                print("O Arquivo {}{} já existe na pasta do Colaborador {}".format(nome_arquivo,ext_arquivo,funcionários[saida]))
                                            else:
                                                shutil.copyfile(caminho_Completo, local_salvar)

                    else:
                        b+=1

        elif busca == "equipe":
            print("\nModo de Busca selecionado 'Equipe'")
            a = "1"

            while a == "1":

                entrada = str(input("Digite a Equipe que deseja buscar: ")).upper()
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

                        #Abrir diretório
                        abrir_pasta = str(input("\nDeseja abrir o diretótio referênte ao colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                        if abrir_pasta == "sim":
                            caminho = "C:\\Users\\{}\\Documents\\Colaboradores\\{}".format(usuario, funcionários[saida])
                            caminho=os.path.realpath(caminho)
                            os.startfile(caminho)
                            copiar_arquivos = str(input("\nDeseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                            if copiar_arquivos == "sim":

                                #Buscar Arquivos e Diretórios na raiz
                                for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                    for arquivo in arquivos:
                                        if funcionários[saida] in arquivo: #Verifica se consta o nome do funcionário (Nome que completo que consta na base de dados) em algum arquivo
                                            caminho_Completo = os.path.join(raiz, arquivo)
                                            nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                            local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                            if os.path.exists(local_salvar):
                                                print("O Arquivo {}{} já existe na pasta do Colaborador {}".format(nome_arquivo,ext_arquivo,funcionários[saida]))
                                            else:
                                                shutil.copyfile(caminho_Completo, local_salvar) #Copia os Arquivos com o nome deste funcionário para a sua pasta
                                
                        else:
                            copiar_arquivos = str(input("\nDeseja copiar todos os arquivos existentes com o nome do colaborador {}?\n'sim' ou 'não'?\n".format(funcionários[saida]))).lower()
                            if copiar_arquivos == "sim":
                                for raiz, diretorios, arquivos in os.walk("C:\\Users\\{}".format(usuario)):
                                    for arquivo in arquivos:
                                        if funcionários[saida] in arquivo:
                                            caminho_Completo = os.path.join(raiz, arquivo)
                                            nome_arquivo, ext_arquivo = os.path.splitext(arquivo)
                                            local_salvar = "C:\\Users\\{}\\Documents\\Colaboradores\\{}\\{}{}".format(usuario, funcionários[saida],nome_arquivo,ext_arquivo)
                                            if os.path.exists(local_salvar):
                                                print("O Arquivo {}{} já existe na pasta do Colaborador {}".format(nome_arquivo,ext_arquivo,funcionários[saida]))
                                            else:
                                                shutil.copyfile(caminho_Completo, local_salvar)

                    else:
                        b+=1      

        else:
            print("Dados inválidos... Digite novamente")
            return buscar_informação()
        
        buscar_informação() #Repete a busca de informações
    
    buscar_informação()

main()
