import os

def main():
    
    funcionários = ['adilton dantas dos santos','allan clarck dos aantos','anderson alves souza','anderson silva santos','bruno souza santos','carlos marcondes de souza nunes','charlan messias dos santos','claudio lopes dos santos','cristiano barroso de sa','cristiano batista de araujo','decio silveira santos','edson pereira damasceno','fernande andrade dos santos pinto','francisco hilario vieira santos','francisco jackson de lima dant','gabriel santos fonseca','irlan dos santos','jefferson dacio de o lins','jose adriano correia santos','jose carlos de jesus','jose carlos dos santos filho','jose mateus otaviano dos santos','jose valteir lima de almeida','kennedy ferreira santos','leandro dos santos doria','leudson santos de souza','lincoln chaves franca','lucas michell santana santos','lugrecio lima dos santos','marcio cley dos santos','marcondes gomes da silva','matheus santos de aragão','matias de jesus santos','pedro henrique da silva lira','pedro mendes de souza','reinaldo santos almeida','thiago gedeon marques de souza','wanderley santos ferreira']
    matriculas = ['180019','180002','180084','180085','180005','180107','180113','180114','180006','180047','180007','180009','180086','180108','180012','180112','180095','180049','180115','180087','180048','180116','180106','180098','082656','180099','180096','180033','180117','180104','180020','180118','180100','180030','180039','180046','180093','180120']
    função = ['eletricista','eletricista','eletricista','motorista|munkeiro','eletricista lv','eletricista','motorista|eletricista lv','eletricista','motorista|munkeiro','encarregado','motorista|munkeiro','aux.eletricista','aux.eletricista','encarregado','eletricista','aux.técnico','encarregado','eletricista lv','motorista|munkeiro','eletricista','eletricista lv','aux.eletricista','encarregado','-','eletricista','encarregado','motorista|munkeiro','programador','encarregado','motorista|munkeiro','encarregado','aux.eletricista','aux.eletricista','aux.técnico','planejador','motorista|eletricista lv','eletricista','eletricista']
    equipe = ['lm 03','lm 03','lm 02','lm 02','lm 01','-','-','-','lm 03','lv 01','lm 01','lm 01','lm 03','-','lm 03','-','lm 01','lv 01','-','lm 02','lv 01','-','lm 04','quase fora','lm 01','lm 02','chefe transporte','-','lv','-','lm 01','-','lm 02','-','-','lv 01','lm 02','-']

    busca = str(input("Quer buscar por 'nome', 'matrícula', 'função' ou 'equipe'?\n")).lower()
    
    if busca == "nome":
        print("\nModo de Busca selecionado 'Nome'")
        a = "1"

        while a == "1":

            entrada = str(input("Digite o nome do funcionário: ")).lower()
            b=0

            for i in funcionários:
                if entrada in i:
                    saida = b 
                    print("\nNome:{}".format(funcionários[saida]).upper())
                    print("Matricula:{}".format(matriculas[saida]).upper())
                    print("Função:{}".format(função[saida]).upper())
                    print("Equipe:{}\n".format(equipe[saida]).upper())

                    caminho = "C:\\Users\\gabriel.fonseca\\Documents\\Colaboradores\\{}".format(funcionários[saida])
                    caminho=os.path.realpath(caminho)
                    os.startfile(caminho)

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
                print("Função:{}".format(função[saida]).upper())
                print("Equipe:{}\n".format(equipe[saida]).upper())
                a = "0"

                caminho = "C:\\Users\\gabriel.fonseca\\Documents\\Colaboradores\\{}".format(funcionários[saida])
                caminho = os.path.realpath(caminho)
                os.startfile(caminho)

            else:
                print("\nMatrícula incorreta... Digite novamente...")

    elif busca == "função" or busca == "funcao":
        print("\nModo de Busca selecionado 'Função'")
        a = "1"

        while a == "1":

            entrada = str(input("Digite a Função que deseja buscar: ")).lower()
            b = 0

            for i in função:
                if entrada == i:
                    saida = b
                    print("\nNome:{}".format(funcionários[saida]).upper())
                    print("Matricula:{}".format(matriculas[saida]).upper())
                    print("Função:{}".format(função[saida]).upper())
                    print("Equipe:{}\n".format(equipe[saida]).upper())
                    a = "0" 
                    b+=1 

                    caminho = "C:\\Users\\gabriel.fonseca\\Documents\\Colaboradores\\{}".format(funcionários[saida])
                    caminho=os.path.realpath(caminho)
                    os.startfile(caminho)

                else:
                    b+=1

    elif busca == "equipe":
        print("\nModo de Busca selecionado 'Equipe'")
        a = "1"

        while a == "1":

            entrada = str(input("Digite a Equipe que deseja buscar: ")).lower()
            b = 0

            for i in equipe:
                if entrada == i:
                    saida = b
                    print("\nNome:{}".format(funcionários[saida]).upper())
                    print("Matricula:{}".format(matriculas[saida]).upper())
                    print("Função:{}".format(função[saida]).upper())
                    print("Equipe:{}\n".format(equipe[saida]).upper())
                    a = "0" 
                    b+=1 

                    caminho = "C:\\Users\\gabriel.fonseca\\Documents\\Colaboradores\\{}".format(funcionários[saida])".format(funcionários[saida])
                    caminho=os.path.realpath(caminho)
                    os.startfile(caminho)

                else:
                    b+=1      

    else:
        print("Dados inválidos... Digite novamente")
        return main()
    
    main()

main()
