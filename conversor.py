f = open("pedidos.txt", "r")
pedidos = f.readlines()
gravar = open('pedidos_saida.txt', 'w')

# for line in pedidos:
#    if line[0].isdigit() or line[0] == " " and line.find('Folha:') == -1:
#        if line[0].isdigit():
#            print("")
#            print(' '.join(line.rstrip().split("-")[1].split()[:-6])) # Cliente
#            print(line.rstrip().split("-")[0]) # Data, etc
#            print(line.rstrip().split("-")[1].split()[-6:]) # Dados do pedido
#        else:
#            print(line.rstrip().split()) # Dados do pedido 2

for line in pedidos:
    if line[0].isdigit() or line[0] == " " and line.find('Folha:') == -1:
        if line[0].isdigit():
            data = line.rstrip().split("-")[0].split()[0]
            pedido = line.rstrip().split("-")[0].split()[2]
            cliente = line.rstrip().split("-")[0].split()[3]+'-'+' '.join(line.rstrip().split("-")[1].split()[:-6])
            gravar.write('{};{};{};{} \n'.format(data, pedido, cliente, ';'.join(line.rstrip().split("-")[1].split()[-6:])))
        else:
            gravar.write('{};{};{};{} \n'.format(data, pedido, cliente, ';'.join(line.rstrip().split())))


f.close()
gravar.close()
