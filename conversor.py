import xlsxwriter

workbook = xlsxwriter.Workbook('Pedidos_Atrasados.xlsx')
tabela = workbook.add_worksheet()
tabela.set_column('A:A', 13)
tabela.set_column('B:B', 13)
tabela.set_column('C:C', 23)
tabela.set_column('D:D', 10)
tabela.set_column('E:E', 9)
tabela.set_column('F:H', 8)
tabela.set_column('I:I', 3)
tabela.set_column('J:J', 11)

estilo = workbook.add_format({'font_color':'white', 'bg_color':'black', 'align':'center'})
centro = workbook.add_format({'align':'center', 'border':1, 'num_format': '0'})

dic = {'SP485':'90100302', 'TP500':'90100025', 'TP500':'90100025', 'TP500':'90100025', 'TP155':'90100025', 
       'TP500':'90100025', 'TP500':'90100025', 'TP500':'90100025', 'TP500':'90100025', 'TP250':'90100025', 
       'TP500':'90100025', 'TP500':'90100025', 'TP500':'90100025', 'TP500':'90100025', 'TP500':'90100025', 
       'TP500':'90100025', 'MPX15':'90100025', 'MPX15':'90100025', 'TP500':'90100025', 'TP500':'90100025', 
       'TP155':'90100025', 'TP155':'90100025', 'TP155':'90100025', 'TP155':'90100025', 'TP155':'90100025', 
       'TP155':'90100025', 'TP155':'90100025', 'TP155':'90100025', 'TP250':'90100025', 'TP250':'90100025', 
       'TP500':'90100025', 'TP155':'90100025', 'TP155':'90100025', 'SJ330':'90100132', 'SJ330':'90100132', 
       'SJ330':'90100132', 'SJ330':'90100132', 'SJ330':'90100132', 'SJ365':'90100162', 'SJ365':'90100162', 
       'SJ365':'90100162', 'SJ328':'90100162', 'SJ365':'90100162', 'SJ365':'90100162', 'SJ365':'90100162', 
       'SJ365':'90100162', 'SJ325':'90100163', 'SJ405':'90100202', 'SJ405':'90100202', 'SJ405':'90100202', 
       'SJ405':'90100202', 'SJ455':'90100242', 'SJ455':'90100242', 'SJ455':'90100242', 'SJ455':'90100242', 
       'SJ455':'90100242', 'SJ455':'90100242', 'SJ455':'90100242', 'SJ455':'90100242', 'SP485':'90100302', 
       'SJ448':'90100302', 'SJ448':'90100302', 'SP485':'90100302', 'SJ448':'90100302', 'SJ448':'90100302', 
       'SJ448':'90100302', 'SJ448':'90100302', 'SJ448':'90100302', 'SJ448':'90100302', 'SP485':'90100302', 
       'SP221':'90100442', 'SP221':'90100442', 'SP520':'90100442', 'SP520':'90100442', 'SP131':'90100442', 
       'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 
       'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 
       'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 'SP128':'90100442', 
       'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 
       'SP221':'90100442', 'SP221':'90100442', 'SP520':'90100442', 'SP520':'90100442', 'SP221':'90100442', 
       'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 
       'SP221':'90100442', 'SP221':'90100442', 'SP520':'90100442', 'SP221':'90100442', 'SP221':'90100442', 
       'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 'SP520':'90100442', 
       'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 
       'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 'SP221':'90100442', 
       'SP520':'90100442', 'SP221':'90100442', 'SP221':'90100442', 'SP520':'90100442', 'SP128':'90100442', 
       'PP125':'90100452', 'PP125':'90100452', 'PP125':'90100452', 'PP125':'90100452', 'CY365':'90101162', 
       'CY365':'90101162', 'CY365':'90101162', 'CY365':'90101162', 'CY365':'90101162', 'CY365':'90101162', 
       'CY365':'90101162', 'CY255':'90101163', 'CY255':'90101163', 'CY505':'90101242', 'CY505':'90101242', 
       'CY505':'90101242'}

f = open("pedidos.txt", "r")
pedidos = f.readlines()
lin, col = 0, 0;gravar=[]

cabecalho = ['Data','Pedido','Cliente','Item','Cor','Qtd.','Exp.','Alm.','S', 'It.']
for dado in cabecalho:
    tabela.write(lin, col, dado, estilo)
    col+=1

for line in pedidos:
    col=0
    if line[0].isdigit() or line[0] == " " and line.find(',') != -1:
        if line[0].isdigit():
            data = line.rstrip().split("-")[0].split()[0]
            pedido = line.rstrip().split("-")[0].split()[2]
            cliente = line.rstrip().split("-")[0].split()[3]+'-'+' '.join(line.rstrip().split("-")[1].split()[:-6])
            item, cor, qtd, exp, alm, s = line.rstrip().split("-")[1].split()[-6:]
            if cor[0].isdigit():
                cor = int(cor)
            try:
                item2 = dic[item]
            except:
                item2 = '0'
            try:    
                gravar = [data, pedido, cliente, item, cor, int(qtd.split(',')[0]), int(exp.split(',')[0]), int(alm.split(',')[0]), int(s), int(item2)]
            except:
                gravar = [0]
#            print gravar
        else:
            item, cor, qtd, exp, alm, s = line.rstrip().split()
            if cor[0].isdigit():
                cor = int(cor)
            try:
                item2 = dic[item]
            except:
                item2 = '0'
            try:
                gravar = [data, pedido, cliente, item, cor, int(qtd.split(',')[0]), int(exp.split(',')[0]), int(alm.split(',')[0]), int(s), int(item2)]
            except:
                gravar = [0]
        lin+=1

    for dado in gravar:
        tabela.write(lin, col, dado, centro)
        col+=1

workbook.close()
