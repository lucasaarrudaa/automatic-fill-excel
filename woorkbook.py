from read import Read

saida = Read().wb('excls\Teste_Saidas.xlsx')
resumo = Read().ws(saida, 'Resumo')

