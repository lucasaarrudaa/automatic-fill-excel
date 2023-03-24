from transform import Transform
from woorkbook import saida, resumo
from ccadm import Ccadm
from cclops import Cclops
from ccmkt import Ccmkt
from resumo import Resumo

#Despesas diretas
Resumo().despesas_diretas()
Ccmkt().despesas_diretas()
Cclops().despesas_diretas()
Ccadm().despesas_diretas()

sum_brl = Transform().convert_to_brl(Ccadm().despesas_diretas(),
                                     Cclops().despesas_diretas(),
                                     Ccmkt().despesas_diretas(),
                                     Resumo().despesas_diretas(),
                                     resumo)

#Despesas Admnistrativas 
Resumo().despesas_admnistrativas()
Ccmkt().despesas_admnistrativas()
Cclops().despesas_admnistrativas()
Ccadm().despesas_admnistrativas()
    
sum_brl = Transform().convert_to_brl(Ccadm().despesas_admnistrativas(),
                                     Cclops().despesas_admnistrativas(),
                                     Ccmkt().despesas_admnistrativas(),
                                     Resumo().despesas_admnistrativas(),
                                     resumo)

Transform().complete(resumo, sum_brl, Resumo().despesas_admnistrativas())


saida.save('tsts.xlsx')



