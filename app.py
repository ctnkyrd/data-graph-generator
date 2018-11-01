import xlsxfunctions
from pgget import Connection
cnn = Connection()

kurum = cnn.getlistofdata('kurum','objectid','analiz_tamamlandi_first is true')



for i in kurum:
    a = xlsxfunctions.KurumTablosu(i[0])
    a.veri_turu()
    a.veriformati()
    # a.projeksiyondatum
    a.veri_eksiksizlik()
    a.mantiksal_tutarlilik()
    a.vk_konumsal()
    a.vk_zamansal() 
    a.vk_tematik() 
    a.vk_guncel()
    a.vk_ogc()
    a.save_excel() 