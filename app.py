import xlsxfunctions
import time
import sys
from pgget import Connection
cnn = Connection()

kurum = cnn.getlistofdata('kurum','objectid','analiz_tamamlandi_first is true')
  


toolbar_width = 61

# setup toolbar
sys.stdout.write("[%s]" % (" " * toolbar_width))
sys.stdout.flush()
sys.stdout.write("\b" * (toolbar_width+1)) # return to start of line, after '['

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
    a.projeksiyon_datum()
    a.metaveri()
    a.save_excel() 

    sys.stdout.write("-")
    sys.stdout.flush()

sys.stdout.write("\n")