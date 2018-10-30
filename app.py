import xlsxfunctions
from pgget import Connection
cnn = Connection()

kurum = cnn.getlistofdata('kurum','objectid','analiz_tamamlandi_first is true')



for i in kurum:
    a = xlsxfunctions.KurumTablosu(i[0])
    a.veri_turu()
    # a.veri_formati()
    a.save_excel()