# -*- coding: utf-8 -*-
import xlsxwriter, sys, os, numpy
import pandas as pd
from pgget import Connection
cnn = Connection()

class KurumTablosu:
    def __init__(self, oid):
        self.oid = oid
        self.k_adi = cnn.getSingledataByOid('kurum', 'k_adi', self.oid)
        self.wb = xlsxwriter.Workbook('created_excels\\'+self.k_adi.decode('utf-8')+'.xlsx')
        self.ek2 = cnn.getsinglekoddata('ek_2_cografi_veri_analizi', 'objectid', 'kurum='+str(self.oid))


        # VeriTürü
        veriler = cnn.getlistofdata('x_ek_2_tucbs_veri_katmani', 'objectid, veri_turu',
                                    'geodurum = true and katman_durumu is true and ek_2='+str(self.ek2))

        self.vt_dijital = 0
        self.vt_basili = 0
        self.vt_bilinmiyor = 0

        for row in veriler:
            if row[1] == 2:
                self.vt_dijital += 1
            elif row[1] == 1:
                self.vt_basili += 1
            elif row[1] == None:
                self.vt_bilinmiyor += 1
                
        if self.vt_dijital == 0:
            self.vt_dijital = None
        if self.vt_basili == 0:
            self.vt_basili = None
        if self.vt_bilinmiyor == 0:
            self.vt_bilinmiyor = None

        # VeriFormatı

        veriformati = cnn.getlistofdata('x_ek_2_tucbs_veri_katmani', 'objectid, veri_turu, veri_formati, veri_formati',
                                    'geodurum = true and katman_durumu is true and ek_2='+str(self.ek2))
        self.cad = 0
        self.raster = 0
        self.vt = 0
        self.raster_bas = 0
        self.raster_dij = 0
        self.vf_bilinmiyor = 0

        for row in veriformati:
            if row [3] in (8,9,10,11):
                self.cad +=1
            elif row [3] in (12,13,14):
                if row[1] == 1:
                    self.raster_bas += 1
                elif row[1] == 2:
                    self.raster_dij += 1
                self.raster +=1
            elif row [3] in (2,3,4,5,6,7,16):
                self.vt +=1
            elif row[3] in (1, 15) or row[3] is None:
                self.vf_bilinmiyor += 1 
            
        if self.cad == 0:
            self.cad = None
        if self.raster == 0:
            self.raster = None

        if self.vt == 0:
            self.vt = None

        if self.raster_bas == 0:
            self.raster_bas = None
        if self.raster_dij == 0:
            self.raster_dij = None
        
        if self.vf_bilinmiyor == 0:
            self.vf_bilinmiyor = None

        # VeriEksiksizlik
     
        veri_eksiksizlik = cnn.getlistofdata('x_ek_2_tucbs_veri_katmani', 'objectid, vk_eksizlik_yeni',
                                    'geodurum = true and katman_durumu is true and ek_2='+str(self.ek2))
        self.vt_eksik = 0
        self.vt_tam = 0
        self.vte_bilinmiyor = 0

        for row in veri_eksiksizlik:
            if row[1] == 2:
                self.vt_eksik += 1
            elif row[1] == 1:
                self.vt_tam += 1
            elif row[1] == None or row[1] == 3:
                self.vte_bilinmiyor += 1
        if self.vt_eksik == 0:
            self.vt_eksik = None
        if self.vt_tam == 0:
            self.vt_tam = None
        if self.vte_bilinmiyor == 0:
            self.vte_bilinmiyor = None


        # MantıksalTutarlılık 
     
        mantiksal_tutarlilik = cnn.getlistofdata('x_ek_2_tucbs_veri_katmani', 'objectid, vk_mantiksal_yeni',
                                    'geodurum = true and katman_durumu is true and ek_2='+str(self.ek2))
        self.vtm_var = 0
        self.vtm_yok = 0
        self.vtm_bilinmiyor = 0

        for row in mantiksal_tutarlilik:
            if row[1] == 1:
                self.vtm_var += 1
            elif row[1] == 2:
                self.vtm_yok += 1
            elif row[1] == None or row[1] == 3:
                self.vtm_bilinmiyor += 1
        if self.vtm_var == 0:
            self.vtm_var = None
        if self.vtm_yok == 0:
            self.vtm_yok = None
        if self.vtm_bilinmiyor == 0:
            self.vtm_bilinmiyor = None


        # KonumsalTutarlılık
     
        vk_konumsal = cnn.getlistofdata('x_ek_2_tucbs_veri_katmani', 'objectid, vk_konumsal_yeni',
                                    'geodurum = true and katman_durumu is true and ek_2='+str(self.ek2))
        self.vt_malti = 0
        self.vt_mustu = 0
        self.vtk_bilinmiyor = 0

        for row in vk_konumsal:
            if row[1] == 1:
                self.vt_malti += 1
            elif row[1] == 2:
                self.vt_mustu += 1
            elif row[1] == None or row[1] == 3:
                self.vtk_bilinmiyor += 1
        if self.vt_malti == 0:
            self.vt_malti = None
        if self.vt_mustu == 0:
            self.vt_mustu = None
        if self.vtk_bilinmiyor == 0:
            self.vtk_bilinmiyor = None


        # ZamansalTutarlılık
     
        vk_zamansal = cnn.getlistofdata('x_ek_2_tucbs_veri_katmani', 'objectid, vk_zamansal_yeni',
                                    'geodurum = true and katman_durumu is true and ek_2='+str(self.ek2))
        self.vtz_var = 0
        self.vtz_yok = 0
        self.vtz_bilinmiyor = 0

        for row in vk_zamansal:
            if row[1] == 1:
                self.vtz_var += 1
            elif row[1] == 2:
                self.vtz_yok += 1
            elif row[1] == None or row[1] == 3:
                self.vtz_bilinmiyor += 1
        if self.vtz_var == 0:
            self.vtz_var = None
        if self.vtz_yok == 0:
            self.vtz_yok = None
        if self.vtz_bilinmiyor == 0:
            self.vtz_bilinmiyor = None


        # TematikTutarlılık
     
        vk_tematik = cnn.getlistofdata('x_ek_2_tucbs_veri_katmani', 'objectid, vk_tematik_yeni',
                                    'geodurum = true and katman_durumu is true and ek_2='+str(self.ek2))
        self.vtt_var = 0
        self.vtt_yok = 0
        self.vtt_bilinmiyor = 0

        for row in vk_tematik:
            if row[1] == 1:
                self.vtt_var += 1
            elif row[1] == 2:
                self.vtt_yok += 1
            elif row[1] == None or row[1] == 3:
                self.vtt_bilinmiyor += 1
        if self.vtt_var == 0:
            self.vtt_var = None
        if self.vtt_yok == 0:
            self.vtt_yok = None
        if self.vtt_bilinmiyor == 0:
            self.vtt_bilinmiyor = None


        # VerilerinGüncelOlmaDurumu
     
        vk_guncel = cnn.getlistofdata('x_ek_2_tucbs_veri_katmani', 'objectid, vk_zamansal_gecerlilik_yeni',
                                    'geodurum = true and katman_durumu is true and ek_2='+str(self.ek2))
        self.vt_guncel = 0
        self.vt_guncel_degil = 0
        self.vtv_bilinmiyor = 0

        for row in vk_guncel:
            if row[1] == 1:
                self.vt_guncel += 1
            elif row[1] == 2:
                self.vt_guncel_degil += 1
            elif row[1] == None or row[1] == 3:
                self.vtv_bilinmiyor += 1
        if self.vt_guncel == 0:
            self.vt_guncel = None
        if self.vt_guncel_degil == 0:
            self.vt_guncel_degil = None
        if self.vtv_bilinmiyor == 0:
            self.vtv_bilinmiyor = None

        # WebServisDurumu

        vk_ogc = cnn.getlistofdata('x_ek_2_tucbs_veri_katmani', 'objectid, servis_wms_var, servis_wfs_var',
                                    'geodurum = true and katman_durumu is true and ek_2='+str(self.ek2))
        self.wms = 0
        self.wfs = 0
        self.total = 0
        self.wms_n = 0
        self.wfs_n = 0

        for row in vk_ogc:
            if row[1]:
                self.wms += 1
            if row[2]: 
                self.wfs += 1
            self.total +=1
        
        if self.wms == 0:
            self.wms = None
            self.wms_n = self.total
        else:
            self.wms_n = self.total - self.wms
        if self.wfs == 0:
            self.wfs = None
            self.wfs_n = self.total
        else:
            self.wfs_n = self.total - self.wfs
        if self.wms_n == 0:
            self.wms_n = None
        if self.wfs_n == 0:
            self.wfs_n = None
        
        # ProjeksiyonveDatum
        self.projeksiyonDatum = []
        allDatum = cnn.getlistofdata('kod_ek_2_projeksiyon p, kod_ek_2_datum d', 'p.objectid, d.objectid', 'true')
        
        pd_allDatum = [[row[0],row[1],0] for row in allDatum]


        pd_projeksiyonDatum = cnn.getlistofdata('x_ek_2_tucbs_veri_katmani', 'objectid, projeksiyon, datum',
                                    'geodurum = true and katman_durumu is true and veri_tipi = 1 and ek_2='+str(self.ek2))

        # print self.k_adi.decode('utf-8')

        for row in pd_projeksiyonDatum:
            for en,datum in enumerate(pd_allDatum):
                if row[1] == datum[0] and row[2] == datum[1]:
                    count = datum[2]
                    count += 1
                    pd_allDatum[en] = [datum[0],datum[1],count]

        last_pd = []
        last_num = []

        for en,row in enumerate(pd_allDatum):
            if row[2] != 0:
                p = cnn.getsinglekoddata('kod_ek_2_projeksiyon', 'kod', 'objectid='+str(row[0])).decode('utf-8')
                d = cnn.getsinglekoddata('kod_ek_2_datum', 'kod', 'objectid='+str(row[1])).decode('utf-8')
                c = row[2]
                last_pd.append([p,d,c])
                last_num.append(row)

        result = {}
        
        for row in last_pd:
            if row[0] in result:
                result[row[0]].append((row[1], row[2]))
            else:
                result[row[0]] = [(row[1], row[2])]
        
        self.result =  result

        # data = numpy.array(last_pd)
        self.df = None
        # if len(last_pd) > 0:
        #     self.df = pd.DataFrame(data=data[0:,2:],index=data[1:,0],columns=data[0:,0:])
        

        # MetaveriDurumu

        metaveri = cnn.getlistofdata('x_ek_2_tucbs_veri_katmani', 'objectid, mv_metaveri_var, mv_standart',
                                    'geodurum = true and katman_durumu is true and ek_2='+str(self.ek2))
        self.tucbs_mv = 0
        self.kurum_mv = 0
        self.ulusal_mv = 0
        self.yok_mv = 0

        for row in metaveri:
            if row[1]:
                if row[2] == 1: 
                    self.tucbs_mv +=1
                elif row[2] == 2:
                    self.ulusal_mv +=1 
                else:
                    self.kurum_mv += 1
            else:
                self.yok_mv += 1
        
        if self.yok_mv == 0:
            self.yok_mv = None

        if self.tucbs_mv == 0:
            self.tucbs_mv = None
        if self.kurum_mv == 0:
            self.kurum_mv = None
        if self.ulusal_mv == 0:
            self.ulusal_mv = None
    def save_excel(self):
        wb = self.wb
        wb.close()

# BasicGraph - VeriTürü 

    def veri_turu(self):
        wb = self.wb
        ws = wb.add_worksheet(u'VeriTuru')

        bold = wb.add_format({'bold': 1})

        # Add the worksheet data that the charts will refer to.
        headings = [u'Dijital Veri', u'Basılı Veri', u'Bilinmiyor']
        data = [
            [self.vt_dijital, self.vt_basili, self.vt_bilinmiyor]
        ]

        ws.write_row('A1', headings, bold)
        ws.write_row('A2', data[0])
        #
        # Create a stacked chart sub-type.
        #
        chart2 = wb.add_chart({'type': 'column'})

        # Configure the first series.
        if self.vt_bilinmiyor == None:
            chart2.add_series({
                'categories': '=VeriTuru!$A$1:$B$1',
                'values':     '=VeriTuru!$A$2:$B$2',
                'data_labels': {'value': True},
            })
        else:
            chart2.add_series({
                'categories': '=VeriTuru!$A$1:$C$1',
                'values':     '=VeriTuru!$A$2:$C$2',
                'data_labels': {'value': True},
            })

        # Add a chart title and some axis labels.
        chart2.set_title ({'name': u'Veri Türü'})
        # chart2.set_x_axis({'name': 'Test number'})
        chart2.set_y_axis({'name': 'Adet'})

        # Set an Excel chart style.
        chart2.set_style(12)

        # Insert the chart into the worksheet (with an offset).
        ws.insert_chart('D1', chart2, {'x_offset': 0, 'y_offset': 0})

# StackedGraph - VeriFormatı 


    def veriformati(self):
        wb = self.wb
        ws = wb.add_worksheet(u'VeriFormati')

        bold = wb.add_format({'bold': 1})

        # Add the worksheet data that the charts will refer to.
        headings = ['Veri', u'Dijital Veri', u'Basılı Veri']
        data = [
            ['NCZ, DWG', 'Raster', u'Veritabanı', 'Bilinmiyor'],
            [self.cad, self.raster_dij, self.vt, self.vf_bilinmiyor],
            [None, self.raster_bas, None, None],
        ]

        ws.write_row('A1', headings, bold)
        ws.write_column('A2', data[0])
        ws.write_column('B2', data[1])
        ws.write_column('C2', data[2])
        #
        # Create a stacked chart sub-type.
        #
        chart2 = wb.add_chart({'type': 'column', 'subtype': 'stacked'})

        # Configure the first series.
        chart2.add_series({
            'name':       '=VeriFormati!$B$1',
            'categories': '=VeriFormati!$A$2:$A$5',
            'values':     '=VeriFormati!$B$2:$B$5',
            'data_labels': {'value': True},
        })

        # Configure second series.
        chart2.add_series({
            'name':       '=VeriFormati!$C$1',
            'categories': '=VeriFormati!$A$2:$A$5',
            'values':     '=VeriFormati!$C$2:$C$5',
            'data_labels': {'value': True},
        })

        # Add a chart title and some axis labels.
        chart2.set_title ({'name': u'Veri Formatı'})
        # chart2.set_x_axis({'name': 'Test number'})
        chart2.set_y_axis({'name': u'Adet'})

        # Set an Excel chart style.
        chart2.set_style(12)

        # Insert the chart into the worksheet (with an offset).
        ws.insert_chart('E1', chart2, {'x_offset': 0, 'y_offset': 0})

# BasicGraph - VeriEksiksizlik 

    def veri_eksiksizlik(self):
        wb = self.wb
        ws = wb.add_worksheet(u'VeriEksizlik')

        bold = wb.add_format({'bold': 1})

        # Add the worksheet data that the charts will refer to.
        headings = [u'Eksik', u'Tam', u'Bilinmiyor']
        data = [
            [self.vt_eksik, self.vt_tam, self.vte_bilinmiyor]
        ]

        ws.write_row('A1', headings, bold)
        ws.write_row('A2', data[0])
        #
        # Create a stacked chart sub-type.
        #
        chart2 = wb.add_chart({'type': 'column'})

        # Configure the first series.
        if self.vte_bilinmiyor == None:
            chart2.add_series({
                'categories': '=VeriEksizlik!$A$1:$B$1',
                'values':     '=VeriEksizlik!$A$2:$B$2',
                'data_labels': {'value': True},
            })
        else:
            chart2.add_series({
                'categories': '=VeriEksizlik!$A$1:$C$1',
                'values':     '=VeriEksizlik!$A$2:$C$2',
                'data_labels': {'value': True},
            })

        # Add a chart title and some axis labels.
        chart2.set_title ({'name': u'Veri Eksiksizliği'})
        # chart2.set_x_axis({'name': 'Test number'})
        chart2.set_y_axis({'name': 'Adet'})

        # Set an Excel chart style.
        chart2.set_style(12)

        # Insert the chart into the worksheet (with an offset).
        ws.insert_chart('D1', chart2, {'x_offset': 0, 'y_offset': 0})

# BasicGraph - MantıksalTutarlılık

    def mantiksal_tutarlilik(self):
        wb = self.wb
        ws = wb.add_worksheet(u'MantiksalTutarlilik')

        bold = wb.add_format({'bold': 1})

        # Add the worksheet data that the charts will refer to.
        headings = [u'Var', u'Yok', u'Bilinmiyor']
        data = [
            [self.vtm_var, self.vtm_yok, self.vtm_bilinmiyor]
        ]

        ws.write_row('A1', headings, bold)
        ws.write_row('A2', data[0])
        #
        # Create a stacked chart sub-type.
        #
        chart2 = wb.add_chart({'type': 'column'})

        # Configure the first series.
        if self.vtm_bilinmiyor == None:
            chart2.add_series({
                'categories': '=MantiksalTutarlilik!$A$1:$B$1',
                'values':     '=MantiksalTutarlilik!$A$2:$B$2',
                'data_labels': {'value': True},
            })
        else:
            chart2.add_series({
                'categories': '=MantiksalTutarlilik!$A$1:$C$1',
                'values':     '=MantiksalTutarlilik!$A$2:$C$2',
                'data_labels': {'value': True},
            })

        # Add a chart title and some axis labels.
        chart2.set_title ({'name': u'Mantıksal Tutarlılık'})
        # chart2.set_x_axis({'name': 'Test number'})
        chart2.set_y_axis({'name': 'Adet'})

        # Set an Excel chart style.
        chart2.set_style(12)

        # Insert the chart into the worksheet (with an offset).
        ws.insert_chart('D1', chart2, {'x_offset': 0, 'y_offset': 0})

# BasicGraph - KonumsalTutarlılık

    def vk_konumsal(self):
        wb = self.wb
        ws = wb.add_worksheet(u'KonumsalDogruluk')

        bold = wb.add_format({'bold': 1})

        # Add the worksheet data that the charts will refer to.
        headings = [u'1m ve Altı', u'1m Üstü', u'Bilinmiyor']
        data = [
            [self.vt_malti, self.vt_mustu, self.vtk_bilinmiyor]
        ]

        ws.write_row('A1', headings, bold)
        ws.write_row('A2', data[0])
        #
        # Create a stacked chart sub-type.
        #
        chart2 = wb.add_chart({'type': 'column'})

        # Configure the first series.
        if self.vtk_bilinmiyor == None:
            chart2.add_series({
                'categories': '=KonumsalDogruluk!$A$1:$B$1',
                'values':     '=KonumsalDogruluk!$A$2:$B$2',
                'data_labels': {'value': True},
            })
        else:
            chart2.add_series({
                'categories': '=KonumsalDogruluk!$A$1:$C$1',
                'values':     '=KonumsalDogruluk!$A$2:$C$2',
                'data_labels': {'value': True},
            })

        # Add a chart title and some axis labels.
        chart2.set_title ({'name': u'Konumsal Doğruluk'})
        # chart2.set_x_axis({'name': 'Test number'})
        chart2.set_y_axis({'name': 'Adet'})

        # Set an Excel chart style.
        chart2.set_style(12)

        # Insert the chart into the worksheet (with an offset).
        ws.insert_chart('D1', chart2, {'x_offset': 0, 'y_offset': 0})

# BasicGraph - ZamansalTutarlılık

    def vk_zamansal(self):
        wb = self.wb
        ws = wb.add_worksheet(u'ZamansalDogruluk')

        bold = wb.add_format({'bold': 1})

        # Add the worksheet data that the charts will refer to.
        headings = [u'Var', u'Yok', u'Bilinmiyor']
        data = [
            [self.vtz_var, self.vtz_yok, self.vtz_bilinmiyor]
        ]

        ws.write_row('A1', headings, bold)
        ws.write_row('A2', data[0])
        #
        # Create a stacked chart sub-type.
        #
        chart2 = wb.add_chart({'type': 'column'})

        # Configure the first series.
        if self.vtz_bilinmiyor == None:
            chart2.add_series({
                'categories': '=ZamansalDogruluk!$A$1:$B$1',
                'values':     '=ZamansalDogruluk!$A$2:$B$2',
                'data_labels': {'value': True},
            })
        else:
            chart2.add_series({
                'categories': '=ZamansalDogruluk!$A$1:$C$1',
                'values':     '=ZamansalDogruluk!$A$2:$C$2',
                'data_labels': {'value': True},
            })

        # Add a chart title and some axis labels.
        chart2.set_title ({'name': u'Zamansal Doğruluk'})
        # chart2.set_x_axis({'name': 'Test number'})
        chart2.set_y_axis({'name': 'Adet'})

        # Set an Excel chart style.
        chart2.set_style(12)

        # Insert the chart into the worksheet (with an offset).
        ws.insert_chart('D1', chart2, {'x_offset': 0, 'y_offset': 0})

# BasicGraph - TematikTutarlılık

    def vk_tematik(self):
        wb = self.wb
        ws = wb.add_worksheet(u'TematikDogruluk')

        bold = wb.add_format({'bold': 1})

        # Add the worksheet data that the charts will refer to.
        headings = [u'Var', u'Yok', u'Bilinmiyor']
        data = [
            [self.vtt_var, self.vtt_yok, self.vtt_bilinmiyor]
        ]

        ws.write_row('A1', headings, bold)
        ws.write_row('A2', data[0])
        #
        # Create a stacked chart sub-type.
        #
        chart2 = wb.add_chart({'type': 'column'})

        # Configure the first series.
        if self.vtt_bilinmiyor == None:
            chart2.add_series({
                'categories': '=TematikDogruluk!$A$1:$B$1',
                'values':     '=TematikDogruluk!$A$2:$B$2',
                'data_labels': {'value': True},
            })
        else:
            chart2.add_series({
                'categories': '=TematikDogruluk!$A$1:$C$1',
                'values':     '=TematikDogruluk!$A$2:$C$2',
                'data_labels': {'value': True},
            })

        # Add a chart title and some axis labels.
        chart2.set_title ({'name': u'Tematik Doğruluk'})
        # chart2.set_x_axis({'name': 'Test number'})
        chart2.set_y_axis({'name': 'Adet'})

        # Set an Excel chart style.
        chart2.set_style(12)

        # Insert the chart into the worksheet (with an offset).
        ws.insert_chart('D1', chart2, {'x_offset': 0, 'y_offset': 0})

# BasicGraph - VerilerinGüncelOlmaDurumu

    def vk_guncel(self):
        wb = self.wb
        ws = wb.add_worksheet(u'VeriGuncelOlmaDurumu')

        bold = wb.add_format({'bold': 1})

        # Add the worksheet data that the charts will refer to.
        headings = [u'Güncel', u'Güncel Değil', u'Bilinmiyor']
        data = [
            [self.vt_guncel, self.vt_guncel_degil, self.vtv_bilinmiyor]
        ]

        ws.write_row('A1', headings, bold)
        ws.write_row('A2', data[0])
        #
        # Create a stacked chart sub-type.
        #
        chart2 = wb.add_chart({'type': 'column'})

        # Configure the first series.
        if self.vtv_bilinmiyor == None:
            chart2.add_series({
                'categories': '=VeriGuncelOlmaDurumu!$A$1:$B$1',
                'values':     '=VeriGuncelOlmaDurumu!$A$2:$B$2',
                'data_labels': {'value': True},
            })
        else:
            chart2.add_series({
                'categories': '=VeriGuncelOlmaDurumu!$A$1:$C$1',
                'values':     '=VeriGuncelOlmaDurumu!$A$2:$C$2',
                'data_labels': {'value': True},
            })

        # Add a chart title and some axis labels.
        chart2.set_title ({'name': u'Verilerin Güncel Olma Durumu'})
        # chart2.set_x_axis({'name': 'Test number'})
        chart2.set_y_axis({'name': 'Adet'})

        # Set an Excel chart style.
        chart2.set_style(12)

        # Insert the chart into the worksheet (with an offset).
        ws.insert_chart('D1', chart2, {'x_offset': 0, 'y_offset': 0})

# StackedGraph - WebServisDurumu

    def vk_ogc(self):
        workbook = self.wb
        worksheet = workbook.add_worksheet('WebServis')

        # Add the worksheet data that the charts will refer to.
        headings = [u'Servis', u'Yayınlanıyor', u'Yayınlanmıyor']
        data = [
            ['WMS', 'WFS'],
            [self.wms, self.wfs],
            [self.wms_n, self.wfs_n],
        ]
        bold = workbook.add_format({'bold': 1})

        worksheet.write_row('A1', headings, bold)
        worksheet.write_column('A2', data[0])
        worksheet.write_column('B2', data[1])
        worksheet.write_column('C2', data[2])
        chart3 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})

        # Configure the first series.
        chart3.add_series({
            'name':       '=WebServis!$B$1',
            'categories': '=WebServis!$A$2:$A$3',
            'values':     '=WebServis!$B$2:$B$3',
            'data_labels': {'value': True},
        })

        # Configure second series.
        chart3.add_series({
            'name':       '=WebServis!$C$1',
            'categories': '=WebServis!$A$2:$A$3',
            'values':     '=WebServis!$C$2:$C$3',
            'data_labels': {'value': True},
        })

        # Add a chart title and some axis labels.
        chart3.set_title ({'name': 'Web Servis Durumu'})
        chart3.set_y_axis({'name': 'Adet'})

        # Set an Excel chart style.
        chart3.set_style(12)

        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart('D1', chart3, {'x_offset': 0, 'y_offset': 0})
    
# StackedGraph - ProjeksiyonveDatum
    def projeksiyon_datum(self):
        wb = self.wb
        ws = wb.add_worksheet(u'ProjeksiyonDatum')

        result = self.result
        datum_position = {}
        max_colunm = 0
        columns=['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
        cur_row_num = 1
        for en,p in enumerate(result):
            cell = columns[en+1] + '1'
            ws.write(cell, p)
            if max_colunm < en+1:
                max_colunm = en+1
            datums = result[p]
            for d in datums:
                datum = d[0]
                count = d[1]
                if datum not in datum_position:
                    datum_position[datum] = cur_row_num
                    cur_row_num += 1
                    ws.write('A'+str(cur_row_num), datum)
                    ws.write(columns[en+1]+str(cur_row_num), count)
                else:
                    ws.write('A'+str(datum_position[datum]+1), datum)
                    ws.write(columns[en+1]+str(datum_position[datum]+1), count)
            
        # # Add the worksheet data that the charts will refer to.
        # headings = ['Veri', u'Dijital Veri', u'Basılı Veri']
        # data = [
        #     ['NCZ, DWG', 'Raster', u'Veritabanı', 'Bilinmiyor'],
        #     [self.cad, self.raster_dij, self.vt, self.vf_bilinmiyor],
        #     [None, self.raster_bas, None, None],
        # ]

        # ws.write_row('A1', headings, bold)
        # ws.write_column('A2', data[0])
        # ws.write_column('B2', data[1])
        # ws.write_column('C2', data[2])
        # #
        # # Create a stacked chart sub-type.
        # #
        if max_colunm > 0:    
            chart2 = wb.add_chart({'type': 'column', 'subtype': 'stacked'})

            # Configure the first series.

            for x in range(max_colunm):
                chart2.add_series({
                    'name':         '=ProjeksiyonDatum!$'+columns[x+1]+'1',
                    'categories':   '=ProjeksiyonDatum!$A$2:$A$'+str(cur_row_num),
                    'values':       '=ProjeksiyonDatum!$'+columns[x+1]+'2:$'+columns[x+1]+'$'+str(cur_row_num),
                    'data_labels': {'value': True},
                })

            # # Configure second series.
            # chart2.add_series({
            #     'name':       '=ProjeksiyonDatum!$C$1',
            #     'categories': '=ProjeksiyonDatum!$A$2:$A$5',
            #     'values':     '=ProjeksiyonDatum!$C$2:$C$5',
            #     'data_labels': {'value': True},
            # })

            # Add a chart title and some axis labels.
            chart2.set_title ({'name': u'Projeksiyon ve Datum'})
            # chart2.set_x_axis({'name': 'Test number'})
            chart2.set_y_axis({'name': u'Adet'})

            # # Set an Excel chart style.
            chart2.set_style(12)

            # # Insert the chart into the worksheet (with an offset).
            ws.insert_chart('E1', chart2, {'x_offset': 0, 'y_offset': 0})

            # BasicGraph-Metaveri Durumu

# StackedGraph - Metaveri
    def metaveri(self):
        wb = self.wb
        ws = wb.add_worksheet(u'MetaveriDurum')

        bold = wb.add_format({'bold': 1})

        # Add the worksheet data that the charts will refer to.
        headings = [u'TUCBS', u'Ulusal Metaveri Profili', u'Kurum', u'Yok']
        data = [
            [u'Var', u'Yok'],
            [self.tucbs_mv, self.ulusal_mv, self.kurum_mv, None],
            [None, None, None, self.yok_mv],
        ]

        ws.write_row('B1', headings, bold)
        ws.write_row('B2', data[1])
        ws.write_row('B3', data[2])
        ws.write_column('A2', data[0])
        #
        # Create a stacked chart sub-type.
        #
        chart2 = wb.add_chart({'type': 'column', 'subtype': 'stacked'})

        # Configure the first series.
        chart2.add_series({
            'name':       '=MetaveriDurum!$B$1',
            'categories': '=MetaveriDurum!$A$2:$A$3',
            'values':     '=MetaveriDurum!$B$2:$B$3',
            'data_labels': {'value': True},
        })

        # Configure second series.
        chart2.add_series({
            'name':       '=MetaveriDurum!$C$1',
            'categories': '=MetaveriDurum!$A$2:$A$3',
            'values':     '=MetaveriDurum!$C$2:$C$3',
            'data_labels': {'value': True},
        })
        chart2.add_series({
            'name':       '=MetaveriDurum!$D$1',
            'categories': '=MetaveriDurum!$A$2:$A$3',
            'values':     '=MetaveriDurum!$D$2:$D$3',
            'data_labels': {'value': True},
        })
        chart2.add_series({
            'name':       '=MetaveriDurum!$E$1',
            'categories': '=MetaveriDurum!$A$2:$A$3',
            'values':     '=MetaveriDurum!$E$2:$E$3',
            'data_labels': {'value': True},
        })

        # Add a chart title and some axis labels.
        chart2.set_title ({'name': u'Metaveri'})
        # chart2.set_x_axis({'name': 'Test number'})
        chart2.set_y_axis({'name': u'Adet'})

        # Set an Excel chart style.
        chart2.set_style(12)

        # Insert the chart into the worksheet (with an offset).
        ws.insert_chart('E1', chart2, {'x_offset': 0, 'y_offset': 0})