# -*- coding: utf-8 -*-
import xlsxwriter
from pgget import Connection
cnn = Connection()

class KurumTablosu:
    def __init__(self, oid):
        self.oid = oid
        self.k_adi = cnn.getSingledataByOid('kurum', 'k_adi', self.oid)
        self.wb = xlsxwriter.Workbook(self.k_adi.decode('utf-8')+'.xlsx')
        self.ek2 = cnn.getsinglekoddata('ek_2_cografi_veri_analizi', 'objectid', 'kurum='+str(self.oid))
        # veri türü
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

    def save_excel(self):
        wb = self.wb
        wb.close()

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
        #######################################################################
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


        
    def veri_formati(self):
        wb = self.wb
        ws = wb.add_worksheet(u'VeriFormati')

        bold = wb.add_format({'bold': 1})

        # Add the worksheet data that the charts will refer to.
        headings = ['Number', u'Dijital Veri', u'Basılı Veri']
        data = [
            ['Bilinmiyor', 'NCZ, DWG', 'Raster', u'Veritabanı'],
            [3, 3, 2, 15],
            [None, None, 1, 0]
        ]

        ws.write_row('A1', headings, bold)
        ws.write_column('A2', data[0])
        ws.write_column('B2', data[1])
        ws.write_column('C2', data[2])
        #######################################################################
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
        chart2.set_style(2)

        # Insert the chart into the worksheet (with an offset).
        ws.insert_chart('D1', chart2, {'x_offset': 0, 'y_offset': 0})