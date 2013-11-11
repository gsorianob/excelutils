# -*- coding: utf-8 -*-
from datetime import date, datetime, time
from django.http import HttpResponse
from django.test import TestCase
from excelutils.excel_manager import WriterExcelManager
from excelutils.excel_manager import RowAndColumnsMustBePositiveIntegersError
from excelutils.assertions import assertIsNotNone, assertRaises, assertIsInstance


class WriterExcelManagerTests(TestCase):

    def test_crear_excel_crea_el_libro(self):
        #Setup
        excel = WriterExcelManager()

        # Excersice
        libro = excel.workbook

        # Verify
        assertIsNotNone(libro)

    def test_crear_excel_lo_crea_con_una_hoja(self):
        #Setup
        excel = WriterExcelManager()

        # Excersice
        cantidad = excel.get_sheet_qty()

        # Verify
        self.assertEquals(1, cantidad)

    def test_crear_excel_deja_el_cursor_en_la_primer_posicion(self):
        #Setup
        excel = WriterExcelManager()

        # Excersice
        row, column = excel.get_active_cell()

        # Verify
        self.assertEquals((0, 0), (row, column))

    def test_crear_excel_deja_el_archivo_listo_para_usar(self):
        #Setup
        excel = WriterExcelManager()

        # Excersice
        excel.write('Hola mundo')

        # Verify
        # que no haya lanzado ninguna excepción.

    def test_write_mueve_el_cursor_a_la_derecha(self):
        #Setup
        excel = WriterExcelManager()

        # Excersice
        excel.write('Hola mundo')

        # Verify
        row, column = excel.get_active_cell()
        self.assertEquals((0, 1), (row, column))

    def test_puedo_escribir_70000_filas(self):
        """ El máximo de filas en un excel es de 65536.
        """
        #Setup
        excel = WriterExcelManager()

        # Excersice
        for i in xrange(70000):
            excel.append_row([i])

        # Verify
        cantidad = excel.get_sheet_qty()
        self.assertEquals(2, cantidad)

    def test_puedo_escribir_1000_columnas(self):
        """ El máximo de columnas en un excel es de 65536.
        """
        #Setup
        excel = WriterExcelManager()

        # Excersice
        for i in xrange(70000):
            excel.write(i)

        # Verify
        cantidad = excel.get_sheet_qty()
        self.assertEquals(1, cantidad)

    def test_write_puede_mover_el_cursor_para_abajo(self):
        #Setup
        excel = WriterExcelManager()

        # Excersice
        excel.write('Hola mundo', continue_in_new_line=True)

        # Verify
        row, column = excel.get_active_cell()
        self.assertEquals((1, 0), (row, column))

    def test_set_active_cell_cambia_la_celda_activa(self):
        #Setup
        excel = WriterExcelManager()

        # Excersice
        excel.set_active_cell(7, 5)

        # Verify
        row, column = excel.get_active_cell()
        self.assertEquals((7, 5), (row, column))

    def test_set_active_cell_lanza_excepcion_si_fila_es_menor_a_cero(self):
        #Setup
        excel = WriterExcelManager()

        # Excersice

        # Verify
        assertRaises(RowAndColumnsMustBePositiveIntegersError, excel.set_active_cell, -1, 5)

    def test_set_active_cell_lanza_excepcion_si_columna_es_menor_a_cero(self):
        #Setup
        excel = WriterExcelManager()

        # Excersice

        # Verify
        assertRaises(RowAndColumnsMustBePositiveIntegersError, excel.set_active_cell, 1, -5)

    def test_set_active_cell_lanza_excepcion_si_fila_y_columna_son_negativos(self):
        #Setup
        excel = WriterExcelManager()

        # Excersice

        # Verify
        assertRaises(RowAndColumnsMustBePositiveIntegersError, excel.set_active_cell, -101, -55)

    def test_agrego_una_hoja_y_get_sheet_qty_me_retorna_2(self):
        #Setup
        excel = WriterExcelManager()

        # Excersice
        excel.add_sheet()

        # Verify
        cantidad = excel.get_sheet_qty()
        self.assertEquals(2, cantidad)

    def test_write_row_mueve_el_cursor_para_abajo(self):
        #Setup
        excel = WriterExcelManager()

        # Excersice
        excel.write_row([1, 2, 3])

        # Verify
        row, column = excel.get_active_cell()
        self.assertEquals((1, 0), (row, column))

    def test_write_row_mueve_el_cursor_uno_para_abajo_por_mas_que_se_encuentre_en_otra_posicion(self):
        #Setup
        excel = WriterExcelManager()
        excel.set_active_cell(0, 5)

        # Excersice
        excel.write_row([1, 2, 3])

        # Verify
        row, column = excel.get_active_cell()
        self.assertEquals((1, 0), (row, column))

    def test_append_row_mueve_el_cursor_para_abajo(self):
        #Setup
        excel = WriterExcelManager()

        # Excersice
        excel.append_row([1, 2, 3])

        # Verify
        row, column = excel.get_active_cell()
        self.assertEquals((1, 0), (row, column))

    def test_append_row_mueve_el_cursor_dos_para_abajo_si_no_se_encuentra_en_la_columna_0(self):
        #Setup
        excel = WriterExcelManager()
        excel.set_active_cell(0, 5)

        # Excersice
        excel.append_row([1, 2, 3])

        # Verify
        row, column = excel.get_active_cell()
        self.assertEquals((2, 0), (row, column))

    def test_puedo_imprimir_multiples_tipos_de_dato_y_no_falla(self):
        #Setup
        excel = WriterExcelManager()
        excel.set_active_cell(0, 5)

        # Excersice
        excel.write('string')
        excel.write(u'unicode')
        excel.write('áéíóú'.decode('latin1'))
        excel.write('áéíóú'.decode('utf-8'))
        excel.write('áéíóú'.decode('utf-16'))
        excel.write(unichr(40960))
        excel.write(1)
        excel.write(1L)
        excel.write(date(2012, 12, 13))
        excel.write(datetime(2012, 12, 13, 15, 15))
        excel.write(time(15, 15))
        excel.write(1.5)
        excel.write(None)
        excel.write([1, 2, 3])
        excel.write((1, 2, 3))
        excel.write({1: 1, 2: 2})

        # Verify
        # el hecho de no lanzar una excepción

    def test_convert_real_to_str_with_2_decimals_its_ok(self):
        #Setup
        excel = WriterExcelManager()

        # Excersice
        to_string = excel.convert_real_to_str_with_decimals(2)
        result = to_string(2 / 3.0)

        # Verify
        self.assertEquals('0.67', result)

    def test_convert_real_to_str_with_5_decimals_its_ok(self):
        #Setup
        excel = WriterExcelManager()

        # Excersice
        to_string = excel.convert_real_to_str_with_decimals(5)
        result = to_string(2 / 3.0)

        # Verify
        self.assertEquals('0.66667', result)

    def test_download_response_returns_a_HttpResponse(self):
        #Setup
        excel = WriterExcelManager()
        rows = []

        # Excersice
        result = excel.download_response(rows)

        # Verify
        assertIsInstance(result, HttpResponse)

    def test_download_response_returns_a_HttpResponse_with_content_dispoition(self):
        #Setup
        excel = WriterExcelManager()
        rows = []

        # Excersice
        result = excel.download_response(rows)

        # Verify
        content_disposition_expected = 'attachment; filename=file.xls'
        self.assertEquals(result['Content-Disposition'], content_disposition_expected)

    def test_download_response_returns_a_HttpResponse_with_content_type(self):
        #Setup
        excel = WriterExcelManager()
        rows = []

        # Excersice
        result = excel.download_response(rows)

        # Verify
        content_disposition_expected = 'application/vnd.ms-excel'
        self.assertEquals(result['Content-Type'], content_disposition_expected)

    def test_download_response_returns_a_HttpResponse_with_excel_file(self):
        #Setup
        excel = WriterExcelManager()
        rows = []

        # Excersice
        result = excel.download_response(rows)

        # Verify
        excel_file = excel.get_for_download()
        self.assertEquals(result.content, excel_file)

    def test_ajax_download_response_returns_a_JSON_HttpResponse(self):
        #Setup
        excel = WriterExcelManager()
        rows = []

        # Excersice
        result = excel.ajax_download_response(rows, '/tmp/', '')

        # Verify
        assertIsInstance(result, HttpResponse)

    def test_ajax_download_response_returns_a_JSON_HttpResponse_with_content_type(self):
        #Setup
        excel = WriterExcelManager()
        rows = []

        # Excersice
        result = excel.ajax_download_response(rows, '/tmp/', '')

        # Verify
        content_disposition_expected = 'text/html; charset=utf-8'
        self.assertEquals(result['Content-Type'], content_disposition_expected)

    def test_ajax_download_response_returns_a_JSON_HttpResponse_with_dict(self):
        #Setup
        excel = WriterExcelManager()
        rows = []

        # Excersice
        result = excel.ajax_download_response(rows, '/tmp/', '')

        # Verify
        excel_file = excel.get_for_download()
        self.assertEquals(result.content, '{"export_path": "file.xls"}')
