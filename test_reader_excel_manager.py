# -*- coding: utf-8 -*-
from django.test import TestCase
from excelutils.excel_manager import ReaderExcelManager, BookDoesNotExistsError, WrongVersionError
from excelutils.excel_manager import RowAndColumnsMustBePositiveIntegersError
from excelutils.utils_decorators import ignore_this_test


class ReaderManagerTests(TestCase):

    def test_si_no_existe_el_libro_lanza_la_excepcion_de_BookDoesNotExistsError(self):
        #Setup
        filename = 'Este archivo no existe.xls'

        # Excersice

        # Verify
        self.assertRaises(BookDoesNotExistsError, ReaderExcelManager, filename)

    @ignore_this_test
    def test_si_es_un_excel_de_2007_lanza_la_excepcion_de_BookDoesNotExistsError(self):
        #Setup
        filename = 'excelutils/archivo_excel_2007.xlsx'

        # Excersice

        # Verify
        self.assertRaises(WrongVersionError, ReaderExcelManager, filename)

    def test_si_es_un_excel_de_2003_NO_lanza_ninguna_excepcion(self):
        #Setup
        filename = 'excelutils/excel_ejemplo.xls'

        # Excersice
        excel = ReaderExcelManager(filename)

        # Verify
        # El hecho de que no haya lanzado una excepción es suficiente.

    def test_set_active_cell_lanza_excepcion_si_fila_es_menor_a_cero(self):
        #Setup
        filename = 'excelutils/excel_ejemplo.xls'
        excel = ReaderExcelManager(filename)

        # Excersice

        # Verify
        self.assertRaises(RowAndColumnsMustBePositiveIntegersError, excel.set_active_cell, -1, 5)

    def test_set_active_cell_lanza_excepcion_si_columna_es_menor_a_cero(self):
        #Setup
        filename = 'excelutils/excel_ejemplo.xls'
        excel = ReaderExcelManager(filename)

        # Excersice

        # Verify
        self.assertRaises(RowAndColumnsMustBePositiveIntegersError, excel.set_active_cell, 1, -5)

    def test_set_active_cell_lanza_excepcion_si_fila_y_columna_son_negativos(self):
        #Setup
        filename = 'excelutils/excel_ejemplo.xls'
        excel = ReaderExcelManager(filename)

        # Excersice

        # Verify
        self.assertRaises(RowAndColumnsMustBePositiveIntegersError, excel.set_active_cell, -101, -55)

    def test_me_muevo_con_set_active_y_al_leer_me_devuelve_el_resultado_esperado(self):
        #Setup
        filename = 'excelutils/excel_ejemplo.xls'
        excel = ReaderExcelManager(filename)
        excel.set_active_cell(4,0)
        expected = (u'fila 5', ReaderExcelManager.TYPE_TEXT)

        # Excersice
        result = excel.read()

        # Verify
        self.assertEquals(result, expected)

    def test_leo_un_string_y_me_devuelve_el_tipo_text(self):
        #Setup
        filename = 'excelutils/excel_ejemplo.xls'
        excel = ReaderExcelManager(filename)
        excel.set_active_cell(4,0)
        expected = (u'fila 5', ReaderExcelManager.TYPE_TEXT)

        # Excersice
        result = excel.read()

        # Verify
        self.assertEquals(result, expected)

    def test_leo_un_numero_con_punto_y_me_devuelve_el_tipo_float(self):
        #Setup
        filename = 'excelutils/excel_ejemplo.xls'
        excel = ReaderExcelManager(filename)
        excel.set_active_cell(6,1)
        expected = (2.5, ReaderExcelManager.TYPE_FLOAT)

        # Excersice
        result = excel.read()

        # Verify
        self.assertEquals(result, expected)

    def test_leo_un_numero_entero_y_me_devuelve_el_tipo_float(self):
        #Setup
        filename = 'excelutils/excel_ejemplo.xls'
        excel = ReaderExcelManager(filename)
        excel.set_active_cell(3,1)
        expected = (1.0, ReaderExcelManager.TYPE_FLOAT)

        # Excersice
        result = excel.read()

        # Verify
        self.assertEquals(result, expected)

    def test_leo_un_numero_con_coma_y_me_devuelve_el_tipo_text(self):
        #Setup
        filename = 'excelutils/excel_ejemplo.xls'
        excel = ReaderExcelManager(filename)
        excel.set_active_cell(5,1)
        expected = (u'1,5', ReaderExcelManager.TYPE_TEXT)

        # Excersice
        result = excel.read()

        # Verify
        self.assertEquals(result, expected)

    def test_leo_un_valor_y_el_cursor_se_mueve_uno_a_la_derecha(self):
        #Setup
        filename = 'excelutils/excel_ejemplo.xls'
        excel = ReaderExcelManager(filename)

        # Excersice
        excel.read()

        # Verify
        row, column = excel.get_active_cell()
        self.assertEquals((0, 1), (row, column))

    def test_leo_un_valor_y_el_cursor_se_mueve_a_la_primer_celda_de_la_fila_siguiente_si_estoy_en_la_ultima_celda_de_la_fila(self):
        #Setup
        filename = 'excelutils/excel_ejemplo.xls'
        excel = ReaderExcelManager(filename)
        excel.set_active_cell(0,3)

        # Excersice
        excel.read()

        # Verify
        row, column = excel.get_active_cell()
        self.assertEquals((1, 0), (row, column))

    def test_go_to_first_cell_in_row_funciona_ok(self):
        #Setup
        filename = 'excelutils/excel_ejemplo.xls'
        excel = ReaderExcelManager(filename)
        excel.set_active_cell(0,3)

        # Excersice
        excel.go_to_first_cell_in_row()

        # Verify
        row, column = excel.get_active_cell()
        self.assertEquals((0, 0), (row, column))

    def test_go_to_first_column_in_row_funciona_ok(self):
        #Setup
        filename = 'excelutils/excel_ejemplo.xls'
        excel = ReaderExcelManager(filename)
        excel.set_active_cell(3,0)

        # Excersice
        excel.go_to_first_cell_in_column()

        # Verify
        row, column = excel.get_active_cell()
        self.assertEquals((0, 0), (row, column))

    def test_reset_sheet_funciona_ok(self):
        #Setup
        filename = 'excelutils/excel_ejemplo.xls'
        excel = ReaderExcelManager(filename)
        excel.set_active_cell(3,3)

        # Excersice
        excel.reset_sheet()

        # Verify
        row, column = excel.get_active_cell()
        self.assertEquals((0, 0), (row, column))

    def test_la_lectura_funcion_correctamente(self):
        #Setup
        expected = [[(u'Excel de prueba', 1), ('', 0), ('', 0), ('', 0)],
                    [(u'fila 2', 1), ('', 0), ('', 0), (u'Columna 4', 1)],
                    [(u'fila 3', 1), ('', 0), ('', 0), ('', 0)],
                    [(u'fila 4', 1), (1.0, 2), ('', 0), ('', 0)],
                    [(u'fila 5', 1), (1.5, 2), ('', 0), ('', 0)],
                    [(u'fila 6', 1), (u'1,5', 1), ('', 0), ('', 0)],
                    [(u'fila 7', 1), (2.5, 2), ('', 0), ('', 0)],
                    [(u'fila 8', 1), ('', 0), ('', 0), ('', 0)],
                    [(u'fila 9', 1), ('', 0), ('', 0), ('', 0)],
                    [(u'fila 10', 1), ('', 0), ('', 0), ('', 0)],
                    [(u'fila 11', 1), ('', 0), ('', 0), ('', 0)],
                    [(u'fila 12', 1), ('', 0), ('', 0), ('', 0)],
                    [(u'fila 13', 1), ('', 0), ('', 0), ('', 0)],
                    [(u'fila 14', 1), ('', 0), ('', 0), ('', 0)],
                    [(u'fila 15', 1), ('', 0), ('', 0), ('', 0)]]
        filename = 'excelutils/excel_ejemplo.xls'
        excel = ReaderExcelManager(filename)

        # Excersice
        result = excel.read_all_sheet()

        # Verify
        self.assertEquals(result, expected)

