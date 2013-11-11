# -*- coding: utf-8 -*-
import datetime
import simplejson
import xlrd
import xlwt

from cStringIO import StringIO
from django.http import HttpResponse
from os import path
from excelutils.string_utils import safe_unicode_to_str

class ExcelException(Exception):
    pass

class BookDoesNotExistsError(ExcelException):

    def __unicode__(self):
        return u'No existe el archivo Excel, primero debe crear el libro.'

    def __str__(self):
        return unicode(self).encode('utf-8')


class EmptyBookError(ExcelException):

    def __unicode__(self):
        return u'El archivo se encuentra vacío, por lo que primero debe agregarle una hoja.'

    def __str__(self):
        return unicode(self).encode('utf-8')


class SheetDoesNotExistsError(ExcelException):

    def __init__(self, sheet_qty):
        self.sheet_qty = sheet_qty

    def __unicode__(self):
        return u'La hoja que esta buscando no se encuentra. El archivo sólo posee %s hojas' % self.sheet_qty

    def __str__(self):
        return unicode(self).encode('utf-8')


class CurrentSheetIsFullError(ExcelException):

    def __unicode__(self):
        msg = u'La hoja en la que quiere escribir se encuentra llena. ' \
              u'Si desea agregar más información tiene que agregar una página'
        return msg

    def __str__(self):
        return unicode(self).encode('utf-8')


class RowAndColumnsMustBePositiveIntegersError(ExcelException):

    def __init__(self, row_number, column_number):
        self.row_number = row_number
        self.column_number = column_number

    def __unicode__(self):
        msg = u'La fila y la columna deben ser mayor o igual a cero. Se encontraron: fila: "%s" y columna: "%s"'
        return msg % (self.row_number, self.column_number)

    def __str__(self):
        return unicode(self).encode('utf-8')


class WrongVersionError(ExcelException):

    def __unicode__(self):
        return u'Esta versión es compatible con archivos excel previos a Excel 2010.'

    def __str__(self):
        return unicode(self).encode('utf-8')


class WriterExcelManager:
    workbook = None
    __sheet_count = 0
    __sheets = []
    __current_sheet = None
    __current_row = None
    __current_column = None
    MAX_ROW = 65536
    MAX_COLUMN = 256

    def __init__(self, encoding="utf-8", sheet_name=None):
        '''Crea el archivo excel con el encoding especificado (utf-8 por
        defecto) y le agrega una hoja con el nombre que recibe por parámetro
        (por defecto, 'Sheet 1').
        @encoding: Codificación de caracteres que usará el archivo.
        Por defecto 'utf-8', aunque se podría pasar 'latin1' o cualquier otra.
        @sheet_name: Nombre de la primer hoja del archivo. Por defecto, se
        usará 'sheet_name'.
        '''
        self.workbook = xlwt.Workbook(encoding=encoding)
        self.add_sheet(sheet_name)
        self.writers = {
            type(datetime.datetime): self.__convert_datetime_to_str,
            type(datetime.date): self.__convert_date_to_str,
            type(datetime.time): self.__convert_time_to_str,
            unicode: safe_unicode_to_str,
            type(None): lambda value: '',
            int: lambda value: value,
            float: lambda value: value,
            long: lambda value: value
        }

    def __convert_date_to_str(self, fecha):
        """Escribe una fecha con el formato dd-mm-aaaa.
        """
        return fecha.srftime('%d-%m-%Y')

    def __convert_datetime_to_str(self, fecha_hora):
        """Escribe una fecha con el formato dd-mm-aaaa hh:mm donde hh
        tiene un rango de 00 a 24.
        """
        return fecha_hora.srftime('%d-%m-%Y %H:%M')

    def __convert_time_to_str(self, tiempo):
        """Escribe una hora con el formato hh:mm donde hh tiene un rango
        de 00 a 24.
        """
        return tiempo.srftime('%H:%M')

    def convert_real_to_str_with_decimals(self, decimals):
        """Crea un nuevo 'writer' para escribir números reales redondeando
        a una X cantidad de decimales.
        @decimals: Indica la cantidad de decimales que se mostraran en cada
        celda (X).
        """
        def convert_real_to_str(value):
            return str(round(value, decimals))
        return convert_real_to_str

    def add_writer(self, type, writer):
        """Agrega un método o función a los existentes para que esta clase
        conozca cómo escribir cada tipo de dato. Por defecto usa str.
        """
        self.writers[type] = writer

    def add_sheet(self, sheet_name=None):
        """Agrega una nueva hoja al archivo con el nombre que le pasen o
        'Sheet nro' donde nro es el número de hoja que se inserta si no le
        pasan ninguno.
        @sheet_name: Nombre de la hoja, 'Sheet nro' por defecto.
        """
        if not sheet_name:
            sheet_name = 'Sheet %s' % (self.__sheet_count + 1, )

        try:
            self.__current_sheet = self.workbook.add_sheet(sheet_name)
        except AttributeError:
            raise BookDoesNotExistsError

        self.__sheet_count += 1
        self.__current_row = 0
        self.__current_column = 0
        self.__sheets.append(self.__current_sheet)

    def get_sheet_qty(self):
        """Retorna la cantidad de hojas que tiene el libro.
        """
        return self.__sheet_count

    def set_active_sheet(self, idx):
        """Cambia la hoja activa del archivo excel.
        @idx: Indica el número de hoja que se usará a partir de este
        momento.
        """
        try:
            self.workbook.set_active_sheet(idx)
        except AttributeError:
            raise BookDoesNotExistsError
        except IndexError:
            raise SheetDoesNotExistsError(self.__sheet_count)
        self.__current_sheet = self.workbook.get_active_sheet()

    def get_active_cell(self):
        """Retorna la posición de la celda activa como una tupla de
        (fila, columna).
        """
        return self.__current_row, self.__current_column

    def set_active_cell(self, row_number, column_number):
        """Cambia la posición en la que se encuentre el cursor dentro de la
        misma hoja.
        @row_number: Fila de la nueva celda.
        @column_number: Columna de la nueva celda.
        """
        if row_number < 0 or column_number < 0:
            raise RowAndColumnsMustBePositiveIntegersError(row_number, column_number)
        self.__current_row = row_number
        self.__current_column = column_number

    def write(self, value, continue_in_new_line=False):
        """Escribe de una forma legible el valor pasado en la celda que se
        encuentra del archivo. Se pueden pisar o definir nuevas formas de
        escribir en el archivo usando el método add_writer.
        @value: Valor a escribir en el archivo.
        @continue_in_new_line: Indica si una vez escrito el valor en el archivo
        continúa en una nueva línea o en la celda que le sigue a la derecha.
        """
        if self.workbook is None:
            raise BookDoesNotExistsError
        writer = self.writers.get(type(value), str)
        try:
            self.__current_sheet.write(
                self.__current_row,
                self.__current_column,
                writer(value)
            )
        except AttributeError:
            raise EmptyBookError
        except ValueError:
            raise CurrentSheetIsFullError
        if continue_in_new_line:
            self.__current_row += 1
            self.__current_column = 0
        else:
            self.__current_column += 1
        if self.__current_column >= self.MAX_COLUMN:
            self.__current_row += 1
            self.__current_column = 1

    def write_row(self, row_values):
        """Agrega, desde donde se encuentre parado, cada valor de row_values
        uno a continuación del otro (cada uno en una celda distinta) hasta
        llegar al último y luego hace un "salto de línea".
        """
        for value in row_values:
            writer = self.writers.get(type(value), str)
            try:
                self.__current_sheet.write(
                    self.__current_row,
                    self.__current_column,
                    writer(value)
                )
            except AttributeError:
                raise EmptyBookError
            self.__current_column += 1
        self.__current_row += 1
        self.__current_column = 0

    def append_row(self, row_values):
        """Verifica que se encuentre en la primer celda de una fila (si esto
        no se cumple baja una fila y se para en la primer columna) y llama al
        método write_row.
        """
        if self.workbook is None:
            raise BookDoesNotExistsError
        if self.__current_column != 0:
            self.__current_row += 1
            self.__current_column = 0
        if self.__current_row == self.MAX_ROW:
            self.add_sheet()
        self.write_row(row_values)

    def save(self, filename):
        """Guarda el archivo excel con el nombre que reciba.
        Si se le pasa algo del tipo StringIO lo guarda en ese "buffer".
        """
        self.workbook.save(filename)

    def get_for_download(self):
        """Guarda el archivo en un buffer para ser devuelto en un response.
        """
        output = StringIO()
        self.workbook.save(output)
        output.seek(0)
        return output.read()

    @staticmethod
    def download_response(rows, filename='file.xls', sheet_name=None, writers={}):
        """Método estático que permite crear un archivo excel a partir
        de una lista de filas con un nombre del archivo como de la hoja
        y retorna directamente el response correspondiente para que se
        pueda usar en una view.
        Utiliza el encoding utf-8 para el archivo.
        @rows: Lista de lista de valores que se escribiran en las celdas.
        @filename: Nombre del archivo, por defecto 'file.xls'.
        @sheet_name: Nombre de la única hoja que tendrá el archivo, por
        defecto 'Sheet 1'
        @writers: Actualiza los writers por lo que se puede redefinir la
        forma en que se escribe un tipo de dato en particular.
        Por defecto, vacío.
        """
        excel = WriterExcelManager(sheet_name=sheet_name)
        excel.writers.update(writers)
        for row in rows:
            excel.write_row(row)

        content_type = 'application/vnd.ms-excel'
        response = HttpResponse(excel.get_for_download(), mimetype=content_type)
        response['Content-Disposition'] = 'attachment; filename=%s' % filename

        return response

    @staticmethod
    def ajax_download_response(rows, local_path, export_path, filename='file.xls', sheet_name=None, writers={}):
        """Método estático que permite crear un archivo excel a partir
        de una lista de filas, guardarlo en el disco del servidor,  con un
        nombre del archivo como de la hoja y retorna directamente el
        response correspondiente para que se pueda usar en una view.
        Utiliza el encoding utf-8 para el archivo.
        Se usa cuando se necesita bloquear la pantalla con el blockUI de
        jQuery.
        @rows: Lista de lista de valores que se escribiran en las celdas.
        @local_path: Directorio en el servidor donde se guardará el archivo.
        @export_path: URL desde la que se descargará el archivo.
        @filename: Nombre del archivo, por defecto 'file.xls'.
        @sheet_name: Nombre de la única hoja que tendrá el archivo, por
        defecto 'Sheet 1'
        @writers: Actualiza los writers por lo que se puede redefinir la
        forma en que se escribe un tipo de dato en particular.
        Por defecto, vacío.
        """
        local_filename = path.join(local_path, filename)
        export_filename = path.join(export_path, filename)
        excel = WriterExcelManager(sheet_name=sheet_name)
        excel.writers.update(writers)
        for row in rows:
            excel.write_row(row)

        excel.save(local_filename)
        data = {
            'export_path': export_filename
        }

        response = HttpResponse(simplejson.dumps(data))
        return response


class ReaderExcelManager:
    workbook = None
    __sheet_count = 0
    __sheets_names = []
    __current_sheet = None
    __current_row = None
    __current_column = None
    TYPE_EMPTY = 0
    TYPE_TEXT = 1
    TYPE_FLOAT = 2


    def _get_columns_quantity(self):
        if not self.__current_sheet:
            raise EmptyBookError
        return self.__current_sheet.ncols

    def _get_rows_quantity(self):
        if not self.__current_sheet:
            raise EmptyBookError
        return self.__current_sheet.nrows

    columns_quantity = property(_get_columns_quantity)
    rows_quantity = property(_get_rows_quantity)


    def __init__(self, filename, sheet_number=None):
        try:
            self.workbook = xlrd.open_workbook(filename)
        except IOError:
            raise BookDoesNotExistsError
        except xlrd.XLRDError:
            raise WrongVersionError
        self.__sheets_names = self.workbook.sheet_names()
        self.__sheet_count = len(self.__sheets_names)
        if self.__sheet_count == 0:
            raise EmptyBookError
        if sheet_number is None:
            sheet_number = 0
        if sheet_number >= self.__sheet_count:
            raise SheetDoesNotExistsError(self.__sheet_count)
        if self.__sheet_count:
            self.__current_sheet = self.workbook.sheet_by_index(sheet_number)
        self.__current_column = 0
        self.__current_row = 0

    def reset_sheet(self):
        self.__current_column = 0
        self.__current_row = 0

    def go_to_first_cell_in_row(self):
        self.__current_column = 0

    def go_to_first_cell_in_column(self):
        self.__current_row = 0

    def get_active_cell(self):
        """Retorna la posición de la celda activa como una tupla de
        (fila, columna).
        """
        return self.__current_row, self.__current_column

    def set_active_cell(self, row_number, column_number):
        """Cambia la posición en la que se encuentre el cursor dentro de la
        misma hoja.
        @row_number: Fila de la nueva celda.
        @column_number: Columna de la nueva celda.
        """
        if row_number < 0 or column_number < 0:
            raise RowAndColumnsMustBePositiveIntegersError(row_number, column_number)
        self.__current_row = row_number
        self.__current_column = column_number

    def _read_cell(self, row_number, column_number):
        try:
            value = self.__current_sheet.cell_value(row_number, column_number)
            type =  self.__current_sheet.cell_type(row_number, column_number)
        except IndexError:
            import pdb
            pdb.set_trace()
            raise
        return value, type

    def read(self):
        value_type = self._read_cell(self.__current_row, self.__current_column)
        self.__current_column += 1
        if self.__current_column >= self.columns_quantity:
            self.__current_row += 1
            self.__current_column = 0
        return value_type

    def read_current_row(self):
        row = []
        current_row = self.__current_row
        while current_row == self.__current_row:
            row.append(self.read())
        return row

    def read_sheet(self):
        rows = []
        while self.__current_row < self.rows_quantity:
            rows.append(self.read_current_row())
        return rows

    def read_all_sheet(self):
        self.reset_sheet()
        return self.read_sheet()

    def change_sheet_by_index(self, sheet_number):
        self.__current_sheet = self.workbook.sheet_by_index(sheet_number)

    def change_sheet_by_name(self, sheet_name):
        self.__current_sheet = self.workbook.sheet_by_name(sheet_name)
