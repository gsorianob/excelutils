#encoding: utf-8
import unicodedata


def safe_unicode_to_str(text, encoding='utf-8', errors='ignore'):
    """ Convierte de forma segura un unicode a string.
    @string: String a convertir.
    @encoding: indica el encoding con el que se intentará convertir el string.
    Por defecto es 'utf-8', pero también puede ser 'latin1' o cualquier otro.
    @errors: Indica qué hacer con los errores que aparezcan. Las opciones
    pueden ser:
        - 'ignore': Reemplaza los caracteres que no pueda mostrar por aquellos
        a los que se asemejan, por ejemplo reemplaza la 'é' por 'e', etc.
        - 'xmlcharrefreplace': Reemplaza los caracteres que no pueda mostrar
        por su código en xml, por ejemplo reemplaza la 'é' por '#769;', etc.
    Por defecto es 'ignore'.
    """
    if isinstance(text, str):
        return text
    return unicodedata.normalize('NFKD', text).encode(encoding, errors)


def safe_str_to_unicode(text, encoding='utf-8', errors='ignore'):
    """ Convierte de forma segura un string a unicode.
    @string: Unicode a convertir.
    @encoding: indica el encoding con el que se intentará convertir el string.
    Por defecto es 'utf-8', pero también puede ser 'latin1' o cualquier otro.
    @errors: Indica qué hacer con los errores que aparezcan. Las opciones
    pueden ser:
        - 'ignore': Reemplaza los caracteres que no pueda mostrar por aquellos
        a los que se asemejan, por ejemplo reemplaza la 'é' por 'e', etc.
        - 'xmlcharrefreplace': Reemplaza los caracteres que no pueda mostrar
        por su código en xml, por ejemplo reemplaza la 'é' por '#769;', etc.
    Por defecto es 'ignore'.
    """
    if isinstance(text, unicode):
        return text
    return text.decode(encoding, errors)


def encode_str_decode_to_unicode(text, str_encoding='utf-8', unicoce_encoding='utf-8', errors='ignore'):
    if isinstance(text, str):
        text = text.encode(str_encoding)
    return safe_str_to_unicode(text, unicoce_encoding, errors)


def agregar_espacios_luego_de_cada_coma_y_cada_punto(string=''):
    return string.replace(',', ', ').replace('.', '. ')


def seaparar_palabras_por_pipes_parentesis_guiones_puntos_y_comas(string=''):
    return agregar_espacios_luego_de_cada_coma_y_cada_punto(string)\
        .replace('(', ' (').replace(')', ') ').replace('|', ' | ').replace('-', ' - ')


def get_short_text(text, limit=100, ending=' ...'):
    if not text:
        return ''
    post_length = len(ending)
    text_length = limit - post_length
    words = text.split(' ')
    # start with de first word.
    short_text = words.pop(0)
    while words and len(short_text + ' %s' % words[0]) <= text_length:
        # append the first word and remove from list.
        short_text += ' %s' % words.pop(0)
    if len(words) == 1 and len(words[0]) <= post_length:
        short_text += ' %s' % words[0]
    elif words:
        short_text += ending
    return short_text[:limit]


def quitar_multiples_espacios_consecutivos(string):
    return ' '.join(s for s in string.split(' ') if s)

def reemplazar_caracteres_latinos(string):
    string = string.replace(u'á', 'a')
    string = string.replace(u'é', 'e')
    string = string.replace(u'í', 'i')
    string = string.replace(u'ó', 'o')
    string = string.replace(u'ú', 'u')
    string = string.replace(u'ü', 'u')
    string = string.replace(u'ñ', 'n')
    string = string.replace(u'Á', 'A')
    string = string.replace(u'É', 'E')
    string = string.replace(u'Í', 'I')
    string = string.replace(u'Ó', 'O')
    string = string.replace(u'Ú', 'U')
    string = string.replace(u'Ü', 'U')
    string = string.replace(u'Ñ', 'N')
    return string
