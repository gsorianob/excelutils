# encoding: utf-8

def assertIn(first, second, msg=None):
    """Test that first is in second.
    """
    try:
        assert first in second
    except AssertionError:
        if msg:
            raise AssertionError('Fail: %s' % msg)
        else:
            raise AssertionError('Fail: "%s" not found in "%s"' % (first, second))


def assertNotIn(first, second, msg=None):
    """Test that first is not in second.
    """
    try:
        assert first not in second
    except AssertionError:
        if msg:
            raise AssertionError('Fail: %s' % msg)
        else:
            raise AssertionError('Fail: "%s" found in "%s"' % (first, second))


def assertRaises(exception, callable, *args, **kwargs):
    """Test that an exception is raised when callable is called with any
    positional or keyword arguments that are also passed to assertRaises().
    The test passes if exception is raised, is an error if another exception
    is raised, or fails if no exception is raised. To catch any of a group
    of exceptions, a tuple containing the exception classes may be passed
    as exception.
    """
    try:
        callable(*args, **kwargs)
    except exception:
        return
    raise AssertionError('Fail: %s not raised' % (exception.__name__, ))


def assertIsInstance(obj, types):
    """Test that this object is instance of that type or one of types.
    """
    if not isinstance(obj, types):
        raise AssertionError('Fail: Object "%s", type "%s"  is not instance of %s' % (obj, type(obj), types))


def assertIsNotInstance(obj, types):
    """Test that this object is instance of that type or one of types.
    """
    if isinstance(obj, types):
        raise AssertionError('Fail: Object "%s", type "%s"  is instance of %s' % (obj, type(obj), types))


def assertIsNone(expr, msg=None):
    """Test that expr is None.
    """
    try:
        assert expr is None
    except AssertionError:
        if msg:
            raise AssertionError('Fail: %s' % msg)
        else:
            raise AssertionError('Fail: "%s" is not None' % (expr, ))


def assertIsNotNone(expr, msg=None):
    """Test that expr is not None.
    """
    try:
        assert expr is not None
    except AssertionError:
        if msg:
            raise AssertionError('Fail: %s' % msg)
        else:
            raise AssertionError('Fail: "%s" is None' % (expr, ))


def assertLength(iterable, quantity, msg=None):
    """Test that iterable has quantity elements
    """
    length = len(iterable)
    try:
        assert length == quantity
    except AssertionError:
        if msg:
            raise AssertionError('Fail: %s' % msg)
        else:
            raise AssertionError('Fail: "%s" has "%s" elements, not %s' % (iterable, length, quantity))


def assertEqualsLists(self, lista_actual, lista_esperada, tipo_elementos=''):
    lista_esperada = sorted(lista_esperada, key=lambda x: x.id)
    lista_actual = sorted(lista_actual, key=lambda x: x.id)
    try:
        self.assertEquals(lista_esperada, lista_actual)
    except AssertionError:
        interseccion = [r for r in lista_esperada if r in lista_actual]
        esperados = [r for r in lista_esperada if r not in lista_actual]
        no_esperados = [r for r in lista_actual if r not in lista_esperada]
        msg = '\n%s:\n  Esperados y no encontrados: "%s"\n  Encontrados y no esperados:"%s"\n  Esperados y encontrados:"%s"'
        raise AssertionError(msg % (tipo_elementos, esperados, no_esperados, interseccion))

def assertRegexpMatches(text, regexp):
    """
    Test that a regexp search matches text. In case of failure, the error
    message will include the pattern and the text (or the pattern and the
    part of text that unexpectedly matched). regexp may be a regular expression
     object or a string containing a regular expression suitable for use by
     re.search().
    """
    import re
    result = re.search(regexp, text)
    try:
        assert result is not None
    except AssertionError:
        raise AssertionError('Fail: Not found "%s" in "%s"' % (regexp, text))
