# -*- coding: utf-8 -*-
from django.test import TestCase
from excelutils.string_utils import agregar_espacios_luego_de_cada_coma_y_cada_punto, get_short_text

class StringUtilsTest(TestCase):

    def test_agregar_espacios_en_blanco_a_string_con_comas(self):
        expected = u'Esto, debería, estar, separado, por, coma, y, espacios'
        result = agregar_espacios_luego_de_cada_coma_y_cada_punto(u'Esto,debería,estar,separado,por,coma,y,espacios')
        self.assertEqual(expected, result)


class GetShortTextTests(TestCase):

    def test_none_retuns_empty_text(self):
        # setup
        origin_text = None
        expected_text = ''

        # exercise
        target = get_short_text(origin_text)

        # verify
        self.assertEquals(expected_text, target)

    def test_empty_text_retuns_empty_text(self):
        # setup
        origin_text = ''
        expected_text = ''

        # exercise
        target = get_short_text(origin_text)

        # verify
        self.assertEquals(expected_text, target)

    def test_short_text_retuns_the_same_text(self):
        # setup
        origin_text = 'Hola mundo'
        expected_text = 'Hola mundo'

        # exercise
        target = get_short_text(origin_text)

        # verify
        self.assertEquals(expected_text, target)

    def test_long_text_retuns_the_short_text_with_pos(self):
        # setup
        origin_text = 'Hola mundo'
        expected_text = 'Hola ...'

        # exercise
        target = get_short_text(origin_text, 8)

        # verify
        self.assertEquals(expected_text, target)

    def test_long_text_retuns_the_shorter_text_if_has_to_cut_words(self):
        # setup
        origin_text = 'Creating test database for alias'
        expected_text = 'Creating ...'

        # exercise
        target = get_short_text(origin_text, 16)

        # verify
        self.assertEquals(expected_text, target)

    def test_if_first_word_is_too_long_will_cut_it_and_not_append_the_ending(self):
       # setup
       origin_text = 'Creating test database for alias'
       expected_text = 'C'

       # exercise
       target = get_short_text(origin_text, 1)

       # verify
       self.assertEquals(expected_text, target)

    def test_could_append_the_last_if_its_short_than_ending(self):
       # setup
       origin_text = 'Creating te'
       expected_text = 'Creating te'

       # exercise
       target = get_short_text(origin_text, 12)

       # verify
       self.assertEquals(expected_text, target)

    def test_long_text_retuns_the_shorter_text_with_spaces(self):
        # setup
        origin_text = 'Creating   test database for alias'
        expected_text = 'Creating   test database ...'

        # exercise
        target = get_short_text(origin_text, 28)

        # verify
        self.assertEquals(expected_text, target)

    def test_long_text_retuns_the_shorter_text_if_has_to_cut_words(self):
        # setup
        origin_text = 'Creating test database for alias'
        expected_text = 'Creating ...'

        # exercise
        target = get_short_text(origin_text, 16)

        # verify
        self.assertEquals(expected_text, target)
