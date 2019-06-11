#!/usr/bin/env python3
# -*- encoding: utf-8 -*-

'''
© 2018, Florian Mortgat

Représentation d’entiers en numération bijective à 26 lettres, en particulier
pour la numérotation des noms des colonnes dans Excel ("A" = 1, "B" = 2, etc.
jusqu’à "XFD" = 16384.

Fonctionne aussi avec d’autres numérations bijectives, mais ça sert moins
souvent.

Utilisation ultra-basique :
>>> from bijnum import AZ
>>> AZ.aaa2n('AB')
28
>>> AZ.n2aaa(28)
'AB'

'''

class Bij:
    '''
    Class that helps you convert numbers into their Excel column name
    representation.  Zero is represented as an empty string.

    It can have other applications using other sets of symbols than the
    alphabet.

    Examples:
        1 = A
        2 = B
        ...
        26 = Z
        27 = AA
        28 = AB
        ...
        702 = ZZ
        703 = AAA
    '''
    def __init__(self, letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'):
        self.LETTERS = letters
        self.LETTER_VALUES = { letter: 1 + letters.find(letter) for letter in letters }
        self.BASE = len(letters)
    def aaa2n(self, aaa):
        n = 0
        positions = len(aaa) - 1
        for letter in aaa:
            letter_value = self.LETTER_VALUES[letter]
            n += letter_value * self.BASE ** positions
            positions -= 1
        return n
    def n2aaa(self, n):
        aaa = ''
        length = self.length_of_aaa_for_n(n)
        #print n, length
        remainder = n
        for i in range(length):
            p = self.BASE ** (length - 1)
            reserved = self.lowest_for_length(length - 1)
            q = (remainder - reserved) // p
            remainder = remainder - p * q
            aaa += self.LETTERS[q-1]
            length -= 1
        return aaa
    def lowest_for_length(self, length):
        if not length: return 0
        return sum((self.BASE**i for i in range(length)))
    def highest_for_length(self, length):
        return self.lowest_for_length(length + 1) - 1
    def length_of_aaa_for_n(self, n):
        i = 0
        while n >= self.lowest_for_length(i + 1):
            i += 1
        return i
        divider = 0
        positions = 0
        while True:
            if positions < 2:
                divider = self.BASE**positions
            else:
                divider += self.BASE**positions
            q = n // divider
            remainder = n % divider
            if n < divider or (remainder == 0 and q < self.BASE and positions):
                return positions
            positions += 1
    def enumerate(self, iterable):
        n = 0
        for item in iterable:
            yield (n, self.n2aaa(n+1), item)
            n += 1

    def check_reversible(self, n):
        try:
            return n == self.aaa2n(self.n2aaa(n))
        except:
            print(n)
            return False

AZ = Bij('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
