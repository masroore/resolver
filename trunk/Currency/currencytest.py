
import sys
sys.setrecursionlimit(100)

import unittest
from decimal import Decimal

from currency import (
    Currency, Pound,
    Yen, Euro,
    Dollar, _GetUnitClass
)


class CurrencyTest(unittest.TestCase):
    
    
    def testCurrency(self):
        self.assertTrue(issubclass(Currency, Decimal))
        self.assertEquals(Currency._symbol, None)
        self.assertRaises(TypeError, Currency)

        
    def testGetUnitClass(self):
        klass = _GetUnitClass('Fish', 'F')
        self.assertTrue(issubclass(klass, Currency))
        self.assertEquals(klass.__name__, 'Fish')
        
        instance = klass(0)
        self.assertEquals(str(instance), 'F0')
        self.assertEquals(repr(instance), '<Fish F0>')


    def testEquality(self):
        Fish = _GetUnitClass('Fish', 'F')
        Eggs = _GetUnitClass('Eggs', 'F')
        f = Fish(0)
        
        self.assertEquals(f, Fish(0))
        
        self.assertNotEquals(f, Fish(1))
        self.assertNotEquals(f, Decimal(0))
        self.assertNotEquals(f, Eggs(0))
        self.assertNotEquals(f, None)
        self.assertNotEquals(f, 0)
        self.assertNotEquals(f, 0.0)
        self.assertNotEquals(f, long(0))
        
        self.assertTrue(f != Decimal(0))

        
    def testAddition(self):
        Fish = _GetUnitClass('Fish', 'F')
        Eggs = _GetUnitClass('Eggs', 'F')
        
        f = Fish(1)
        e = Eggs(1)
        self.assertEquals(f + f, Fish(2))
        self.assertRaises(TypeError, lambda: f + e)
        self.assertRaises(TypeError, lambda: f + 1)
        self.assertRaises(TypeError, lambda: f + 1.0)
        self.assertRaises(TypeError, lambda: f + Decimal(1))
        
        self.assertRaises(TypeError, lambda: 1 + f)
        self.assertRaises(TypeError, lambda: 1.0 + f)
        self.assertRaises(TypeError, lambda: Decimal(1) + f)
        
        
        
if __name__ == '__main__':
    unittest.main()
    