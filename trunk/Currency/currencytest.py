
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

        
    def testAddition(self):
        Fish = _GetUnitClass('Fish', 'F')
        
        self.fail('not finished')
        
        
        
if __name__ == '__main__':
    unittest.main()
    