
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
        klass = _GetUnitClass('fish', 'F')
        self.assertTrue(issubclass(klass, Currency))
        self.assertEquals(klass.__name__, 'fish')
        
        instance = klass(0)
        self.assertEquals(str(instance), 'F0')
        self.assertEquals(repr(instance), '<fish F0>')

        
    def testAddition(self):
        Fish = _GetUnitClass('fish', 'F')
        Eggs = _GetUnitClass('Eggs', 'F')
        
        self.fail('not finished')
        
if __name__ == '__main__':
    unittest.main()
    