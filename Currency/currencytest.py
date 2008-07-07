
import sys
sys.setrecursionlimit(100)

import unittest
from decimal import Decimal

from currency import (
    Currency, Pound,
    Yen, Euro, Matching, 
    Dollar, _GetUnitClass
)


class CurrencyTest(unittest.TestCase):
    
    
    def testCurrency(self):
        self.assertTrue(issubclass(Currency, Decimal))
        self.assertEquals(Currency._symbol, None)
        self.assertRaises(TypeError, Currency)

        
    def testGetUnitClass(self):
        Fish = _GetUnitClass('Fish', 'F')
        self.assertTrue(issubclass(Fish, Currency))
        self.assertEquals(Fish.__name__, 'Fish')
        
        instance = Fish(0)
        self.assertEquals(str(instance), 'F0')
        self.assertEquals(repr(instance), '<Fish F0>')
        
    
    def testMatchingDecorator(self):
        class A(object):
            def __init__(self, value):
                self.value = value
        
        @Matching
        def test(self, other):
            return 3
        
        self.assertEquals(test.__name__, 'test')
        self.assertRaises(TypeError, test, 1, None)
        
        result = test(A(1), A(2))
        self.assertTrue(isinstance(result, A))
        self.assertEquals(result.value, 3)
        
        
    def testConstruction(self):
        Fish = _GetUnitClass('Fish', 'F')
        
        f = Fish(0)
        self.assertEquals(f, Fish(f))
        self.assertEquals(f, Fish(Decimal(0)))
        

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
    