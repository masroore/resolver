from decimal import Decimal
# from System import Decimal

__all__ = [
    'Pound',
    'Dollar',
    'Euro',
    'Yen',
    'GetUnitClass'
]


def Matching(method):
    def inner(self, other, *args, **keywargs):
        if not type(other) == type(self):
            raise TypeError("operation not allowed between %r and %r" % (self, other))
        result = method(self, other, *args, **keywargs)
        return self.__class__(result)
    
    inner.__name__ = method.__name__
    return inner
                        

class Currency(object):
    _symbol = None
    
    def __new__(cls, value):
        self = object.__new__(cls)
        if self._symbol is None:
            raise TypeError("Can't instantiate an abstract Currency")
        if isinstance(value, cls):
            value = value._value
        self._value = Decimal(value)
        return self
        
    
    def __str__(self):
        return '%s%s' % (self._symbol, str(self._value))
        
    
    def __repr__(self):
        return '<%s %s>' % (self.__class__.__name__, str(self))

    
    def __eq__(self, other):
        return type(self) == type(other) and self._value == other._value
    
    
    def __ne__(self, other):
        return not self.__eq__(other)
    
    @Matching
    def __add__(self, other):
        return self._value + other._value
            
    __radd__ = __add__
    
    
    @Matching
    def __sub__(self, other):
        return self._value - other._value
    
    __rsub__ = __sub__
    
    
    def __mul__(self, other):
        if isinstance(other, Currency):
            raise TypeError("Can't multiply %r by %r" % (self, other))
        return self.__class__(self._value * other)
    
    __rmul__ = __mul__

    
    def __div__(self, other):
        if isinstance(other, Currency):
            raise TypeError("Can't divide %r by %r" % (self, other))
        return self.__class__(self._value / other)
    

    def __floordiv__(self, other):
        if isinstance(other, Currency):
            raise TypeError("Can't floor divide %r by %r" % (self, other))
        return self.__class__(self._value // other)
    
    
    def __nonzero__(self):
        return bool(self._value)
    
    
def GetUnitClass(name, symbol):
    return type(name, (Currency,), {'_symbol': symbol})


Pound = GetUnitClass('Pound', '\xa3')
Dollar = GetUnitClass('Dollar', '$')
Euro = GetUnitClass('Euro', u'\u20ac')
Yen = GetUnitClass('Yen', '\xa5')

