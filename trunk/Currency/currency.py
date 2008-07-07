from decimal import Decimal

__all__ = [
    'Pound',
    'Dollar',
    'Euro',
    'Yen'
]


class Currency(Decimal):
    _symbol = None
    
    def __init__(self, *args, **kwargs):
        if self._symbol is None:
            raise TypeError("Can't instantiate an abstract Currency")
        super(Currency, self).__init__(*args, **kwargs)
        
    
    def __str__(self):
        return '%s%s' % (self._symbol, Decimal.__str__(self))
        
    
    def __repr__(self):
        return '<%s %s>' % (self.__class__.__name__, str(self))

    
    def __eq__(self, other):
        return type(self) == type(other) and Decimal.__eq__(self, other)
    
    
    def __ne__(self, other):
        return type(self) != type(other) or Decimal.__ne__(self, other)
    
    
    def __add__(self, other, context=None):
        return type(self) == type(other) and 
            
        

def _GetUnitClass(name, symbol):
    return type(name, (Currency,), {'_symbol': symbol})


Pound = _GetUnitClass('Pound', '\xa3')
Dollar = _GetUnitClass('Dollar', '$')
Euro = _GetUnitClass('Euro', u'\u20ac')
Yen = _GetUnitClass('Yen', '\xa5')

