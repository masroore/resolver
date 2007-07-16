_cache = {}

def GetCache(cacheId):
    return _cache.setdefault(cacheId, {})

def ClearCacheEntry(cacheId):
    _cache[cacheId] = {}
    