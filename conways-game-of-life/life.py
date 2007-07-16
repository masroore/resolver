from random import random, seed

class Grid(object):

    def __init__(self, x, y, density):
        self.width = x
        self.height = y
        self.grid = set()
        self._populate(density)


    def _populate(self, density):
        self.grid.clear()
        for x, y in self.locations():
            if random() < density:
                self.grid.add((x, y))


    def nextGeneration(self):
        self.grid = set([(x , y) for x, y in self.locations() if self._locationLives(x, y)])


    def locations(self):
        for y in range(self.height):
            for x in range(self.width):
                yield (x, y)


    def liveLocations(self):
        for location in self.locations():
            if self._isAlive(*location):
                yield location


    def _isAlive(self, x, y):
        return (x, y) in self.grid


    def _locationLives(self, x, y):
        neighbourCount = self._countLiveNeighbours(x, y)
        return (
            neighbourCount == 3 or
            neighbourCount == 2 and self._isAlive(x, y)
        )


    def _countLiveNeighbours(self, x, y):
        return len([loc for loc in self._getAdjacentLocations(x, y) if self._isAlive(*loc)])


    def _getAdjacentLocations(self, x, y):
        locations = set()
        for dy in (-1, 0, 1):
            for dx in (-1, 0, 1):
                if not (dx == 0 and dy == 0):
                    lx = (x + dx + self.width) % self.width
                    ly = (y + dy + self.height) % self.height
                    locations.add((lx, ly))
        return locations


    def __str__(self):
        onChar = "O"
        offChar = "-"
        chars = []
        for x, y in self.locations():
            if self._isAlive(x, y):
                chars.append(onChar)
            else:
                chars.append(offChar)
            if x == self.width - 1:
                chars.append('\n')
        return ''.join(chars)



