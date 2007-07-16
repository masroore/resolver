from life import Grid

grid = None

def getGrid(width, height, density):
    global grid
    if grid is None:
        grid = Grid(width, height, density)
    return grid