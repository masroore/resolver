from sys import stdout
from random import seed
from time import sleep

from life import Grid

seed(0)
g = Grid(40, 30, 0.2)
while True:
    g.nextGeneration()
    print g
    stdout.flush()
    sleep(0.2)


