from numba import jit
from datetime import datetime as dt
@jit
def xxx():
    for i in range(100):
        for j in range(100):
            for m in range(100):
                for n in range(100):
                    a=i*j*m*n

if __name__ == "__main__":
    print(dt.now())
    xxx()
    print(dt.now())