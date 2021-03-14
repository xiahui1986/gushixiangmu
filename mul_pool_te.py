from multiprocessing import Process,Pool
import multiprocessing as mul
import os, time, random


def job(x):
    return x*x

def long_time_task(name):
    name=str(name)
    print(name)
    print('Run task %s (%s)...' % (name, os.getpid()))
    start = time.time()
    time.sleep(random.random() * 3)
    end = time.time()
    print('Task %s runs %0.2f seconds.' % (name, (end - start)))

if __name__=='__main__':
    mul.freeze_support()
    #print('Parent process %s.' % os.getpid())
    #p=[]
    pool=mul.Pool(5)
    res=pool.map(job,range(5))
    print(res)
    #for i in range(100):
    #    x.apply_async(long_time_task,args=(i,))

