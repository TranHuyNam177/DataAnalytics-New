def __convertInteger__(x:int):
    x = str(x)
    if len(x)==1:
        x = '0'+x
    return x