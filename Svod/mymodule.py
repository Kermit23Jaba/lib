def dell_slo(word):
    a = word.split()
    for i in a:
        if '№' in list(i):
            c = i.replace(',', '')
            return c

def dell_slo_o(a):
    b = a.split()
    c = list(b[0])
    d = c[:-1]
    g =''.join(d)
    return g
