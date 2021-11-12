def t():
    global l
    l = []

def i():
    print(l)
    l = ["t"]


t()
i()
print(l)