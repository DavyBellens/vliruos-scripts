with open('tool-link.txt', 'r+') as f:
    x = [i.removesuffix('\n')+'/modules/4696\n' if i!="\n" else i for i in f.readlines() ]
    f.seek(0)
    f.writelines(x)