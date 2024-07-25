with open('sp-links.txt', 'r+') as f:
    x = ["https://vliruos.sharepoint.com/sites/PROGRAMMAWERKING/"+i.replace(' ', '%20') if i.startswith('/') else i for i in f.readlines() ]
    f.seek(0)
    f.writelines(x)