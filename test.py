import os

print(os.getcwd())
for roots, dirs, files in os.walk(os.getcwd() + r"\Data\POPData"):
    for file in files:
        print(file)
"""        if file[:1] != '~' and file[-5:] == '.xlsx':
            if ((file[:3] == 'JGP') | (file[:3] == 'ALL')):
                name = file[:3]"""
