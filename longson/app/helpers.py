import openpyxl
from models import Janus

janus = open('./liste/janus.txt',mode='rb')
for line in janus:
    row = str(line).split('|')
    janus = Janus()
