from random import shuffle as shuf

writers = ['Cade', 'Adam', 'Will', 'Chris', 'Tom', 'Duncan', 'Emi', 'Zack']
late_writers = ['John']

shuf(writers)
shuf(late_writers)

order = writers + late_writers

print(order)

### Official ordering
#
# ['Will', 'Adam', 'Chris', 'Zack', 'Cade', 'Emi', 'Duncan', 'John']
#
###