a = [['1', 'a'], ['2', 'b'], ['3', 'c']]

output = ''
for i in range(len(a)):
    output += (f'\n{a[i][0]}: {a[i][1]}')

print(output)
