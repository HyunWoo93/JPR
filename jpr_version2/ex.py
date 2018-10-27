import re


p = re.compile('_.+')
a= ['_1', '_2', '1_1', '2_2']
for file in a:
	if p.match(file):
		print(file)
	else:
		print(file + 'nono')