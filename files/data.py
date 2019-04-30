data = []
with open("outcomedata.csv") as f:
	lines = f.readlines()
	for line in lines:
		indata = line.strip().split(':')
		print(indata)
		data.append(indata)


