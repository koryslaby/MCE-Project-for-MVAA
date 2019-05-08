with open('JSToutcomes.csv', 'r') as f:
    lines = f.readlines()
    s = set()
    for line in lines:
        s.add(line.strip().split("::")[0])
    with open('JST.csv', 'w') as g:
        for item in sorted(s):
            g.write(item + "\n")
