with open('outcome.csv', 'r') as f:
    lines = f.readlines()
    s = set()
    for line in lines:
        s.add(line)
    with open('clean.csv', 'w') as g:
        for item in sorted(s):
            g.write(item)
