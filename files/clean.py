with open('files/outcome.csv', 'r') as f:
    lines = f.readlines()
    s = set()
    for line in lines:
        s.add(line.strip())
    with open('outcome.csv', 'w') as g:
        for item in s:
            g.write(item + "\n")