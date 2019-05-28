import sqlite3

con = sqlite3.connect('mce.sqlite3')
cur = con.cursor()

cur.execute('delete from institution;')
cur.execute('delete from course;')
cur.execute('delete from outcome;')

con.commit()

con.close()