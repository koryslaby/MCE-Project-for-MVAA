import sqlite3

con = sqlite3.connect('mce.sqlite3')
cur = con.cursor()

cur.execute('delete from institution;')
cur.execute('delete from course;')
cur.execute('delete from outcome;')
cur.execute('delete from reviewer;')

con.commit()

con.close()

con = sqlite3.connect('mce.sqlite3')
cur = con.cursor()

with open("institution.csv") as f:
    lines = f.readlines()
    for line in lines[1:]:
        inst1 = line.strip().split('::')
        # print(inst1)
        sql = "insert into Institution (InstitutionName, InstitutionAddress, InstitutionCity, InstitutionState, InstitutionZipcode, InstitutionWebsite) values (?, ?, ?, ?, ?, ?);"
        cur.execute(sql, (inst1[0], inst1[1], inst1[2], inst1[3], inst1[4], inst1[5]))		
        
with open("course.csv") as f:
    lines = f.readlines()
    for line in lines[1:]:
        inst2 = line.strip().split('::')
        print(inst2)
        sql = "insert into Course (CourseNumber, CourseName, CourseDescription, CourseCredit, CourseEquivalenceNonOC, InstitutionID) values (?, ?, ?, ?, ?, ?);"
        cur.execute(sql, (inst2[0], inst2[1], inst2[2], inst2[3], inst2[4], inst2[5]))

with open("outcome.csv") as f:
    lines = f.readlines()
    for line in lines[1:]:
        inst3 = line.strip().split('::')
        # print(inst3)
        sql = "insert into Outcome (CourseNumber, OutcomeDescription) values (?, ?);"
        cur.execute(sql, (inst3[0], inst3[1]))

with open("reviewer.csv") as f:
    lines = f.readlines()
    for line in lines[1:]:
        inst4 = line.strip().split('::')
        # print(inst3)
        sql = "insert into Reviewer (ReviewerName, ReviewerPhone, ReviewerEmail) values (?, ?, ?);"
        cur.execute(sql, (inst4[0], inst4[1], inst4[2]))
	
con.commit()

con.close()
