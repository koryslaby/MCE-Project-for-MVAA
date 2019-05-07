import sqlite3

con = sqlite3.connect('mce.sqlite3')
cur = con.cursor()

#data = []
#with open("outcomedata.csv") as f:
#	lines = f.readlines()
#	for line in lines:
#		indata = line.strip().split(':')
#		print(indata[0])

with open("institution.csv") as f:
    lines = f.readlines()
    for line in lines[1:]:
        inst1 = line.strip().split('::')
        sql = "insert into Institution (InstitutionName, InstitutionAddress, InstitutionCity, InstitutionState, InstitutionZipcode, InstitutionWebsite) values (?, ?, ?, ?, ?, ?);"
        cur.execute(sql, (inst1[0], inst1[1], inst1[2], inst1[3], inst1[4], inst1[5]))		
        
with open("course.csv") as f:
    lines = f.readlines()
    for line in lines[1:]:
        inst2 = line.strip().split('::')
        sql = "insert into Course (CourseNumber, CourseName, CourseDescription, CourseCredit, CourseEquivalenceNonOC, InstitutionID) values (?, ?, ?, ?, ?, ?);"
        cur.execute(sql, (inst2[0], inst2[1], inst2[2], inst2[3], inst2[4], inst2[5]))

with open("outcome.csv") as f:
    lines = f.readlines()
    for line in lines[1:]:
        inst3 = line.strip().split('::')
        sql = "insert into Outcome (OutcomeDescription, CourseNumber) values (?, ?);"
        cur.execute(sql, (inst3[1], inst3[0]))

con.commit()

con.close()