BEGIN TRANSACTION;
CREATE TABLE IF NOT EXISTS `Outcome` (
	`OutcomeID`	INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,
	`CourseNumber`	INTEGER NOT NULL,
	`OutcomeDescription`	TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS `Institution` (
	`InstitutionID`	INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,
	`InstitutionName`	TEXT NOT NULL UNIQUE,
	`InstitutionAddress`	TEXT NOT NULL,
	`InstitutionCity`	TEXT NOT NULL,
	`InstitutionState`	TEXT NOT NULL,
	`InstitutionZipcode`	TEXT NOT NULL,
	`InstitutionWebSite`	TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS `Course` (
	`CourseID`	INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,
	`CourseNumber`	TEXT NOT NULL,
	`CourseName`	TEXT,
	`CourseDescription`	TEXT,
	`CourseCredit`	REAL,
	`CourseEquivalenceNonOC`	TEXT,
	`InstitutionID`	INTEGER
);
COMMIT;
