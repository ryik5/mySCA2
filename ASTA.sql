CREATE TABLE IF NOT EXISTS 'ConfigDB' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, ParameterName TEXT, Value TEXT, Description TEXT, DateCreated TEXT, IsPassword TEXT, IsExample TEXT, 
UNIQUE ('ParameterName', 'IsExample') ON CONFLICT REPLACE);

CREATE TABLE IF NOT EXISTS 'PeopleGroupDescription' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, GroupPerson TEXT, GroupPersonDescription TEXT, AmountStaffInDepartment TEXT, Recipient TEXT, 
UNIQUE ('GroupPerson') ON CONFLICT REPLACE);

CREATE TABLE IF NOT EXISTS 'PeopleGroup' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, FIO TEXT, NAV TEXT, GroupPerson TEXT, ControllingHHMM TEXT, ControllingOUTHHMM TEXT, 
Shift TEXT, Comment TEXT, Department TEXT, PositionInDepartment TEXT, DepartmentId TEXT, City TEXT, Boss TEXT, 
UNIQUE ('FIO', 'NAV', 'GroupPerson', 'DepartmentId') ON CONFLICT REPLACE);

CREATE TABLE IF NOT EXISTS 'ListOfWorkTimeShifts' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, NAV TEXT, DayStartShift TEXT, MoStart REAL,MoEnd REAL, TuStart REAL,TuEnd REAL, WeStart REAL,WeEnd REAL, ThStart REAL,ThEnd REAL, FrStart REAL,FrEnd REAL, SaStart REAL,SaEnd REAL, SuStart REAL,SuEnd REAL, Status Text, Comment TEXT, DayInputed TEXT, 
UNIQUE ('NAV', 'DayStartShift') ON CONFLICT REPLACE);

CREATE TABLE IF NOT EXISTS 'TechnicalInfo' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, PCName TEXT, POName TEXT, POVersion TEXT, LastDateStarted TEXT, CurrentUser TEXT, FreeRam TEXT, GuidAppication TEXT);

CREATE TABLE IF NOT EXISTS 'BoldedDates' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, DayBolded TEXT, NAV TEXT, DayType TEXT, DayDescription TEXT, DateCreated TEXT);

CREATE TABLE IF NOT EXISTS 'LastTakenPeopleComboList' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, ComboList TEXT, DateCreated TEXT, 
UNIQUE ('ComboList') ON CONFLICT REPLACE);

CREATE TABLE IF NOT EXISTS 'Mailing' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, RecipientEmail TEXT, GroupsReport TEXT, NameReport TEXT, Description TEXT, Period TEXT, Status TEXT, SendingLastDate TEXT, TypeReport TEXT, DayReport TEXT, DateCreated TEXT, 
UNIQUE ('RecipientEmail', 'GroupsReport', 'NameReport', 'Description', 'Period', 'TypeReport', 'DayReport') ON CONFLICT REPLACE);

CREATE TABLE IF NOT EXISTS 'MailingException' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, RecipientEmail TEXT, NameReport TEXT, Description TEXT, DayReport TEXT, DateCreated TEXT, 
UNIQUE ('RecipientEmail', 'NameReport') ON CONFLICT REPLACE);

CREATE TABLE IF NOT EXISTS 'SelectedCityToLoadFromWeb' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, City TEXT, DateCreated TEXT, 
UNIQUE ('City') ON CONFLICT REPLACE);

