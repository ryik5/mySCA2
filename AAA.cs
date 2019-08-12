using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA
{
    class AAA
    {
        //todo
        //will make Update local DB structure

        /*
-= -= update scheme =-
1    Start a transaction.
2    Run PRAGMA schema_version to determine the current schema version number. This number will be needed for step 6 below.
3    Activate schema editing using PRAGMA writable_schema=ON.
4    Run an UPDATE statement to change the definition of table X in the sqlite_master table: 
	UPDATE sqlite_master SET sql=... WHERE type='table' AND name='X';
5    Caution: Making a change to the sqlite_master table like this will render the database corrupt and unreadable 
	if the change contains a syntax error. It is suggested that careful testing of 
	the UPDATE statement be done on a separate blank database prior to using it on a database containing important data.
6    If the change to table X also affects other tables or indexes or triggers are views within schema, then 
	run UPDATE statements to modify those other tables indexes and views too. For example, 
	if the name of a column changes, all FOREIGN KEY constraints, triggers, indexes, and views 
	that refer to that column must be modified.
7    Caution: Once again, making changes to the sqlite_master table like 
	this will render the database corrupt and unreadable if the change contains an error. 
	Carefully test this entire procedure on a separate test database 
	prior to using it on a database containing important data and/or 
	make backup copies of important databases prior to running this procedure.
8    Increment the schema version number using PRAGMA schema_version=X 
	where X is one more than the old schema version number found in step 2 above.
9    Disable schema editing using PRAGMA writable_schema=OFF.
10    (Optional) Run PRAGMA integrity_check to verify that the schema changes did not damage the database.
11    Commit the transaction started on step 1 above. 

-= If some future version of SQLite adds new ALTER TABLE capabilities, 
	those capabilities will very likely use one of the two procedures outlined above.


==================
-= rename table =-

ALTER TABLE existing_table
RENAME TO new_table;


=======================
-= Adding a new column  //new column cannot have a UNIQUE or PRIMARY KEY constraint

ALTER TABLE table
ADD COLUMN column_definition;


======================


-=  ALTER TABLE – Other actions =-

PRAGMA foreign_keys=off; 
BEGIN TRANSACTION; 

ALTER TABLE table RENAME TO temp_table; 
CREATE TABLE table
( 
   column_definition,
   ...
);
 
INSERT INTO table (column_list)
  SELECT column_list
  FROM temp_table;
 
DROP TABLE temp_table;
 
COMMIT;
 
PRAGMA foreign_keys=on;
 

         
-= DROP COLUMN example =-

BEGIN TRANSACTION;
 
ALTER TABLE equipment RENAME TO temp_equipment;
 
CREATE TABLE equipment (
 name text NOT NULL,
 model text NOT NULL,
 serial integer NOT NULL UNIQUE
);
 
INSERT INTO equipment 
SELECT
 name, model, serial
FROM
 temp_equipment;
 
DROP TABLE temp_equipment;
 
COMMIT;

         
         
         
         
         */

    }
}
