%let pgm=utl-create-ms-access-mdb-table-from-sas-table-without-sas-access-using-powershell;

%stop_submission;

Create ms access mdb table from sas table without sas access using powershell

  CONTENTS

     1 download simple.mdb
       https://tinyurl.com/2azm47wd
     2 powershell example using utl_psbegin and utl_psend drop down

 you need tagsets.sql (https://tinyurl.com/y9nfugth)
 you need to install oledb (https://tinyurl.com/3em6k2z7)

github
https://tinyurl.com/y9v55j5z
https://github.com/rogerjdeangelis/utl-create-ms-access-mdb-table-from-sas-table-without-sas-access-using-powershell

macros
https://tinyurl.com/y9nfugth
https://github.com/rogerjdeangelis/utl-macros-used-in-many-of-rogerjdeangelis-repositories

github
https://tinyurl.com/5mpbc23d
https://github.com/rogerjdeangelis/utl-creating-sqlite-and-postgresql-tables-from-sas-datasets-without-sas-access-and-a-blueprint

related repo
github
https://tinyurl.com/3em6k2z7
https://github.com/rogerjdeangelis/utl-export-ms-access-table-from-mdb-database-to-sas-table-without-sas-access-using-powershell

download and install oledb for access
https://tinyurl.com/2azm47wd
https://www.microsoft.com/en-us/download/details.aspx?id=54920

download simple.mdb (download raw)
https://tinyurl.com/2azm47wd
https://github.com/rogerjdeangelis/utl-export-ms-access-table-from-mdb-database-to-sas-table-without-sas-access-using-powershell/blob/main/simple.mdb

/**************************************************************************************************************************/
/* INPUT          | PROCESS                                                             |     OUTPUT                      */
/* =====          | =======                                                             |     ======                      */
/* WORK,CLASS     | STEPS                                                               | MS ACCES MDB TABLE CLASS        */
/*                |                                                                     |                                 */
/*  NAME  SEX AGE | 1  use tagsets to get sas create table statements                   |  NAME  SEX AGE                  */
/*                | 2  edit the sas create script for ms access                         |                                 */
/* Alfred  M   14 | 3  powershell script w included powershell create scriptfrom 2      |  Alfred  M   14                 */
/* Alice   F   13 |                                                                     |  Alice   F   13                 */
/* Barbara F   13 | EDIT TAGSETS>SQL OUTPUT FOR MS ACCESS                               |  Barbara F   13                 */
/* Carol   F   14 |                                                                     |  Carol   F   14                 */
/* Henry   M   14 | %utl_tagsets_sql;                                                   |  Henry   M   14                 */
/* James   M   12 | options ls=256;                                                     |  James   M   12                 */
/*                | ods tagsets.sql file="c:/temp/sqlcreins.sql";                       |                                 */
/*                | proc print data=class;                                              |                                 */
/* data class;    | run;quit;                                                           |                                 */
/*   input        | ods _all_ close; ** very important;                                 |                                 */
/*     name$      | data _null_;                                                        |                                 */
/*     sex$ age;  |  infile "c:/temp/sqlcreins.sql" firstobs=2 end=dne;                 |                                 */
/* cards4;        |    file "c:/temp/sqlcreins.ps1";                                    |                                 */
/* Alfred  M 14   |  input;                                                             |                                 */
/* Alice   F 13   |  select;                                                            |                                 */
/* Barbara F 13   |   when (dne) do;                                                    |                                 */
/* Carol   F 14   |      _infile_  = quote(strip(_infile_));                            |                                 */
/* Henry   M 14   |      put _infile_ ;                                                 |                                 */
/* James   M 12   |      putlog _infile_;                                               |                                 */
/* ;;;;           |      _infile_=')';                                                  |                                 */
/* run;quit;      |   end;                                                              |                                 */
/*                |   when (index(_infile_,"Create")) do;                               |                                 */
/*                |      _infile_=cats('$createTableSql=',quote(strip(_infile_)));      |                                 */
/*                |      put _infile_;                                                  |                                 */
/*                |      putlog _infile_;                                               |                                 */
/*                |      _infile_='$connection.Execute($createTableSql)';               |                                 */
/*                |      put _infile_;                                                  |                                 */
/*                |      putlog _infile_;                                               |                                 */
/*                |      _infile_ = '$insertStatements = @(' ;                          |                                 */
/*                |   end;                                                              |                                 */
/*                |   otherwise  _infile_  = cats(quote(strip(_infile_)),',');          |                                 */
/*                |  end;                                                               |                                 */
/*                |  put _infile_;                                                      |                                 */
/*                |  putlog _infile_;                                                   |                                 */
/*                |  run;quit;                                                          |                                 */
/*                |                                                                     ----------------------------------*/
/*                |                                                                                                       */
/*                |                                                                                                       */
/*                |  /*---- THE CODE ABOVE CONVERTS THIS                                                                  */
/*                |                                                                                                       */
/*                |  c:/temp/sqlcreins.sql                                                                                */
/*                |                                                                                                       */
/*                |  proc sql;                                                                                            */
/*                |   Create table class(name varchar(200),sex varchar(200),age float);                                   */
/*                |  Insert into class(name, sex, age) Values('Alfred', 'M', 14);                                         */
/*                |  Insert into class(name, sex, age) Values('Alice', 'F', 13);                                          */
/*                |  Insert into class(name, sex, age) Values('Barbara', 'F', 13);                                        */
/*                |  Insert into class(name, sex, age) Values('Carol', 'F', 14);                                          */
/*                |  Insert into class(name, sex, age) Values('Henry', 'M', 14);                                          */
/*                |  Insert into class(name, sex, age) Values('James', 'M', 12);                                          */
/*                |                                                                                                       */
/*                |  ---- TO THIS postscript code c:/temp/sqlcreins.ps1                                                   */
/*                |                                                                                                       */
/*                |  $createTableSql ="Create table class(name varchar(200),sex varchar(200),age float);"                 */
/*                |  $connection.Execute($createTableSql)                                                                 */
/*                |  $insertStatements = @(                                                                               */
/*                |  "Insert into class(name, sex, age) Values('Alfred', 'M', 14);",                                      */
/*                |  "Insert into class(name, sex, age) Values('Alice', 'F', 13);",                                       */
/*                |  "Insert into class(name, sex, age) Values('Barbara', 'F', 13);",                                     */
/*                |  "Insert into class(name, sex, age) Values('Carol', 'F', 14);",                                       */
/*                |  "Insert into class(name, sex, age) Values('Henry', 'M', 14);",                                       */
/*                |  "Insert into class(name, sex, age) Values('James', 'M', 12);"                                        */
/*                |  )                                                                                                    */
/*                |  ----*/                                                                                               */
/*                |                                                                                                       */
/*                |                                                                                                       */
/*                |  POWERSHELL OLEDB CREATE ACCESS MDB TABLE                                                             */
/*                |  ========================================                                                             */
/*                |  %utl_psbegin;                                                                                        */
/*                |  parmcards4;                                                                                          */
/*                |  # Set file paths                                                                                     */
/*                |  $csvPath = "D:\csv\have.csv"                                                                         */
/*                |  $accessDbPath = "D:\mdb\simple.mdb"                                                                  */
/*                |  $tableName = "class"                                                                                 */
/*                |                                                                                                       */
/*                |  # Connect to Access database                                                                         */
/*                |  $connection = New-Object -ComObject ADODB.Connection                                                 */
/*                |  $connection.ConnectionString="Provider=Microsoft.ACE.OLEDB.16.0;Data Source=$accessDbPath"           */
/*                |  $connection.Open()                                                                                   */
/*                |                                                                                                       */
/*                |  # Drop table if it exists                                                                            */
/*                |  try {                                                                                                */
/*                |      $connection.Execute("DROP TABLE $tableName")                                                     */
/*                |  } catch {}                                                                                           */
/*                |                                                                                                       */
/*                |  # Create new table                                                                                   */
/*                |  . c:\temp\sqlcreins.ps1                                                                              */
/*                |                                                                                                       */
/*                |  foreach ($sql in $insertStatements) {                                                                */
/*                |      $connection.Execute($sql)                                                                        */
/*                |  }                                                                                                    */
/*                |                                                                                                       */
/*                |  # Close connection                                                                                   */
/*                |  $connection.Close()                                                                                  */
/*                |  $connection = $null                                                                                  */
/*                |  ;;;;                                                                                                 */
/*                |  %utl_psend;                                                                                          */
/**************************************************************************************************************************/

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
*/

data class;
  input
    name$
    sex$ age;
cards4;
Alfred  M 14
Alice   F 13
Barbara F 13
Carol   F 14
Henry   M 14
James   M 12
;;;;
run;quit;

/**************************************************************************************************************************/
/* WORK,CLASS                                                                                                             */
/*                                                                                                                        */
/*  NAME  SEX AGE                                                                                                         */
/*                                                                                                                        */
/* Alfred  M   14                                                                                                         */
/* Alice   F   13                                                                                                         */
/* Barbara F   13                                                                                                         */
/* Carol   F   14                                                                                                         */
/* Henry   M   14                                                                                                         */
/* James   M   12                                                                                                         */
/**************************************************************************************************************************/

/*
 _ __  _ __ ___   ___ ___  ___ ___
| `_ \| `__/ _ \ / __/ _ \/ __/ __|
| |_) | | | (_) | (_|  __/\__ \__ \
| .__/|_|  \___/ \___\___||___/___/
|_|
*/

PROCESS
=======
STEPS

1  use tagsets to get sas create table statements
2  edit the sas create script for ms access
3 powershell script w included powershell create script

EDIT TAGSETS>SQL OUTPUT FOR MS ACCESS
=====================================

%utl_tagsets_sql;
options ls=256;
ods tagsets.sql file="c:/temp/sqlcreins.sql";
proc print data=class;
run;quit;
ods _all_ close; ** very important;
data _null_;
infile "c:/temp/sqlcreins.sql" firstobs=2 end=dne;
file "c:/temp/sqlcreins.ps1";
input;
select;
when (dne) do;
_infile_  = quote(strip(_infile_));
put _infile_ ;
putlog _infile_;
_infile_=')';
end;
when (index(_infile_,"Create")) do;
_infile_=cats('$createTableSql=',quote(strip(_infile_)));
put _infile_;
putlog _infile_;
_infile_='$connection.Execute($createTableSql)';
put _infile_;
putlog _infile_;
_infile_ = '$insertStatements = @(' ;
end;
otherwise  _infile_  = cats(quote(strip(_infile_)),',');
end;
put _infile_;
putlog _infile_;
run;quit;

/*---- THE CODE ABOVE CONVERTS THIS

c:/temp/sqlcreins.sql

proc sql;
Create table class(name varchar(200),sex varchar(200),age float);
Insert into class(name, sex, age) Values('Alfred', 'M', 14);
Insert into class(name, sex, age) Values('Alice', 'F', 13);
Insert into class(name, sex, age) Values('Barbara', 'F', 13);
Insert into class(name, sex, age) Values('Carol', 'F', 14);
Insert into class(name, sex, age) Values('Henry', 'M', 14);
Insert into class(name, sex, age) Values('James', 'M', 12);

---- TO THIS postscript code c:/temp/sqlcreins.ps1

$createTableSql ="Create table class(name varchar(200),sex varchar(200),age float);"
$connection.Execute($createTableSql)
$insertStatements = @(
"Insert into class(name, sex, age) Values('Alfred', 'M', 14);",
"Insert into class(name, sex, age) Values('Alice', 'F', 13);",
"Insert into class(name, sex, age) Values('Barbara', 'F', 13);",
"Insert into class(name, sex, age) Values('Carol', 'F', 14);",
"Insert into class(name, sex, age) Values('Henry', 'M', 14);",
"Insert into class(name, sex, age) Values('James', 'M', 12);"
)
----*/


POWERSHELL OLEDB CREATE ACCESS MDB TABLE
========================================

%utl_psbegin;
parmcards4;
# Set file paths
$csvPath = "D:\csv\have.csv"
$accessDbPath = "D:\mdb\simple.mdb"
$tableName = "class"

# Connect to Access database
$connection = New-Object -ComObject ADODB.Connection
$connection.ConnectionString="Provider=Microsoft.ACE.OLEDB.16.0;Data Source=$accessDbPath"
$connection.Open()

# Drop table if it exists
try {
$connection.Execute("DROP TABLE $tableName")
} catch {}

# Create new table
. c:\temp\sqlcreins.ps1

foreach ($sql in $insertStatements) {
$connection.Execute($sql)
}

# Close connection
$connection.Close()
$connection = $null

;;;;
%utl_psend;

/*           _               _
  ___  _   _| |_ _ __  _   _| |_
 / _ \| | | | __| `_ \| | | | __|
| (_) | |_| | |_| |_) | |_| | |_
 \___/ \__,_|\__| .__/ \__,_|\__|
                |_|
*/

d;/mdb/simple.mdb

Table class

  NAME  SEX AGE

 Alfred  M   14
 Alice   F   13
 Barbara F   13
 Carol   F   14
 Henry   M   14
 James   M   12

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
