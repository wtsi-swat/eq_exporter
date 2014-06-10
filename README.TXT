Instructions for building the "EQ Exporter" program, for exporting data from the "EQ" program to either a MySQL database or Microsoft Excel (.xlsx) file.

This program has been designed for computers/tablets running versions of the Microsoft Windows operating system. It is known to work with Windows XP, Windows 7 and Windows 8. The .Net framework from Microsoft
must be installed on a computer before using EQ. See this link for more information: http://www.microsoft.com/en-gb/download/details.aspx?id=30653

1. Install Microsoft Visual Studio for Windows Desktop (denoted hereafter as "VS". Note: the free "Express" version works as of this writing, i.e. the 2013 version)
2. Open VS and choose File -> Open Project
3. Open file: Projects -> EQ Exporter -> EQ Exporter.sln
4. Click "Start" in the main menu and VS should compile and run the project.

To use EQ Exporter:

Exporting data from EQ is a 2-stage process. The first part is performed using the EQ program itself. The second part is done using this program.

1. In "EQ", on the "Existing Questionnaire" tab, choose a form-type from the list under "export data", then click the "export selected form" button, which will copy all the data for that form to the directory of your choice. 

2. In "EQ Exporter", select either "Excel file" or "MySQL Database", then click "Export Data Now".

3. For "Excel file", select the directory that you saved your data in 1 (above), then the data will be saved as the file EQ-data.xlsx
The data is organised such that all data for each participant is placed on a single row.

4. For "MySQL Database", first set up your database with a table defined as follows:

create table eq_data(

    id int unsigned auto_increment,
	questionnaire_name varchar(500),
	participant_id varchar(500),
	question_code varchar(500),
	answer text,
    primary key(id)
    
) engine=InnoDB;

The data is organised differently from the Excel file, i.e. each row only contains data for a single field.

In EQ Exporter, you will be asked for your MySQL connection details. Enter these, then click "OK". Your data will then be loaded into MySQL.


Notes: To use EQ Exporter, you will also need to have the following files in the same directory as "EQ Exporter.exe"
MySql.Data.dll
EPPlus.dll
EPPlus.xml




