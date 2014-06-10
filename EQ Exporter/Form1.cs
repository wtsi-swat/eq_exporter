
/*
Copyright (c) 2014 Genome Research Ltd.
Author: Stephen Rice <sr7@sanger.ac.uk>
This file is part of EQ Exporter.
EQ-Exporter is free software: you can redistribute it and/or modify it under
the terms of the GNU General Public License as published by the Free Software
Foundation; either version 3 of the License, or (at your option) any later
version.
This program is distributed in the hope that it will be useful, but WITHOUT
ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
details.
You should have received a copy of the GNU General Public License along with
this program. If not, see <http://www.gnu.org/licenses/>.
*/

﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;



namespace EQ_Exporter
{
    public partial class Form1 : Form
    {

        public string globalSurveyName { get; set; }
        public string dbName { get; set; }
        public string dbUserName { get; set; }
        public string dbHostName { get; set; }
        public string dbPort { get; set; }
        public string dbPassword { get; set; }

        public bool dbConnectionParamsSet { get; set; }

        
        
        
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            //user clicked button to build excel file

            //open a dir dialog to get the source dir

            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = "Select the folder where your EQ data files were copied to.";

            string dataDir;
            string filename;
            string surveyName;
            string partID;

            Match match;

            //list of all results
            List<EQresult> resultList = new List<EQresult>();



            // Show the FolderBrowserDialog.
            DialogResult result = folderBrowserDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                dataDir = folderBrowserDialog.SelectedPath;

                //get all the files in this dir
                DirectoryInfo di = new DirectoryInfo(dataDir);


                foreach (FileInfo file in di.GetFiles())
                {

                    //filename
                    filename = file.Name;

                    //is this a valid datafile format?
                    match = Regex.Match(filename, @"^participant_data_(.+)_(.+)\.txt$");

                    if (match.Success)
                    {

                        partID = match.Groups[1].Value;
                        surveyName = match.Groups[2].Value;


                    }
                    else
                    {
                        //the older format might be used, which does not contain the survey-id

                        match = Regex.Match(filename, @"^final_data_(.+)\.txt$");

                        if (match.Success)
                        {
                            partID = match.Groups[1].Value;

                            //we need to ask the user for the name of the survey
                            //have we aksed this already?

                            if (globalSurveyName == null)
                            {

                                Form2 textInForm = new Form2(this);
                                textInForm.ShowDialog();

                                //survey name should appear in globalSurveyName



                            }

                            surveyName = globalSurveyName;




                        }
                        else
                        {

                            //both matches failed
                            continue;


                        }


                    }

                    
                    
                    //read the file
                    readfile(file.FullName, surveyName, partID, resultList);




                }

                //did the user want to save as xlsx or mysql?

                RadioButton selectedButton= null;
            foreach (RadioButton rb in groupBox1.Controls)
            {

                if (rb.Checked)
                {
                    selectedButton = rb;
                    break;


                }



            }

                string rbText= selectedButton.Text;

                if (rbText == "MySQL Database")
                {

                    //get db connection params
                    Form3 dbConnForm = new Form3(this);
                    dbConnForm.ShowDialog();

                    if (! dbConnectionParamsSet)
                    {

                        //user has cancelled
                        return;


                    }
                    
                    
                    
                    //try and connect to the DB

                    string connStr = "SERVER=" + dbHostName + ";DATABASE=" + dbName + ";UID=" + dbUserName + ";PASSWORD=" + dbPassword + ";PORT=" + dbPort;

                    DB db = new DB(connStr);

                    //open connection

                    try
                    {
                        db.connect();

                    }
                    catch (Exception ex)
                    {
                        //connection error

                        MessageBox.Show("There was a problem connecting to the database: " + ex.Message);
                        return;



                    }
                    
                    
                    
                    //export to mysql
                    saveAsMySQL(resultList, db);

                }
                else
                {

                    //export to excel
                    //all file data has been read. save as a single Excel file.
                    saveAsExcel(di, resultList);


                    MessageBox.Show("Your data has been saved in file: EQ-data.xlsx");



                }




            }




        }


        


        private void saveAsMySQL( List<EQresult> resultList, DB db)
        {


            try
            {

                foreach (EQresult result in resultList)
                {

                    db.addParticipantData(result);

                }

                MessageBox.Show("Data has been saved to the database OK");



            }

            catch (Exception ex)
            {


                MessageBox.Show("There was a problem sending the data to the database: " + ex.Message);
               


            }

            finally
            {
                db.close();



            }




        }







        private void saveAsExcel(DirectoryInfo di, List<EQresult> resultList)
        {

            //open an excel file in this dir for writing.

            FileInfo newFile = new FileInfo(di.FullName + @"\EQ-data.xlsx");
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(di.FullName + @"\EQ-data.xlsx");
            }


            //reformat the data into all info for a specific participant on the same line

            EQdataCollection eqd = new EQdataCollection();

            //add results
            eqd.Build(resultList);

            //build ordered set of questions
            eqd.BuildCodeSet();



            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("EQ Data");

                //Add the headers
                worksheet.Cells[1, 1].Value = "Questionnaire Name";
                worksheet.Cells[1, 2].Value = "Participant ID";
                //worksheet.Cells[1, 3].Value = "Question Code";
                //worksheet.Cells[1, 4].Value = "Answer";


                eqd.AddHeaders(worksheet);
                eqd.BuildExcel(worksheet);



                //format the headers
                /*
                using (var range = worksheet.Cells[1, 1, 1, 4])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                    range.Style.Font.Color.SetColor(Color.White);
                }
                 
 

                int row = 2;

                


                foreach (EQresult result in resultList)
                {

                    worksheet.Cells[row, 1].Value = result.surveyID;
                    worksheet.Cells[row, 2].Value = result.partID;
                    worksheet.Cells[row, 3].Value = result.qCode;
                    worksheet.Cells[row, 4].Value = result.answer;

                    row++;



                }
                 */
 




                package.Save();



            }



        }





       




        private void readfile(string filepath, string surveyID, string partID, List<EQresult> resultList)
        {

            StreamReader dh = new StreamReader(filepath);
            char[] splitOn = { '\t' };
            string qCode;
            string answer;

            EQresult eqRes;

            try
            {

                while (dh.EndOfStream == false)
                {
                    string line = dh.ReadLine();


                    //split into qcode and answer
                    string[] items = line.Split(splitOn);

                    //first item is the qCode
                    //second item is the data for that qCode
                    qCode = items[0];
                    answer = items[1];

                    eqRes = new EQresult();

                    eqRes.qCode = qCode;
                    eqRes.answer = answer;
                    eqRes.partID = partID;
                    eqRes.surveyID = surveyID;

                    resultList.Add(eqRes);


                }

                

            }


            catch (Exception ex1)
            {
                MessageBox.Show("Error reading file:" + ex1.Message );


            }

            finally
            {

                if (dh != null)
                {

                    dh.Close();

                }

            }

            




        }





    }







}
