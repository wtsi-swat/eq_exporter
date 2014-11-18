
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

            //user clicked button to export data

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


                //did the user want to do variable splitting?
                if (checkBox1.Checked)
                {
                    //yes

                    //get the skipped text
                    string skipText = textBox1.Text;

                    //get no-answer
                    string noAnswerText= textBox2.Text;

                    //get don't know
                    string dontKnowText = textBox3.Text;

                    //get not applicable
                    string notAppText = textBox4.Text;

                    if (string.IsNullOrWhiteSpace(skipText) || string.IsNullOrWhiteSpace(noAnswerText)  || string.IsNullOrWhiteSpace(dontKnowText) || string.IsNullOrWhiteSpace(notAppText))
                    {

                        MessageBox.Show("You must have values for each of the 4 skip types");
                        return;


                    }


                    //split vars

                    splitVars(resultList, skipText, noAnswerText, dontKnowText, notAppText);




                }
                







                //did the user want to save as xlsx or mysql?

                RadioButton selectedButton = null;
                foreach (RadioButton rb in groupBox1.Controls)
                {

                    if (rb.Checked)
                    {
                        selectedButton = rb;
                        break;


                    }



                }

                string rbText = selectedButton.Text;

                if (rbText == "MySQL Database")
                {

                    //get db connection params
                    Form3 dbConnForm = new Form3(this);
                    dbConnForm.ShowDialog();

                    if (!dbConnectionParamsSet)
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





        private void saveAsMySQL(List<EQresult> resultList, DB db)
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


        private EQresult copyResult(EQresult orig, string newCode, string newAnswer)
        {

            EQresult res = new EQresult();
            res.partID = orig.partID;
            res.surveyID = orig.surveyID;
            res.qCode = newCode;
            res.answer = newAnswer;

            return res;



        }


        private bool checkPatternMatch(string qCode, Regex pattern)
        {

            Match match = pattern.Match(qCode);

            if (match.Success)
            {
                return true;


            }
            else
            {

                return false;

            }




        }




        private void splitVars(List<EQresult> resultList, string missingStr, string noAnswerText, string dontKnowText, string notAppText)
        {

            //search for compound vars which can be split into chunks where each chunk is a new var
            //missingStr is whatever was chosen to represent missing values


            //delete list: what we need to remove from the data list
            List<EQresult> dList = new List<EQresult>();


            //add list: new things to add
            List<EQresult> aList = new List<EQresult>();



            Regex pattAVEW = new Regex(@"([WMY]):(\d+)");
            Regex pattDEXAM = new Regex(@"^(\d+)/(\d+)/(\d+) .+");
            Regex pattTAC2 = new Regex(@"^TAC2[A-E]-R$");
            Regex pattBlood= new Regex(@"^(.+):(.+)$");
            Regex pattTOB = new Regex(@"^(\d):(.+)$");


            //blood samples
            //mapping of the qCode of the result to the qCode of the corresponding barcode
            var bloodMap= new Dictionary<string, string>();

            bloodMap["PST"]= "PSTB1";
            bloodMap["EDTA1A"]= "EDTAB1A";
            bloodMap["EDTA1B"]= "EDTAB1B";
            bloodMap["EDTA1C"]= "EDTAB1C";
            bloodMap["PL1"]= "PLTB1";
            bloodMap["SPU"]= "SPUB";
            bloodMap["NAF1"]= "NAFB1";
            bloodMap["NAF2"]= "NAFB2";
            bloodMap["PST2"]= "PSTB2";


            foreach (EQresult res in resultList)
            {

                string qCode = res.qCode;
                string answer = res.answer;


                //blood sample?
                if(bloodMap.Keys.Contains(qCode)){

                    //split the value into barcode and result
                    //e.g. PST	1:111

                    

                    EQresult newRes;
                    EQresult newResBC;  //barcode

                    //was this skipped
                    if (answer == missingStr || answer == noAnswerText || answer == dontKnowText || answer == notAppText)
                    {
                        //create second result (barcode)
                        newRes = copyResult(res, bloodMap[qCode], answer);
                        aList.Add(newRes);

                    }
                    else{

                        //split into barcode and result

                        Match match = pattBlood.Match(answer);
                        string result = match.Groups[1].Value;
                        string bc= match.Groups[2].Value;

                        //result
                        newRes = copyResult(res, qCode, result);
                        aList.Add(newRes);

                        //barcode
                        newResBC = copyResult(res, bloodMap[qCode], bc);
                        aList.Add(newResBC);

                        //delete the original
                        dList.Add(res);


                    }


                }



                //AVEW	W:123

                else if (qCode == "AVEW")
                {
                    //split into Weeks, Months, Years.
                    

                    EQresult avewRes;
                    EQresult avemRes;
                    EQresult aveyRes;

                    //might have been skipped
                    if (answer == missingStr || answer == noAnswerText || answer == dontKnowText || answer == notAppText)
                    {

                        avewRes = copyResult(res, "AVEW", answer);
                        avemRes = copyResult(res, "AVEM", answer);
                        aveyRes = copyResult(res, "AVEY", answer);


                    }
                    else
                    {

                        Match match = pattAVEW.Match(answer);

                        string wmy = match.Groups[1].Value;
                        string time = match.Groups[2].Value;

                        


                        if (wmy == "W")
                        {

                            //create 3 new results and remove this one
                            avewRes = copyResult(res, "AVEW", time);
                            avemRes = copyResult(res, "AVEM", missingStr);
                            aveyRes = copyResult(res, "AVEY", missingStr);



                        }
                        else if (wmy == "M")
                        {
                            //create 3 new results and remove this one
                            avewRes = copyResult(res, "AVEW", missingStr);
                            avemRes = copyResult(res, "AVEM", time);
                            aveyRes = copyResult(res, "AVEY", missingStr);


                        }
                        else if (wmy == "Y")
                        {
                            //create 3 new results and remove this one
                            avewRes = copyResult(res, "AVEW", missingStr);
                            avemRes = copyResult(res, "AVEM", missingStr);
                            aveyRes = copyResult(res, "AVEY", time);


                        }
                        else
                        {

                            throw new Exception("Illegal prefix for code:AVEW");
                        }


                    }

                    

                    //remove original res
                    dList.Add(res);

                    //add 3 new ones
                    aList.Add(avewRes);
                    aList.Add(avemRes);
                    aList.Add(aveyRes);




                }

                else if (qCode == "AVEW-2")
                {

                    //duplicate of AVEW: delete

                    dList.Add(res);



                }

                else if (qCode == "CRWD")
                {

                    //delete CRWD
                    dList.Add(res);



                }

                else if (qCode == "START")
                {

                    //change this to SITECODE

                    EQresult newRes = copyResult(res, "SITECODE", answer);
                    aList.Add(newRes);

                    //delete original
                    dList.Add(res);


                }

                


                else if (checkPatternMatch(qCode, pattTAC2))
                {

                    //delete
                    dList.Add(res);





                }


                else if (qCode == "TOB9")
                {

                    EQresult days;
                    EQresult weeks;
                    EQresult months;
                    EQresult years;

                    if (answer == missingStr || answer == noAnswerText || answer == dontKnowText || answer == notAppText)
                    {
                        days = copyResult(res, "TOB9D", answer);
                        weeks = copyResult(res, "TOB9W", answer);
                        months = copyResult(res, "TOB9M", answer);
                        years = copyResult(res, "TOB9Y", answer);



                    }
                    else
                    {

                        Match match = pattTOB.Match(answer);

                        string dwmy= match.Groups[1].Value;
                        string timespan = match.Groups[2].Value;

                        string answerDays="0";
                        string answerWeeks="0";
                        string answerMonths="0";
                        string answerYears="0";

                        if (dwmy == "1")
                        {

                            answerDays = timespan;
                           

                        }
                        else if (dwmy == "2")
                        {

                            answerWeeks = timespan;
                        }

                        else if (dwmy == "3")
                        {

                            answerMonths = timespan;
                        }

                        else if (dwmy == "8")
                        {

                            answerYears = timespan;


                        }
                        else
                        {

                            throw new Exception("unknown timespan for TOB9");
                        }




                        days = copyResult(res, "TOB9D", answerDays);
                        weeks = copyResult(res, "TOB9W", answerWeeks);
                        months = copyResult(res, "TOB9M", answerMonths);
                        years = copyResult(res, "TOB9Y", answerYears);



                    }

                    //remove original res
                    dList.Add(res);

                    //add 3 new ones
                    aList.Add(days);
                    aList.Add(weeks);
                    aList.Add(months);
                    aList.Add(years);




                }





                else if (qCode == "DEXAM")
                {
                    //DEXAM	22/08/2014 13:38:42

                    //split into d/m/y and ignore timestamp part
                    

                    EQresult dexam;
                    EQresult mexam;
                    EQresult yexam;

                    if (answer == missingStr || answer == noAnswerText || answer == dontKnowText || answer == notAppText)
                    {
                        dexam = copyResult(res, "DEXAM", answer);
                        mexam = copyResult(res, "MEXAM", answer);
                        yexam = copyResult(res, "YEXAM", answer);



                    }
                    else
                    {

                        Match match = pattDEXAM.Match(answer);

                        string days = match.Groups[1].Value;
                        string months = match.Groups[2].Value;
                        string years = match.Groups[3].Value;

                       

                        dexam = copyResult(res, "DEXAM", days);
                        mexam = copyResult(res, "MEXAM", months);
                        yexam = copyResult(res, "YEXAM", years);



                    }

                    

                    //remove original res
                    dList.Add(res);

                    //add 3 new ones
                    aList.Add(dexam);
                    aList.Add(mexam);
                    aList.Add(yexam);




                }




            }

            //build a new result-list, i.e. remove things in dList and add things in aList
            foreach (EQresult res in dList)
            {
                //delete from main list
                resultList.Remove(res);


            }

            //add new items
            resultList.AddRange(aList);

            






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
                MessageBox.Show("Error reading file:" + ex1.Message);


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
