
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
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;


namespace EQ_Exporter
{
    class EQparticipant
    {

        //map qCode to answer for this participant


        private Dictionary<string, string> questionHash;

        public EQparticipant()
        {

            questionHash = new Dictionary<string, string>();

        }

        public void addResult(EQresult result)
        {
            string qCode= result.qCode;
            string answer = result.answer;

            questionHash[qCode] = answer;




        }

        public SortedSet<string> getQcodeSet()
        {
            //get all qCodes
            SortedSet<string> qCodeSet = new SortedSet<string>();

            foreach (string qCode in questionHash.Keys)
            {
                qCodeSet.Add(qCode);


            }

            return qCodeSet;



        }


        public void BuildExcel(string qName, ExcelWorksheet worksheet, SortedSet<string> qCodeSet, string partID, int row)
        {

            //create a row in this table for this participant
            //add questionnaire name and participant-id
            worksheet.Cells[row, 1].Value = qName;
            worksheet.Cells[row, 2].Value = partID;

            int col = 3;

           
            //show the answer to each question

            foreach (string qCode in qCodeSet)
            {

                if (questionHash.ContainsKey(qCode))
                {

                    worksheet.Cells[row, col].Value = questionHash[qCode];

                }
                else
                {
                    worksheet.Cells[row, col].Value = "";

                }


                col++;




            }







        }







    }
}
