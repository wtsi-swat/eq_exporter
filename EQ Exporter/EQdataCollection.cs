
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
    class EQdataCollection
    {

        //map of questionnaire name to EQuestionnaire
        private Dictionary<string, EQuestionnaire> qHash;

        //ordered set of qCodes
        private SortedSet<string> qCodeSet;


        public EQdataCollection()
        {

            qHash = new Dictionary<string, EQuestionnaire>();


        }

        public void AddHeaders(ExcelWorksheet worksheet)
        {

            int col = 3;
            
            foreach (string qCode in qCodeSet)
            {
                worksheet.Cells[1, col].Value = qCode;
                col++;


            }

        }


        public void BuildExcel(ExcelWorksheet worksheet)
        {

            //for each questionnaire: add data
            String qName;
            EQuestionnaire eq;
            int row = 2;


            foreach (KeyValuePair<string, EQuestionnaire> kv in qHash)
            {

                qName= kv.Key;
                eq= kv.Value;


                eq.BuildExcel(qName, worksheet, qCodeSet, ref row);     //note: row is passed by reference so each questionnaire uses different rows


            }




        }


        public void BuildCodeSet()
        {
            //get a set of all qCodes
            qCodeSet = new SortedSet<string>();

            foreach (EQuestionnaire eq in qHash.Values)
            {
                qCodeSet.UnionWith(eq.getQcodeSet());



            }




        }



        public void Build(List<EQresult> resultList)
        {

            string qName;
            EQuestionnaire eq;
            
            //for each result, assign to a questionnaire
            foreach (EQresult result in resultList)
            {
                qName = result.surveyID;

                if (qHash.ContainsKey(qName))
                {
                    qHash[qName].addResult(result);


                }
                else
                {
                    //new Q.
                    eq = new EQuestionnaire();

                    qHash[qName] = eq;

                    eq.addResult(result);


                }




            }




        }
        


    }
}
