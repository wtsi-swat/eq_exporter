
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
    class EQuestionnaire
    {

        
        //map participantID to object
        private Dictionary<string, EQparticipant> partHash;

        public EQuestionnaire()
        {
            partHash = new Dictionary<string, EQparticipant>();


        }


        public void BuildExcel(string qName, ExcelWorksheet worksheet, SortedSet<string> qCodeSet, ref int row)
        {

            //build a row in the table for each participant
            string partID;
            EQparticipant eqp;


            foreach (KeyValuePair<string, EQparticipant> kv in partHash)
            {

                partID = kv.Key;
                eqp = kv.Value;


                eqp.BuildExcel(qName, worksheet, qCodeSet, partID, row);
                row++;


            }







        }





        public SortedSet<string> getQcodeSet()
        {
            //get all qCodes
            SortedSet<string> qCodeSet = new SortedSet<string>();

            foreach (EQparticipant eqp in partHash.Values)
            {
                qCodeSet.UnionWith(eqp.getQcodeSet());


            }

            return qCodeSet;

        }


        public void addResult(EQresult result)
        {

            //which participant?
            string part = result.partID;

            //get EQparticipant

            EQparticipant eqPart;

            if (partHash.ContainsKey(part))
            {
                eqPart = partHash[part];
                eqPart.addResult(result);




            }
            else
            {
                eqPart = new EQparticipant();
                partHash[part] = eqPart;
                eqPart.addResult(result);



            }



        }









    }
}
