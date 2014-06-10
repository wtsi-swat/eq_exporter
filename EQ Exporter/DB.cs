
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
using MySql.Data.MySqlClient;


namespace EQ_Exporter
{
    class DB
    {

        private string connectionStr;
        private MySqlConnection conn = null;

        public DB(string connectionStr)
        {
            this.connectionStr = connectionStr;

        }

        public void connect()
        {

            //try and connect
            conn = new MySqlConnection(connectionStr);
            conn.Open();



        }

        public void close()
        {

            try
            {
                conn.Close();

            }
            catch
            {

            }



        }


       


        public void addParticipantData(EQresult result)
        {
           
            string query = "insert into eq_data(questionnaire_name, participant_id, question_code, answer) values (@qName, @partID, @qCode, @answer)";

            MySqlCommand cmd = new MySqlCommand(query, conn);
            cmd.Parameters.AddWithValue("@qName", result.surveyID);
            cmd.Parameters.AddWithValue("@partID", result.partID);
            cmd.Parameters.AddWithValue("@qCode", result.qCode);
            cmd.Parameters.AddWithValue("@answer", result.answer);


            cmd.ExecuteNonQuery();



        }






    }
}
