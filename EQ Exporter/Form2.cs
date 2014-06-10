
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

namespace EQ_Exporter
{
    public partial class Form2 : Form
    {
        private Form1 baseForm;
        
        
        public Form2(Form1 baseForm)
        {
            InitializeComponent();

            this.baseForm = baseForm;


        }

        private void button1_Click(object sender, EventArgs e)
        {

            //OK button clicked

            //pass the text back to the calling form

            baseForm.globalSurveyName = textBox1.Text;

            //close this form
            this.Dispose();



        }
    }
}
