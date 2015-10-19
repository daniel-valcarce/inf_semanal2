using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Inf_Semanal
{
    public partial class Formcsv : Form
    {
        public Formcsv()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook newWorkbook = app.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet hojaexcel;
            hojaexcel = (Microsoft.Office.Interop.Excel.Worksheet)app.Worksheets[1];
           
            app.Visible = true;
        }
    }
}
