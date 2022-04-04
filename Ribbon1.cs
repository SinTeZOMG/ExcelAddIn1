using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;


namespace ExcelAddIn1
{
    
    public partial class Ribbon1
    {
        public event System.Action Button1Clicked; 
        public event System.Action Button2Clicked;
        public event System.Action Button3Clicked;
        public event System.Action Button4Clicked;
        public event System.Action Button5Clicked;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (Button1Clicked != null)
                Button1Clicked();
        }

      

        private void button2_Click_1(object sender, RibbonControlEventArgs e)
        {
            if (Button2Clicked != null)
                Button2Clicked();
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {

        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            if (Button3Clicked != null)
                Button3Clicked();

        }
        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            if (Button4Clicked != null)
                Button4Clicked();

        }
        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            if (Button5Clicked != null)
                Button5Clicked();

        }

    }
}
