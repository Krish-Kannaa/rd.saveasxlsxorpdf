using OfficeOpenXml;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EPPLUS_SaveAsXlsxOrPdf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Save_btn_Click(object sender, EventArgs e)
        {
            SaveAsPdf("C:\\Users\\kannan\\Downloads\\test.pdf");
        }

        private bool SaveAsPdf(string saveAsLocation)
        {
            string saveas = (saveAsLocation.Split('.')[0]) + ".pdf";
            try
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(saveAsLocation);

                //Save the document in PDF format

                workbook.SaveToFile(saveas, Spire.Xls.FileFormat.PDF);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }
        //public static MemoryStream ConvertXLSXtoPDF(MemoryStream stream)
        //{
          
        //}
            
    }
}
