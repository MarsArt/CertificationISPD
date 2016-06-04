using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;


namespace CertificationISPD
{
  public  class FSTECFindCertificationProgram : IFindInfoinRegistry
    {

        public string PathToDB { set; get; }
        public Random rd;

        public FSTECFindCertificationProgram(string PTDB) {
            PathToDB = PTDB;
            rd = new Random();
        }

        public FSTECFindCertificationProgram()
        {
            PathToDB = "FSTEC.xlsx";
            rd = new Random();
        }

        public List<CertifProgramDescr> getListCertOld(string pattern)
        {
            Excel.Application excelapp = new Excel.Application();
            excelapp.Visible = false;
            Excel.Workbooks excelappworkbooks = excelapp.Workbooks;
            System.Windows.Forms.MessageBox.Show(PathToDB);
            string stolb="-1";
            string strok="-1";
            Excel.Range excelcells=null;
            try
            {
                Excel.Workbook excelappworkbook = excelapp.Workbooks.Open(@Path.Combine(Application.StartupPath, "FSTEC.xlsx"),
                                                        Type.Missing, true, Type.Missing,
                                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                        Type.Missing, Type.Missing);
                Excel.Sheets excelsheets = excelappworkbook.Worksheets;
                Excel.Worksheet excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
                excelcells = excelworksheet.Cells.Find(pattern, Type.Missing, Type.Missing, Excel.XlLookAt.xlPart, Type.Missing,
                                                           Excel.XlSearchDirection.xlNext,
                                                          Type.Missing, Type.Missing, Type.Missing);
                 stolb = Convert.ToString(excelcells.Column);
                 strok = Convert.ToString(excelcells.Rows.Row);
            }
            catch (Exception ex) { System.Windows.Forms.MessageBox.Show(ex.Message); }

          
          
            System.Windows.Forms.MessageBox.Show("Строка: "+ strok + "\n Столбец: "+ stolb + "\n поиск: "+ excelcells.Value2);
            excelapp.Quit();
            return new List<CertifProgramDescr>();
        }
        public List<CertifProgramDescr> getListCert(string pattern)
        {
            /* Excel.Application excelapp = new Excel.Application();
             excelapp.Visible = false;
             Excel.Workbooks excelappworkbooks = excelapp.Workbooks;
             string stolb = "-1";
             string strok = "-1";
             Excel.Range currentFind = null;
             Excel.Range firstFind = null;

             Excel.Range tempVal = null;

             Excel.Workbook excelappworkbook = excelapp.Workbooks.Open(@Path.Combine(Application.StartupPath, "FSTEC.xlsx"),
                                                         Type.Missing, true, Type.Missing,
                                                         Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                         Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                         Type.Missing, Type.Missing);
                 Excel.Sheets excelsheets = excelappworkbook.Worksheets;
                 Excel.Worksheet excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);



             currentFind = excelworksheet.Cells.Find(pattern, Type.Missing, Type.Missing, 
                                                     Excel.XlLookAt.xlPart, Type.Missing,
                                                     Excel.XlSearchDirection.xlNext,
                                                     Type.Missing, Type.Missing, Type.Missing);
             while (currentFind != null)
             {
                 // Keep track of the first range you find. 
                 if (firstFind == null)
                 {
                     firstFind = currentFind;
                 }

                 // If you didn't move to a new range, you are done.
                 else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                       == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                 {
                     break;
                 }

                 tempVal = currentFind;
                 MessageBox.Show(tempVal[4][4].Value2); //ok
                 stolb = Convert.ToString(currentFind.Column);
                 strok = Convert.ToString(currentFind.Rows.Row);
                 MessageBox.Show("Строка: " + strok + "\n Столбец: " + stolb + "\n поиск: " + currentFind.Value2);
                 currentFind = excelworksheet.Cells.FindNext(currentFind); 
             }    
             excelapp.Quit();*/
            List<CertifProgramDescr> LRD = new List<CertifProgramDescr>();
            for(int i=0;i< rd.Next(1, 5); i++)
              LRD.Add(new CertifProgramDescr(i.ToString(), DateTime.Now, DateTime.Now, pattern+"name"+i.ToString(), "DESK", "Labor", "Zey"));
            return LRD;
        }
    }
}
