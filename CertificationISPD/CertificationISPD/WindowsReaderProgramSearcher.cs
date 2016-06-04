using System;
using System.Collections.Generic;
using System.Management;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace CertificationISPD
{
    public class WindowsReaderProgramSearcher : IReaderProgram
    {
        public List<SetupProgramDescr> ListSetupWidthRelation { set; get; }
       // public <List>

        public WindowsReaderProgramSearcher() {
            ListSetupWidthRelation = new List<SetupProgramDescr>();
        }
        public WindowsReaderProgramSearcher(List<SetupProgramDescr> scr)
        {
            ListSetupWidthRelation = scr;
        }

        public void ReadProgram()
        {
            ManagementObjectSearcher searcher3 = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_Service");

            ListSetupWidthRelation.Add(new SetupProgramDescr("asdsad", "sad", "fds", "asd", "asdsad")); //Очень плохо надо поменять на порождающий паттерн.
            foreach (ManagementObject queryObj in searcher3.Get())
            {
                try
                {
                    /*Проверять все параметры и добавлять если есть хотя бы DisplayName*/
                    ListSetupWidthRelation.Add(new SetupProgramDescr(
                                                    (queryObj["Caption"].ToString() != null) ? queryObj["Caption"].ToString() : "no info",
                                                    (queryObj["Description"].ToString() != null) ? queryObj["Description"].ToString() : "no info",
                                                    (queryObj["Name"].ToString() != null) ? queryObj["Name"].ToString() : "no info",
                                                    (queryObj["DisplayName"].ToString() != null) ? queryObj["DisplayName"].ToString() : "no info",
                                                    (queryObj["PathName"].ToString() != null) ? queryObj["PathName"].ToString() : "no info"
                                                     ));
                 
                }
                catch (Exception ex) {; /*MessageBox.Show(ex.Message);*/ }
            }
        }


        public void fillListCert()
        {
            FSTECFindCertificationProgram FFCPI = new FSTECFindCertificationProgram();
           foreach (SetupProgramDescr pr in ListSetupWidthRelation) {
                pr.FFCP = FFCPI;
                pr.setRelation();
            }

         /*   int i = 1;
            int k = 0;
            string stolb;
            excelapp = new Excel.Application();
            excelapp.Visible = false;
            excelappworkbooks = excelapp.Workbooks;
            excelappworkbook = excelapp.Workbooks.Open(Application.StartupPath + @"\template\2.xls",
                      Type.Missing, true, Type.Missing,
    "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
     Type.Missing, Type.Missing, Type.Missing, Type.Missing,
     Type.Missing, Type.Missing);
            excelsheets = excelappworkbook.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            excelcells = excelworksheet.Cells.Find("q", Missing.Value, Missing.Value, Excel.XlLookAt.xlPart, Missing.Value,
               Excel.XlSearchDirection.xlNext,
               Missing.Value, Missing.Value, Missing.Value);

            stolb = Convert.ToString(excelcells.Column);
            strok = Convert.ToString(excelcells.Rows.Row);
            System.Windows.Forms.MessageBox.Show(getAdres(Convert.ToInt32(poisk), (Convert.ToInt32(stolb)) - 1));
            excelapp.Quit();*/
        }
    }
}
