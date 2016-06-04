using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace CertificationISPD
{
    public class WordReportGenerator : IReportGenerator
    {
        public Dictionary<string, string> RaplaceDict { set; get; }
        public string template { set; get; }
        public string PathToSave { set; get; }

        public WordReportGenerator()
        {
            RaplaceDict = new Dictionary<string, string>();
            PathToSave = "";
        }
        public void Add(string key, string val)
        {
            RaplaceDict.Add(key, val);
        }
        public void Generate()
        {
            if (template != "" || template != null)
            {
                PathToSave = (PathToSave == "" || PathToSave == null) ? "ReportDefault.docx" : PathToSave;
                File.Copy(template, PathToSave, true);
                Word.Application applicat = new Word.Application();
                Word.Document Doc = applicat.Documents.Open(@PathToSave);
                Word.Find findObject = applicat.Selection.Find;
                foreach (KeyValuePair<string, string> keyValue in RaplaceDict)
                {
                    SearchReplace(findObject, keyValue.Key, keyValue.Value);
                }
                Doc.Close();
            }
        }
        public void SearchReplace(Word.Find findObject, string fintpattern, string raplaceSrting)
        {
            findObject.ClearFormatting();
            findObject.Text = fintpattern;
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = raplaceSrting;
            object replaceAll = Word.WdReplace.wdReplaceAll;
            findObject.Execute(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                ref replaceAll, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }
    }
}
