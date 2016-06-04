using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Web.Script.Serialization;
namespace CertificationISPD
{
    public class Loader : ILoaderDB
    {
        public List<ItemSource> source;
        public string SavePath { set; get; }
        public Loader(List<ItemSource> src, string SP)
        {
            source = new List<ItemSource>(src);
            SavePath = SP;
        }
        public Loader()
        {
            source = new List<ItemSource>();
            SavePath = "";
        }

   

        public void addSourceItem(string URL, string FN)
        {
            source.Add(new ItemSource(URL, FN));

        }

        public void removeSourceItem(int id)
        {
            source.RemoveAt(id);
        }


        public void LoadFiles()
        {
            WebClient WC = new WebClient();
            string path="";
            
            foreach(ItemSource IS in source)
                {
                     path = (SavePath=="")? IS.LocalFilename : SavePath + "/" + IS.LocalFilename;
                     WC.DownloadFile(IS.source, path);
                }
        }

        public void loadConfig()
        {
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(List<ItemSource>));
            MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(Properties.Settings.Default.source));
            source = (List<ItemSource>)serializer.ReadObject(ms);
        }
        public void saveConfig()
        {  
            DataContractJsonSerializer jsonFormatter = new DataContractJsonSerializer(typeof(List<ItemSource>));
            MemoryStream stream1 = new MemoryStream();
            jsonFormatter.WriteObject(stream1, source);
            string res = Encoding.UTF8.GetString(stream1.ToArray());
        /*  var json = new JavaScriptSerializer().Serialize(source);
            string res = json.ToString();
            var source = new JavaScriptSerializer().Deserialize<List<ItemSource>>(res);*/
            Properties.Settings.Default.source = res;
            Properties.Settings.Default.Save();
          //  MessageBox.Show(res);
        }
    }
}
