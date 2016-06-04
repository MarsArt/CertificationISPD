using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization.Json;
using System.Runtime.Serialization;

namespace CertificationISPD
{
    [DataContract]
    public class ItemSource
    {
        
    
        [DataMember]
        public Uri source { get; set;}
        [DataMember]
        public string LocalFilename { get; set;}
        public ItemSource(Uri src, string FN) { source = src; LocalFilename = FN;}

        public ItemSource(string uRL, string fN)
        {
            this.LocalFilename = fN;
            this.source = new Uri(uRL);

        }
    }
}
