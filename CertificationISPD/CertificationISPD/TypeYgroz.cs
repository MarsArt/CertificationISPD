//using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CertificationISPD
{
    public class TypeYgroz
    {
        public int code { set;
            get; }
        public string name { set; get; }
        public TypeYgroz()
        {
            code = -1;
            name = "Atcbr ,kz[f ve[f";
        }
        public TypeYgroz(int c, string n)
        {
            code = c;
            name = n;
        }
    }
}
