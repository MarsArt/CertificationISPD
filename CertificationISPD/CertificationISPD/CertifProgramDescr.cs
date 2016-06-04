using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CertificationISPD
{
  public  class CertifProgramDescr
    {
        public string Number { set; get; }
        public DateTime startCert { set; get;}
        public DateTime Dedline { set; get; }
        public string Name { set; get;}
        public string Descr { set; get; }
        public string labor { set; get; }
        public string Zyavitel { set; get; }

       public CertifProgramDescr(string N, DateTime SC, DateTime DL, string NM, string DESC, string Labor, string Zay){
            Number = N;
            startCert = SC;
            Dedline = DL;
            Name = NM;
            Descr = DESC;
            labor = Labor;
            Zyavitel = Zay;

        }
        public CertifProgramDescr(CertifProgramDescr cpd)
        {
            Number = cpd.Number;
            startCert = cpd.startCert;
            Dedline = cpd.Dedline;
            Name = cpd.Name;
            Descr = cpd.Descr;
            labor = cpd.labor;
            Zyavitel = cpd.Zyavitel;

        }

    }
}
