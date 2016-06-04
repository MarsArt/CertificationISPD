using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CertificationISPD
{
   public interface IFindInfoinRegistry
    {
         string PathToDB { set; get; } 
         List<CertifProgramDescr> getListCert(string pattern);
    }
}
