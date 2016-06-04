using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CertificationISPD
{
  public  class SetupCertificationRelation
    {
        public SetupProgramDescr PrSetup { set; get; } 
        public FSTECFindCertificationProgram FFCP { set; get; }
        public List<CertifProgramDescr> RelationProgram { set; get; }
        public SetupCertificationRelation(SetupProgramDescr PS) {
            PrSetup = PS;
            RelationProgram = new List<CertifProgramDescr>();
        }
        public SetupCertificationRelation(SetupProgramDescr PS, FSTECFindCertificationProgram FFC)
        {
            PrSetup = PS;
            FFCP = FFC;
            RelationProgram = new List<CertifProgramDescr>();
        }
        public void addRelationItem(CertifProgramDescr CPD) {
            RelationProgram.Add(CPD);
        }

        public void setRelation() {
            if (FFCP == null) { throw new Exception("FSTECFindCertificationProgram == null"); return; }
            if (PrSetup == null) {throw new Exception("SetupProgramDescr == null"); return; }
            RelationProgram=FFCP.getListCert(PrSetup.Name);
        }
    }
}
