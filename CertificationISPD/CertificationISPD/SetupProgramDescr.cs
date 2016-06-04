using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CertificationISPD
{
    public class SetupProgramDescr
    {
        public string Caption { set; get; }
        public string Description { set; get; }
        public string Name { get; set; }
        public string DisplayName { set; get; }
        public string PathName { set; get; }
        public string Started { set; get; }
        public SetupProgramDescr(string cap, string Desc, string Nam, string DN, string PN)
        {
            Caption = cap;
            Description = Desc;
            Name = Nam;
            DisplayName = DN;
            PathName = PN;
            Started = "yes";
        }
        public SetupProgramDescr(SetupProgramDescr SPD, CertifProgramDescr CPD)
        {
            Caption = SPD.Caption;
            Description = SPD.Description;
            Name = SPD.Name;
            DisplayName = SPD.DisplayName;
            PathName = SPD.PathName;
            Started = SPD.Started;
            addRelationItem(CPD);
        }
        public List<CertifProgramDescr> RelationProgram { set; get; }
        public FSTECFindCertificationProgram FFCP { set; get; }
        public void addRelationItem(CertifProgramDescr CPD)
        {
            if (RelationProgram == null) {
                RelationProgram = new List<CertifProgramDescr>();
            }
            RelationProgram.Add(CPD);
        }

        public void setRelation()
        {
            if (FFCP == null) { throw new Exception("FSTECFindCertificationProgram == null"); return; }
            if (Name == null) { throw new Exception("SetupProgramDescr == null"); return; }
            RelationProgram = FFCP.getListCert(Name);
        }

    }
}
