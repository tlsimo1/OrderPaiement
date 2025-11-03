using OrderPaiement.Classe;
using OrderPaiement.DataAccess;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrderPaiement.BL
{
    public class DETAILFACTURE_BL
    {
        public DETAILFACTURE_DAL DETAILFACTURE;
        public string GetLastID()
        {
            DETAILFACTURE = new DETAILFACTURE_DAL();
            return DETAILFACTURE.GetLastID();
        }
        public void DELETE_DetailFacture(int IDDetail)
        {
            DETAILFACTURE = new DETAILFACTURE_DAL();
            DETAILFACTURE.DELETE_DetailFacture(IDDetail);
        }
        public void Insert_DeatilFacture(List<DETAILFACTURE> Listdetailfac)
        {
            DETAILFACTURE = new DETAILFACTURE_DAL();
            DETAILFACTURE.Insert_DeatilFacture(Listdetailfac);
        }
        public void Update_DeatilFacture(List<DETAILFACTURE> UpdateListdetailfac)
        {
            DETAILFACTURE = new DETAILFACTURE_DAL();
            DETAILFACTURE.UPDATE_DeatilFacture(UpdateListdetailfac);
        }
        public DataTable ImprimerDetailFacture(int idFac)
        {
            DETAILFACTURE = new DETAILFACTURE_DAL();
            return DETAILFACTURE.ImprimerDetailFacture(idFac);
        }
    }
}
