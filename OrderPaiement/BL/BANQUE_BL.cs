using OrderPaiement.DataAccess;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrderPaiement.BL
{
    public class BANQUE_BL
    {
        public BANQUE_DAL BANQUE_DAL;
        public DataTable GetAllBanque()
        {
            BANQUE_DAL = new BANQUE_DAL();
            return BANQUE_DAL.GetAllBanque();
        }
    }
}
