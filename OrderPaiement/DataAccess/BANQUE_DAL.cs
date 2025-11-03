using OrderPaiement.Classe;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrderPaiement.DataAccess
{
    public class BANQUE_DAL
    {
        public DataTable GetAllBanque()
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection cnx = new SqlConnection(Connection.stringConnection))
                {
                    SqlDataAdapter da = new SqlDataAdapter($@"select * from BANQUE ", cnx);
                    da.Fill(dt);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("" + e);
            }
            return dt;
        }
    }
}
