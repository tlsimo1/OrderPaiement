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
    public class DETAILFACTURE_DAL
    {
        public string GetLastID()
        {
            DataTable dt = new DataTable();
            string LastID = "";
            try
            {
                using (SqlConnection cnx = new SqlConnection(Connection.stringConnection))
                {
                    SqlCommand cmd = new SqlCommand($@"SELECT IDENT_CURRENT('FACTURE')", cnx);
                    cnx.Open();
                    var returnValue = cmd.ExecuteScalar();
                    if (returnValue != null)
                        LastID = returnValue.ToString();

                }
            }
            catch (Exception e)
            {
                Console.WriteLine("" + e);
            }
            return LastID;
        }
        public void Insert_DeatilFacture(List<DETAILFACTURE> Listdetailfac)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Connection.stringConnection))
                {
                    connection.Open();
                    foreach (var detailfacture in Listdetailfac)
                    {
                        SqlCommand com = new SqlCommand($@" INSERT INTO [dbo].[DETAIL FACTURE] 
                                                       ([DATE FACTURE]
                                                       ,[NUM FACTURE]
                                                       ,[OBJET]
                                                       ,[MONTANT HT]
                                                       ,[TAUX TVA]
                                                       ,[MONTANT TTC]
                                                       ,[TVA]
                                                       ,[LIGNE]
                                                       ,[FACTURE_ID]
                                                       ,[DATE BANQUE])
                                                       VALUES 
                                                     ('{detailfacture.DateFacture}',
                                                       '{detailfacture.NumFacture}',
                                                       '{detailfacture.Objet}',
                                                       {detailfacture.MontantHT.ToString().Replace(',', '.')},
                                                       '{detailfacture.TauxTAV}',
                                                       {detailfacture.MontantTTC.ToString().Replace(',', '.')},
                                                       {detailfacture.TVA.ToString().Replace(',', '.')},
                                                       '{detailfacture.Ligne}',
                                                       {detailfacture.FactureID},
                                                       '{detailfacture.DateBanque}'
                                                       )", connection);

                        com.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("" + e);
            }

        }

        public void UPDATE_DeatilFacture(List<DETAILFACTURE> UPDATEListdetailfac)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Connection.stringConnection))
                {
                    connection.Open();
                    foreach (var detailfacture in UPDATEListdetailfac)
                    {
                        SqlCommand com = new SqlCommand($@" UPDATE [dbo].[DETAIL FACTURE]
                                                       SET [DATE FACTURE] = '{detailfacture.DateFacture}'
                                                       ,[NUM FACTURE] = '{detailfacture.NumFacture}'
                                                       ,[OBJET] = '{detailfacture.Objet}'
                                                       ,[MONTANT HT] = {detailfacture.MontantHT.ToString().Replace(',', '.')}
                                                       ,[TAUX TVA] = '{detailfacture.TauxTAV}'
                                                       ,[MONTANT TTC] = {detailfacture.MontantTTC.ToString().Replace(',', '.')}
                                                       ,[TVA] = {detailfacture.TVA.ToString().Replace(',', '.')}
                                                       ,[LIGNE] = '{detailfacture.Ligne}'
                                                       ,[DATE BANQUE] = '{detailfacture.DateBanque}'
                                                       WHERE ID_DETAIL ={detailfacture.ID_DETAIL}", connection);

                        com.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("" + e);
            }

        }
        public void DELETE_DetailFacture(int IDDetail)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Connection.stringConnection))
                {
                    connection.Open();
                    
                        SqlCommand com = new SqlCommand($@"DELETE FROM [dbo].[DETAIL FACTURE]
                                                       WHERE ID_DETAIL ={IDDetail} ", connection);

                        com.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("" + e);
            }

        }
        public DataTable ImprimerDetailFacture(int idFac)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection cnx = new SqlConnection(Connection.stringConnection))
                {
                    SqlDataAdapter DA = new SqlDataAdapter($@"SELECT ID_DETAIL, CONVERT (varchar(10), [DATE FACTURE], 103) as 'DATE FACTURE',
                                                          [NUM FACTURE], OBJET, [MONTANT HT], [TAUX TVA], [MONTANT TTC], TVA, LIGNE, FACTURE_ID,
                                                          CONVERT (varchar(10), [DATE BANQUE], 103) as 'DATE BANQUE' FROM  [DETAIL FACTURE] 
                                                          where  FACTURE_ID = '{idFac}'", cnx);

                    DA.Fill(dt);
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
