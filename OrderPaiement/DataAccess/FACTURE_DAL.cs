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
    public class FACTURE_DAL
    {
        public void Insert_Facture(FACTURE facture)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Connection.stringConnection))
                {
                    connection.Open();
                    SqlCommand com = new SqlCommand($@"INSERT INTO [dbo].[FACTURE]
                                                    ([NUMOP]
                                                    ,[DATE SAISIE]
                                                    ,[MODE SAISIE]
                                                    ,[DATE ECHEANCE]
                                                    ,[VALEUR]
                                                    ,[FOURNISSEUR_ID]
                                                    ,[BANQUE_ID])
                                                     VALUES
                                                     ('{facture.NUMOP}',
                                                       '{facture.DateSaisie}',
                                                       '{facture.Mode}',
                                                       '{facture.DateEcheance}',
                                                       '{facture.Valeur}',
                                                       '{facture.Fournisseur_ID}',
                                                       '{facture.Banque_ID}')", connection);

                    com.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("" + e);
            }

        }
        public void Update_Facture(FACTURE Updatefacture)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Connection.stringConnection))
                {
                    connection.Open();
                    SqlCommand com = new SqlCommand($@"UPDATE [dbo].[FACTURE]
                                                   SET [NUMOP] = '{Updatefacture.NUMOP}'
                                                   ,[DATE SAISIE] ='{Updatefacture.DateSaisie}'
                                                   ,[MODE SAISIE] ='{Updatefacture.Mode}'
                                                   ,[DATE ECHEANCE] ='{Updatefacture.DateEcheance}'
                                                   ,[VALEUR] = '{Updatefacture.Valeur}'
                                                   ,[FOURNISSEUR_ID] ={Updatefacture.Fournisseur_ID}
                                                   ,[BANQUE_ID] ={Updatefacture.Banque_ID}
                                                   WHERE ID_FACTURE = {Updatefacture.ID_FACTURE} ", connection);

                    com.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("" + e);
            }

        }
        public void Delete_FactureAndDetail(int IDFAC)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Connection.stringConnection))
                {
                    connection.Open();
                    SqlCommand com = new SqlCommand($@"DELETE FROM [dbo].[FACTURE]
                                                     WHERE ID_FACTURE ={IDFAC}", connection);
                    com.ExecuteNonQuery();
                    SqlCommand com2 = new SqlCommand($@"DELETE FROM [dbo].[DETAIL FACTURE]
                                                     WHERE FACTURE_ID={IDFAC}", connection);
                    com2.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("" + e);
            }

        }
        public DataTable GetALLFACTURE()
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection cnx = new SqlConnection(Connection.stringConnection))
                {
                    SqlDataAdapter DA = new SqlDataAdapter($@"SELECT FACTURE.ID_FACTURE,FOURNISSEUR.ID_FR, FACTURE.NUMOP, 
                                            FOURNISSEUR.[CODE FOURNISSEUR], 
                                            FOURNISSEUR.NOM, FACTURE.[DATE SAISIE], 
                                            FACTURE.[MODE SAISIE], FACTURE.[DATE ECHEANCE], 
                                            FACTURE.VALEUR FROM FACTURE INNER JOIN
                                            FOURNISSEUR ON FACTURE.FOURNISSEUR_ID = FOURNISSEUR.ID_FR",
                                            cnx);
                    
                    DA.Fill(dt);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("" + e);
            }
            return dt;
        }
        public DataTable GetALLFACTUREDETAIL()
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection cnx = new SqlConnection(Connection.stringConnection))
                {
                    SqlDataAdapter DA = new SqlDataAdapter($@"SELECT FACTURE.NUMOP,CONVERT (varchar(10),FACTURE.[DATE SAISIE], 103) as 'DATE SAISIE',
                                                            FACTURE.[MODE SAISIE], 
                                                            CONVERT (varchar(10),FACTURE.[DATE ECHEANCE], 103) as 'DATE ECHEANCE',
                                                            FACTURE.VALEUR, FOURNISSEUR.[CODE FOURNISSEUR], 
                                                            FOURNISSEUR.NOM, FOURNISSEUR.IGE, FOURNISSEUR.[IF], 
                                                            BANQUE.NOM AS BANQUE,
						                                    CONVERT (varchar(10), [DETAIL FACTURE].[DATE FACTURE], 103) as 'DATE FACTURE',[DETAIL FACTURE].[NUM FACTURE], 
						                                    [DETAIL FACTURE].[MONTANT HT], [DETAIL FACTURE].[TAUX TVA], 
                                                            [DETAIL FACTURE].[MONTANT TTC], [DETAIL FACTURE].TVA, 
                                                            [DETAIL FACTURE].OBJET,[DETAIL FACTURE].LIGNE
                                                            FROM BANQUE INNER JOIN
                                                            FACTURE ON BANQUE.ID_BANQUE = FACTURE.BANQUE_ID INNER JOIN
                                                            FOURNISSEUR ON FACTURE.FOURNISSEUR_ID = FOURNISSEUR.ID_FR INNER JOIN
                                                            [DETAIL FACTURE] ON FACTURE.ID_FACTURE = [DETAIL FACTURE].FACTURE_ID", cnx);
                                                            DA.Fill(dt);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("" + e);
            }
            return dt;
        }
        public DataSet GetALLFACTUREdEATILbyID(int IDfac,int IDFR)
        {
            DataSet ds = new DataSet();
            using (SqlConnection cnx = new SqlConnection(Connection.stringConnection))
            {
                cnx.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter("ViewAllDetailFactureByID", cnx);
                sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
                sqlDa.SelectCommand.Parameters.AddWithValue("@IDFac", IDfac);
                sqlDa.SelectCommand.Parameters.AddWithValue("@IDfR", IDFR);
                sqlDa.Fill(ds);
            }
            return ds;
        }
        public DataTable GetALLFACTUREdEATILbyValue(string value)
        {
            DataTable dt = new DataTable();
            using (SqlConnection cnx = new SqlConnection(Connection.stringConnection))
            {
                cnx.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter("sp_ViewAllOrder", cnx);
                sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
                sqlDa.SelectCommand.Parameters.AddWithValue("@sersh", value);
                sqlDa.Fill(dt);
            }
            return dt;
        }
        public DataTable GetALLFACTUREdEATILbyDATE(DateTime datedebut, DateTime datefint)
        {
            DataTable dt = new DataTable();
            using (SqlConnection cnx = new SqlConnection(Connection.stringConnection))
            {
                cnx.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter("sp_SearchAllOrderBydate", cnx);
                sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
                sqlDa.SelectCommand.Parameters.AddWithValue("@dateDebut", datedebut);
                sqlDa.SelectCommand.Parameters.AddWithValue("@dateFin", datefint);
                sqlDa.Fill(dt);
            }
            return dt;
        }
        public DataTable ImprimerFACFR(int idFac)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection cnx = new SqlConnection(Connection.stringConnection))
                {
                    SqlDataAdapter DA = new SqlDataAdapter($@"SELECT FACTURE.NUMOP,CONVERT (varchar(10), FACTURE.[DATE SAISIE], 103) as 'DATE SAISIE' ,
                                                            FACTURE.[MODE SAISIE], 
                                                            CONVERT (varchar(10),FACTURE.[DATE ECHEANCE], 103) as 'DATE ECHEANCE',
                                                            BANQUE.NOM, FOURNISSEUR.[CODE FOURNISSEUR], 
                                                            FOURNISSEUR.NOM as NOMFR, FOURNISSEUR.IGE, FOURNISSEUR.[IF]
                                                            FROM  BANQUE INNER JOIN
                                                            FACTURE ON BANQUE.ID_BANQUE = FACTURE.BANQUE_ID INNER JOIN
                                                            FOURNISSEUR ON FACTURE.FOURNISSEUR_ID = FOURNISSEUR.ID_FR 
                                                            where FACTURE.ID_FACTURE = '{idFac}'",cnx);

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
