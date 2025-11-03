using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OrderPaiement.Classe;

namespace OrderPaiement.DataAccess
{
    public class FOURNISSEUR_DAL
    {
        public DataTable GetDetailFournisseur(string code)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection cnx = new SqlConnection(Connection.stringConnection))
                {
                    SqlDataAdapter da = new SqlDataAdapter($@"select * from FOURNISSEUR where [CODE FOURNISSEUR]  like '%{code}%' ", cnx);

                    da.Fill(dt);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("" + e);
            }
            return dt;
        }
        public DataTable AutoCompletTextBoxFournisseur()
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection connection = new SqlConnection(Connection.stringConnection))
                {
                    SqlDataAdapter da = new SqlDataAdapter(@"select * from FOURNISSEUR", connection);
                    da.Fill(dt);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("" + e);
            }
            return dt;
        }
        public string GetIDFournisseurByName(string code)
        {
            
            string id = "";
            try
            {
                using (SqlConnection cnx = new SqlConnection(Connection.stringConnection))
                {
                    SqlCommand cmd = new SqlCommand($@"SELECT ID_FR from FOURNISSEUR where [CODE FOURNISSEUR] like '{code}'", cnx);
                    cnx.Open();
                    var returnValue = cmd.ExecuteScalar();
                    if (returnValue != null)
                        id = returnValue.ToString();

                }
            }
            catch (Exception e)
            {
                Console.WriteLine("" + e);
            }
            return id;
        }
        public void Insert_Fournisseur(FOURNISSEUR fourniseur)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Connection.stringConnection))
                {
                    connection.Open();
                    SqlCommand com = new SqlCommand($@"INSERT INTO [dbo].[FOURNISSEUR]
                                                    ([CODE FOURNISSEUR]
                                                    ,[NOM]
                                                    ,[IGE]
                                                    ,[IF])
                                                     VALUES
                                                     ('{fourniseur.CodeFournisseur}',
                                                       '{fourniseur.NomFournisseur}',
                                                       '{fourniseur.IGE}',
                                                       '{fourniseur.IF}')", connection);

                    com.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("" + e);
            }

        }
        public void Update_Fournisseur(FOURNISSEUR fourniseur)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Connection.stringConnection))
                {
                    connection.Open();
                    SqlCommand com = new SqlCommand($@"UPDATE [dbo].[FOURNISSEUR]
                                                     SET [CODE FOURNISSEUR] = '{fourniseur.CodeFournisseur}'
                                                    ,[NOM] = '{fourniseur.NomFournisseur}'
                                                    ,[IGE] = '{fourniseur.IGE}'
                                                    ,[IF] = '{fourniseur.IF}'
                                                     WHERE ID_FR ={fourniseur.ID_FR}", connection);

                    com.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("" + e);
            }

        }
        public void Delete_Fournisseur(int IDFournissur)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Connection.stringConnection))
                {
                    connection.Open();
                    SqlCommand com = new SqlCommand($@"DELETE FROM [dbo].[FOURNISSEUR]
                                                  WHERE  ID_FR ={IDFournissur}", connection);
                    com.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("" + e);
            }

        }
        public DataTable GetALLFournisseur()
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection cnx = new SqlConnection(Connection.stringConnection))
                {
                    SqlDataAdapter da = new SqlDataAdapter($@"select * from FOURNISSEUR ", cnx);
                    da.Fill(dt);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("" + e);
            }
            return dt;
        }

        public DataTable GetFournisseurByID(int id)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection cnx = new SqlConnection(Connection.stringConnection))
                {
                    SqlDataAdapter da = new SqlDataAdapter($@"select * from FOURNISSEUR where ID_FR = {id} ", cnx);

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
