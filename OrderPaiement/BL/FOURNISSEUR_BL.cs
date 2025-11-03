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
    public class FOURNISSEUR_BL
    {
        public FOURNISSEUR_DAL FOURNISSEUR_DAL;
        public DataTable GetDetailPiece(string code)
        {
            FOURNISSEUR_DAL = new FOURNISSEUR_DAL();
            return FOURNISSEUR_DAL.GetDetailFournisseur(code);
        }
        public DataTable AutoCompletTextBoxFournisseur()
        {
            FOURNISSEUR_DAL = new FOURNISSEUR_DAL();
            return FOURNISSEUR_DAL.AutoCompletTextBoxFournisseur();
        }
        public string GetIDFournisseurByName(string code)
        {
            FOURNISSEUR_DAL = new FOURNISSEUR_DAL();
            return FOURNISSEUR_DAL.GetIDFournisseurByName(code);
        }
        public void Insert_Fournisseur(FOURNISSEUR fourniseur)
        {
            FOURNISSEUR_DAL = new FOURNISSEUR_DAL();
            FOURNISSEUR_DAL.Insert_Fournisseur(fourniseur);
        }

        public void Update_Fournisseur(FOURNISSEUR fourniseur)
        {
            FOURNISSEUR_DAL = new FOURNISSEUR_DAL();
            FOURNISSEUR_DAL.Update_Fournisseur(fourniseur);
        }
        public void Delete_Fournisseur(int id)
        {
            FOURNISSEUR_DAL = new FOURNISSEUR_DAL();
            FOURNISSEUR_DAL.Delete_Fournisseur(id);
        }
        public DataTable GetALLFournisseur()
        {
            FOURNISSEUR_DAL = new FOURNISSEUR_DAL();
           return FOURNISSEUR_DAL.GetALLFournisseur();
        }
        public DataTable GetFournisseurByID(int id)
        {
            FOURNISSEUR_DAL = new FOURNISSEUR_DAL();
            return FOURNISSEUR_DAL.GetFournisseurByID(id);
        }
    }
}
