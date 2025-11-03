using DocumentFormat.OpenXml.Office2010.Excel;
using OrderPaiement.Classe;
using OrderPaiement.DataAccess;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OrderPaiement.BL
{
    public class FACTURE_BL
    {
        public FACTURE_DAL FACTURE_DAL;
        public void INSERT_FAC(FACTURE facture)
        {
            FACTURE_DAL = new FACTURE_DAL();
            FACTURE_DAL.Insert_Facture(facture);
        }
        public DataTable GetALLFACTUREDETAIL()
        {
            FACTURE_DAL = new FACTURE_DAL();
            return FACTURE_DAL.GetALLFACTUREDETAIL();
        }
        public void UPDATE_FAC(FACTURE UPDATEfacture)
        {
            FACTURE_DAL = new FACTURE_DAL();
            FACTURE_DAL.Update_Facture(UPDATEfacture);
        }
        public DataTable GetALLFACTUREdEATILbyValue(string value)
        {
            FACTURE_DAL = new FACTURE_DAL();
            return FACTURE_DAL.GetALLFACTUREdEATILbyValue(value);
        }
        public DataTable GetALLFACTUREdEATILbyDATE(DateTime datedebut, DateTime datefint)
        {
            FACTURE_DAL = new FACTURE_DAL();
            return FACTURE_DAL.GetALLFACTUREdEATILbyDATE(datedebut, datefint);
        }
        public void Delete_FactureAndDetail(int IDFAC)
        {
            FACTURE_DAL = new FACTURE_DAL();
            FACTURE_DAL.Delete_FactureAndDetail(IDFAC);
        }
        public DataTable GetALLFACTURE()
        {
            FACTURE_DAL = new FACTURE_DAL();
            return FACTURE_DAL.GetALLFACTURE();
        }
        public DataSet GetALLFACTUREDeatil(int ID, int IDFR)
        {
            FACTURE_DAL = new FACTURE_DAL();
            return FACTURE_DAL.GetALLFACTUREdEATILbyID(ID, IDFR);
        }
        public DataTable ImprimerFACFR(int idFac)
        {
            FACTURE_DAL = new FACTURE_DAL();
            return FACTURE_DAL.ImprimerFACFR(idFac);
        }
    }
}
