using OrderPaiement.BL;
using OrderPaiement.Classe;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OrderPaiement.Reports
{
    public partial class Imp_FACTURE_FR: Form
    {
       
        int IDFAC = 0;
        FACTURE_BL FACTURE_BL = new FACTURE_BL();
        DETAILFACTURE_BL DETAILFACTURE_BL = new DETAILFACTURE_BL();
        rpt_FACTUREFOURNISSUER rpt_FACTUREFOURNISSUER = new rpt_FACTUREFOURNISSUER();
        rpt_DetailFacture rpt_DetailFacture = new rpt_DetailFacture();
        public Imp_FACTURE_FR(int IDfAC)
        {
            InitializeComponent();
            string name2 = Environment.UserName;
            string file = ($"C:\\Users\\{name2}\\Downloads\\FacturesFournissur.pdf");
            string file2 = ($"C:\\Users\\{name2}\\Downloads\\DetailFactures.pdf");
            IDFAC = IDfAC;


            DataTable datatable = new DataTable();
            datatable = FACTURE_BL.ImprimerFACFR(IDfAC);
            int count = datatable.Rows.Count;
            DataTableReader reader = datatable.CreateDataReader();
            DataSet ds = new DataSet();
            ds = new OrderPaiment();
            ds.Tables["FACTURE_FR"].Load(reader);
            rpt_FACTUREFOURNISSUER.SetDataSource(ds);
            crystalReportViewer1.ReportSource = rpt_FACTUREFOURNISSUER;

            DataTable dtDetail = new DataTable();
            dtDetail = DETAILFACTURE_BL.ImprimerDetailFacture(IDfAC);
            DataTableReader reader2 = dtDetail.CreateDataReader();
            DataSet ds2 = new DataSet();
            ds2 = new OrderPaiment();
            ds2.Tables["DetailFacture"].Load(reader2);
            rpt_FACTUREFOURNISSUER.Subreports[0].SetDataSource(ds2);

            rpt_FACTUREFOURNISSUER.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, file);
            System.Diagnostics.Process.Start(file);
            crystalReportViewer1.Refresh();

        }


    }
}
