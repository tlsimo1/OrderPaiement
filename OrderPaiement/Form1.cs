using OrderPaiement.BL;
using OrderPaiement.Classe;
using OrderPaiement.Reports;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;


namespace OrderPaiement
{
    public partial class Form1 : Form
    {
        FOURNISSEUR_BL FOURNISSEUR_BL = new FOURNISSEUR_BL();
        BANQUE_BL BANQUE_BL = new BANQUE_BL();
        DETAILFACTURE_BL DETAILFACTURE_BL = new DETAILFACTURE_BL();
        FACTURE_BL FACTUREBL = new FACTURE_BL();
        System.Data.DataTable DTAUTOCOMPLET = new System.Data.DataTable();
        AutoCompleteStringCollection coll = new AutoCompleteStringCollection();
        DateTimePicker dtp = new DateTimePicker();
        DateTimePicker dtp2 = new DateTimePicker();
        System.Drawing.Rectangle _Rectangle;
        System.Data.DataTable dtAllBanque = new System.Data.DataTable();
        int NUMORDER = 0;
        int LASTIDFAC = 0;
        int IDFAC = 0;
        int IDFR = 0;
        int IDFr = 0;

        string currentdatetime = DateTime.Now.ToString("yyyyMMddHHmmss");
        string LogFolder = @"D:\Files\Logs";
        string connectionString = @"Server=DESKTOP-EKJ1P64\SQL2019;Database=Test;Integrated Security=True;";

        string filePath = @"D:\Files\Email.XLSX";
        public Form1()
        {
            InitializeComponent();
            try
            {
                DTAUTOCOMPLET = FOURNISSEUR_BL.AutoCompletTextBoxFournisseur();
                dgv_detail.Rows[0].Cells["dgv_dteDATEFac"].Value = DateTime.Today.ToShortDateString();
                dgv_detail.Controls.Add(dtp);
                dtp.Visible = false;
                dtp.Format = DateTimePickerFormat.Custom;
                dgv_detail.Rows[0].Cells["dgv_BANQUE"].Value = DateTime.Today.ToShortDateString();
                dgv_detail.Controls.Add(dtp2);
                dtp2.Visible = false;
                dtp2.Format = DateTimePickerFormat.Custom;
                dgv_fournisseur.DataSource = FOURNISSEUR_BL.GetALLFournisseur();
            }
            catch (Exception)
            {
                return;
            }
            dtp.TextChanged += Dtp_TextChanged;
            dtp2.TextChanged += Dtp2_TextChanged;
        }

        private void Dtp2_TextChanged(object sender, EventArgs e)
        {
            dgv_detail.CurrentCell.Value = dtp2.Text.ToString();
        }
        private void Dtp_TextChanged(object sender, EventArgs e)
        {
            dgv_detail.CurrentCell.Value = dtp.Text.ToString();

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                Auto();
                cmb_valeur.Text = "-select-";
                cmb_moderemise.Text = "-select-";
                NUMORDER = Convert.ToInt16(DETAILFACTURE_BL.GetLastID());

                if (NUMORDER == 0)
                    NUMORDER = NUMORDER + 1;

                txt_numop.Text = NUMORDER.ToString();
                RemplirComboBanque();
                dgv_listeorderpaiement.DataSource = FACTUREBL.GetALLFACTURE();
                dgv_impression.DataSource = FACTUREBL.GetALLFACTUREDETAIL();
                dgv_fournisseur.DataSource = FOURNISSEUR_BL.GetALLFournisseur();

            }
            catch (Exception)
            {

                return;
            }
        }
        private void txt_codefournisseur_TextChanged(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataTable dtDetailFournisseur = new System.Data.DataTable();
                dtDetailFournisseur = FOURNISSEUR_BL.GetDetailPiece(txt_codefournisseur.Text);
                foreach (DataRow row in dtDetailFournisseur.Rows)
                {
                    txt_nomfournisseur.Text = row["NOM"].ToString();
                    txt_icg.Text = row["IGE"].ToString();
                    txt_if.Text = row["IF"].ToString();
                }
            }
            catch (Exception)
            {

                return;
            }

        }
        public void Auto()
        {
            try
            {
                for (int i = 0; i < DTAUTOCOMPLET.Rows.Count; i++)
                {
                    coll.Add(DTAUTOCOMPLET.Rows[i]["CODE FOURNISSEUR"].ToString());
                }
                txt_codefournisseur.AutoCompleteMode = AutoCompleteMode.Suggest;
                txt_codefournisseur.AutoCompleteSource = AutoCompleteSource.CustomSource;
                txt_codefournisseur.AutoCompleteCustomSource = coll;
            }
            catch (Exception)
            {

                return;
            }
        }
        private void Form1_Scroll(object sender, ScrollEventArgs e)
        {
            dtp.Visible = false;
        }
        private void dgv_detail_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                //if (dgv_listeorderpaiement.CurrentRow.Index != -1)
                //{
                switch (dgv_detail.Columns[e.ColumnIndex].Name)
                {
                    case "dgv_dteDATEFac":
                        _Rectangle = dgv_detail.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
                        dtp.Size = new Size(_Rectangle.Width, _Rectangle.Height);
                        dtp.Location = new System.Drawing.Point(_Rectangle.X, _Rectangle.Y);
                        dtp.Visible = true;
                        break;
                    case "dgv_BANQUE":
                        _Rectangle = dgv_detail.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
                        dtp.Size = new Size(_Rectangle.Width, _Rectangle.Height);
                        dtp.Location = new System.Drawing.Point(_Rectangle.X, _Rectangle.Y);
                        dtp.Visible = true;
                        break;
                }
                //}
            }
            catch (Exception)
            {

                return;
            }

        }
        void RemplirComboBanque()
        {
            try
            {
                dtAllBanque = BANQUE_BL.GetAllBanque();
                DataRow dr = dtAllBanque.NewRow();
                dr[0] = 0;
                dr[1] = "-Select-";
                dtAllBanque.Rows.InsertAt(dr, 0);
                cmb_banque.ValueMember = "ID_BANQUE";
                cmb_banque.DisplayMember = "NOM";
                cmb_banque.DataSource = dtAllBanque;
            }
            catch (Exception)
            {

                return;
            }

        }
        void InsertFactureDetail()
        {
            try
            {
                FACTURE facture = new FACTURE();
                facture.DateSaisie = Convert.ToDateTime(dte_datesaisie.Text);
                if (cmb_moderemise.SelectedIndex == 0)
                {
                    cmb_moderemise.BackColor = System.Drawing.Color.Red;
                    Helpers.IsValid = false;
                }
                else
                {
                    cmb_moderemise.BackColor = SystemColors.Window;
                    facture.Mode = cmb_moderemise.Text == "-Select-" ? "" : cmb_moderemise.Text;
                    Helpers.IsValid = true;
                }
                facture.DateEcheance = Convert.ToDateTime(dte_dateecheance.Text);
                if (cmb_banque.SelectedIndex == 0)
                {
                    cmb_banque.BackColor = System.Drawing.Color.Red;
                    Helpers.IsValid = false;
                }
                else
                {
                    cmb_banque.BackColor = SystemColors.Window;
                    facture.Banque_ID = (int)cmb_banque.SelectedValue;
                    Helpers.IsValid = true;
                }
                if (cmb_valeur.SelectedIndex == 0)
                {
                    cmb_valeur.BackColor = System.Drawing.Color.Red;
                    Helpers.IsValid = false;
                }
                else
                {
                    cmb_valeur.BackColor = SystemColors.Window;
                    facture.Valeur = cmb_valeur.Text == "-Select-" ? "" : cmb_valeur.Text;
                    Helpers.IsValid = true;
                }
                facture.NUMOP = NUMORDER.ToString();
                if (txt_codefournisseur.Text == "")
                {
                    txt_codefournisseur.BackColor = System.Drawing.Color.Red;
                    Helpers.IsValid = false;
                }
                else
                {
                    txt_codefournisseur.BackColor = SystemColors.Window;
                    bool successfullyParsed = int.TryParse(FOURNISSEUR_BL.GetIDFournisseurByName(txt_codefournisseur.Text), out int result);
                    if (successfullyParsed)
                    {
                        facture.Fournisseur_ID = Convert.ToInt16(FOURNISSEUR_BL.GetIDFournisseurByName(txt_codefournisseur.Text));
                        Helpers.IsValid = true;
                    }

                    else
                    {
                        MessageBox.Show("Code fournisseur non valide");
                        Helpers.IsValid = false;
                        txt_codefournisseur.BackColor = System.Drawing.Color.Red;
                        return;
                    }

                }
                LASTIDFAC = Convert.ToInt16(DETAILFACTURE_BL.GetLastID()) + 1;
                FACTUREBL.INSERT_FAC(facture);

                List<DETAILFACTURE> ListDetailFacture = new List<DETAILFACTURE>();
                foreach (DataGridViewRow dgvRow in dgv_detail.Rows)
                {
                    if (dgvRow.IsNewRow) break;
                    else
                    {
                        DETAILFACTURE detailFacture = new DETAILFACTURE();
                        //detailFacture.ID_DETAIL = (int)(dgvRow.Cells["dgv_IDDETAIL"].Value);
                        detailFacture.DateFacture = dgvRow.Cells["dgv_dteDATEFac"].Value == null ? DateTime.Now : Convert.ToDateTime(dgvRow.Cells["dgv_dteDATEFac"].Value);
                        detailFacture.NumFacture = dgvRow.Cells["dgv_dteNUMFACTURE"].Value.ToString();
                        detailFacture.Objet = dgvRow.Cells["dgv_OBJET"].Value.ToString();
                        detailFacture.MontantHT = Convert.ToDouble(dgvRow.Cells["dgv_MONTANTHT"].Value);
                        detailFacture.TVA = Convert.ToDouble(dgvRow.Cells["dgv_TVA"].Value);
                        detailFacture.MontantTTC = Convert.ToDouble(dgvRow.Cells["dgv_MONTANT_TTC"].Value);
                        detailFacture.TauxTAV = dgvRow.Cells["dgv_TAUXTVA"].Value.ToString();
                        detailFacture.Ligne = dgvRow.Cells["dgv_LIGNE"].Value.ToString();
                        detailFacture.DateBanque = dgvRow.Cells["dgv_BANQUE"].Value == null ? DateTime.Now : Convert.ToDateTime(dgvRow.Cells["dgv_BANQUE"].Value);
                        detailFacture.FactureID = LASTIDFAC;
                        ListDetailFacture.Add(detailFacture);
                    }

                }
                DETAILFACTURE_BL.Insert_DeatilFacture(ListDetailFacture);
                MessageBox.Show("order paiement créé avec succès");
            }
            catch (Exception)
            {

                return;
            }

        }
        void UpdateFacture()
        {
            try
            {
                FACTURE Updatefacture = new FACTURE();
                Updatefacture.ID_FACTURE = Convert.ToInt16(dgv_listeorderpaiement.CurrentRow.Cells[0].Value.ToString());
                Updatefacture.DateSaisie = Convert.ToDateTime(dte_datesaisie.Text);
                if (cmb_moderemise.SelectedIndex == 0)
                {
                    cmb_moderemise.BackColor = System.Drawing.Color.Red;
                    Helpers.IsValid = false;
                }
                else
                {
                    cmb_moderemise.BackColor = SystemColors.Window;
                    Updatefacture.Mode = cmb_moderemise.Text == "-Select-" ? "" : cmb_moderemise.Text;
                    Helpers.IsValid = true;
                }
                Updatefacture.DateEcheance = Convert.ToDateTime(dte_dateecheance.Text);
                if (cmb_banque.SelectedIndex == 0)
                {
                    cmb_banque.BackColor = System.Drawing.Color.Red;
                    Helpers.IsValid = false;
                }
                else
                {
                    cmb_banque.BackColor = SystemColors.Window;
                    Updatefacture.Banque_ID = (int)cmb_banque.SelectedValue;
                    Helpers.IsValid = true;
                }
                if (cmb_valeur.SelectedIndex == 0)
                {
                    cmb_valeur.BackColor = System.Drawing.Color.Red;
                    Helpers.IsValid = false;
                }
                else
                {
                    cmb_valeur.BackColor = SystemColors.Window;
                    Updatefacture.Valeur = cmb_valeur.Text == "-Select-" ? "" : cmb_valeur.Text;
                    Helpers.IsValid = true;
                }
                Updatefacture.NUMOP = NUMORDER.ToString();
                if (txt_codefournisseur.Text == "")
                {
                    txt_codefournisseur.BackColor = System.Drawing.Color.Red;
                    Helpers.IsValid = false;
                }
                else
                {
                    txt_codefournisseur.BackColor = SystemColors.Window;
                    bool successfullyParsed = int.TryParse(FOURNISSEUR_BL.GetIDFournisseurByName(txt_codefournisseur.Text), out int result);
                    if (successfullyParsed)
                    {
                        Updatefacture.Fournisseur_ID = Convert.ToInt16(FOURNISSEUR_BL.GetIDFournisseurByName(txt_codefournisseur.Text));
                        Helpers.IsValid = true;
                    }

                    else
                    {
                        MessageBox.Show("Code fournisseur non valide");
                        Helpers.IsValid = false;
                        txt_codefournisseur.BackColor = System.Drawing.Color.Red;
                        return;
                    }

                }
                // LASTIDFAC = Convert.ToInt16(DETAILFACTURE_BL.GetLastID()) + 1;
                FACTUREBL.UPDATE_FAC(Updatefacture);

                List<DETAILFACTURE> ListDetailFacture = new List<DETAILFACTURE>();
                foreach (DataGridViewRow dgvRow in dgv_detail.Rows)
                {
                    if (dgvRow.IsNewRow) break;
                    else
                    {
                        DETAILFACTURE detailFacture = new DETAILFACTURE();
                        detailFacture.ID_DETAIL = Convert.ToInt32(dgvRow.Cells["dgv_IDDETAIL"].Value.ToString());
                        detailFacture.DateFacture = Convert.ToDateTime(dgvRow.Cells["dgv_dteDATEFac"].Value) == null ? DateTime.Now : Convert.ToDateTime(dgvRow.Cells["dgv_dteDATEFac"].Value);
                        detailFacture.NumFacture = dgvRow.Cells["dgv_dteNUMFACTURE"].Value.ToString();
                        detailFacture.Objet = dgvRow.Cells["dgv_OBJET"].Value.ToString();
                        detailFacture.MontantHT = Convert.ToDouble(dgvRow.Cells["dgv_MONTANTHT"].Value);
                        detailFacture.TVA = Convert.ToDouble(dgvRow.Cells["dgv_TVA"].Value);
                        detailFacture.MontantTTC = Convert.ToDouble(dgvRow.Cells["dgv_MONTANT_TTC"].Value);
                        detailFacture.TauxTAV = dgvRow.Cells["dgv_TAUXTVA"].Value.ToString();
                        detailFacture.Ligne = dgvRow.Cells["dgv_LIGNE"].Value.ToString();
                        detailFacture.DateBanque = Convert.ToDateTime(dgvRow.Cells["dgv_BANQUE"].Value) == null ? DateTime.Now : Convert.ToDateTime(dgvRow.Cells["dgv_BANQUE"].Value);
                        detailFacture.FactureID = Updatefacture.ID_FACTURE;
                        ListDetailFacture.Add(detailFacture);
                    }

                }
                DETAILFACTURE_BL.Update_DeatilFacture(ListDetailFacture);
                MessageBox.Show("order paiement modifié avec succès");
            }
            catch (Exception)
            {

                return;
            }
        }
        private void btn_save_Click(object sender, EventArgs e)
        {
            try
            {
                if (IsValide())
                {
                    if (btn_save.Text == "SAVE")
                    {
                        if (float.TryParse(dgv_detail.CurrentRow.Cells[4].Value.ToString(), out float outParse))
                        {
                            InsertFactureDetail();
                            RestCahmps();
                            dgv_impression.DataSource = FACTUREBL.GetALLFACTUREDETAIL();
                            dgv_listeorderpaiement.DataSource = FACTUREBL.GetALLFACTURE();
                        }
                        else
                        {
                            Helpers.IsValid = false;
                            MessageBox.Show("le format de la chaine d'entré est incorrect ");
                            dgv_detail.CurrentRow.Cells[4].Style.BackColor = System.Drawing.Color.Red;
                        }
                            
                    }
                    else
                    {
                        if (float.TryParse(dgv_detail.CurrentRow.Cells[4].Value.ToString(), out float outParse))
                        {
                            UpdateFacture();
                            //RestCahmps();
                            dgv_impression.DataSource = FACTUREBL.GetALLFACTUREDETAIL();
                        }
                        else
                        {
                            Helpers.IsValid = false;
                            MessageBox.Show("le format de la chaine d'entré est incorrect ");
                            dgv_detail.CurrentRow.Cells[4].Style.BackColor = System.Drawing.Color.Red;
                        }
                        
                    }
                    dgv_listeorderpaiement.DataSource = FACTUREBL.GetALLFACTURE();
                    dgv_listeorderpaiement.DataSource = FACTUREBL.GetALLFACTURE();

                }

                else
                    MessageBox.Show("Order Paiment non Valide .Merci de remplir tous les champs ");
            }
            catch (Exception)
            {

                return;
            }

        }
        private void dgv_detail_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            try
            {

                if (dgv_detail.CurrentRow.Cells[8].Value != null)
                {
                    if(dgv_detail.CurrentRow.Cells[4].Value!=null)
                    {
                        if (float.TryParse(dgv_detail.CurrentRow.Cells[4].Value.ToString(), out float outParse))
                        {
                            dgv_detail.CurrentRow.Cells[4].Style.BackColor = SystemColors.Window;
                            string taxuTTC = dgv_detail.CurrentRow.Cells[5].Value.ToString();
                            switch (taxuTTC)
                            {
                                case "10%":
                                    dgv_detail.CurrentRow.Cells[6].Value = double.Parse(dgv_detail.CurrentRow.Cells[4].Value.ToString()) * (0.1);
                                    break;
                                case "12%":
                                    dgv_detail.CurrentRow.Cells[6].Value = double.Parse(dgv_detail.CurrentRow.Cells[4].Value.ToString()) * (0.12);
                                    break;
                                case "15%":
                                    dgv_detail.CurrentRow.Cells[6].Value = double.Parse(dgv_detail.CurrentRow.Cells[4].Value.ToString()) * (0.15);
                                    break;
                                case "18%":
                                    dgv_detail.CurrentRow.Cells[6].Value = double.Parse(dgv_detail.CurrentRow.Cells[4].Value.ToString()) * (0.18);
                                    break;
                                case "20%":
                                    dgv_detail.CurrentRow.Cells[6].Value = double.Parse(dgv_detail.CurrentRow.Cells[4].Value.ToString()) * (0.2);
                                    break;
                                default:
                                    dgv_detail.CurrentRow.Cells[6].Value = double.Parse(dgv_detail.CurrentRow.Cells[4].Value.ToString()) * (0.1);
                                    break;
                            }
                            dgv_detail.CurrentRow.Cells[7].Value = double.Parse(dgv_detail.CurrentRow.Cells[4].Value.ToString()) + double.Parse(dgv_detail.CurrentRow.Cells[6].Value.ToString());
                            Helpers.IsValid = true;
                        }
                        //else
                        //{
                        //    Helpers.IsValid = false;
                        //    MessageBox.Show("le format de la chaine d'entré est incorrect ");
                        //    dgv_detail.CurrentRow.Cells[4].Style.BackColor = System.Drawing.Color.Red;


                        //}
                    }
                    

                }
                CalculateSumOfDatagrid();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }

        }
        private void dgv_listeorderpaiement_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                if (dgv_listeorderpaiement.CurrentRow.Index != -1)
                {
                    DataGridViewRow _dgvCurrentRow = dgv_listeorderpaiement.CurrentRow;
                    IDFAC = Convert.ToInt32(_dgvCurrentRow.Cells[0].Value);
                    IDFR = Convert.ToInt32(_dgvCurrentRow.Cells[1].Value);
                    DataSet ds = FACTUREBL.GetALLFACTUREDeatil(IDFAC, IDFR);
                    //Master - Fill
                    DataRow dr = ds.Tables[0].Rows[0];

                    txt_numop.Text = dr["NUMOP"].ToString();
                    dte_datesaisie.Text = dr["DATE SAISIE"].ToString();
                    cmb_moderemise.Text = dr["MODE SAISIE"].ToString();
                    dte_dateecheance.Text = dr["DATE ECHEANCE"].ToString();
                    cmb_banque.SelectedValue = Convert.ToInt16(dr["BANQUE_ID"].ToString());
                    cmb_valeur.Text = dr["VALEUR"].ToString();

                    DataRow dr1 = ds.Tables[1].Rows[0];
                    dgv_detail.DataSource = ds.Tables[1];
                    dgv_detail.AutoGenerateColumns = false;



                    DataRow dr2 = ds.Tables[2].Rows[0];
                    txt_codefournisseur.Text = dr2["CODE FOURNISSEUR"].ToString();
                    txt_nomfournisseur.Text = dr2["NOM"].ToString();
                    txt_icg.Text = dr2["IGE"].ToString();
                    txt_if.Text = dr2["IF"].ToString();

                    btn_delete.Enabled = true;
                    btn_save.Text = "Update";
                    tbl_ajouterfournisseur.SelectedIndex = 0;


                }
                CalculateSumOfDatagrid();
            }
            catch (Exception)
            {

                return;
            }

        }
        private void btn_imprimer_Click(object sender, EventArgs e)
        {
            try
            {
                int ImpLastIDFac = 0;
                ImpLastIDFac = Convert.ToInt16(DETAILFACTURE_BL.GetLastID());
                if (IDFAC == 0)
                {
                    Imp_FACTURE_FR Imp_FACTURE_FR = new Imp_FACTURE_FR(ImpLastIDFac);
                }
                else { Imp_FACTURE_FR Imp_FACTURE_FR = new Imp_FACTURE_FR(IDFAC); }
            }
            catch (Exception)
            {

                return;
            }

        }
        bool IsValide()
        {
            try
            {
                if (dgv_detail.CurrentRow.Cells[4].Value != null)
                {
                    if (dgv_detail.Rows.Count > 1 && IsDigitsOnly(dgv_detail.CurrentRow.Cells[4].Value.ToString() == null ? "0" : dgv_detail.CurrentRow.Cells[4].Value.ToString())
                && cmb_moderemise.Text != "-Select-" && cmb_banque.Text != "-Select-"
                && cmb_valeur.Text != "-Select-" && txt_codefournisseur.Text != "-Select-")
                    {
                        Helpers.IsValid = true;
                    }
                    //else if (dgv_detail.Rows.Count > 1 || dgv_detail.CurrentRow.Cells[4].Style.BackColor.Name == Color.Red.Name)
                    //{
                    //    Helpers.IsValid = true;
                    //}
                    else
                    {
                        Helpers.IsValid = false;
                    }
                }

            }
            catch (Exception)
            {
                return false;
                ;
            }
            return Helpers.IsValid;
        }
        bool IsDigitsOnly(string str)
        {
            try
            {
                foreach (char c in str)
                {
                    if (c < '0' || c > '9')
                        return false;
                }

                return true;
            }
            catch (Exception)
            {

                return false;
            }

        }
        void CalculateSumOfDatagrid()
        {
            try
            {
                double sumHT = 0;
                double sumTVA = 0;
                double sumTTC = 0;
                for (int i = 0; i < dgv_detail.Rows.Count; ++i)
                {
                    sumHT += Convert.ToDouble(dgv_detail.Rows[i].Cells[4].Value);
                    sumTVA += Convert.ToDouble(dgv_detail.Rows[i].Cells[6].Value);
                    sumTTC += Convert.ToDouble(dgv_detail.Rows[i].Cells[7].Value);
                }
                lbl_ht.Text = sumHT.ToString();
                lbl_tva.Text = sumTVA.ToString();
                lbl_ttc.Text = sumTTC.ToString();
            }
            catch (Exception)
            {

                return;
            }

        }
        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }
        void RestCahmps()
        {
            LASTIDFAC = 0;
            IDFAC = 0;
            IDFR = 0;
            txt_numop.Text = (NUMORDER + 1).ToString();
            txt_nomfournisseur.Text = "";
            txt_icg.Text = "";
            txt_if.Text = "";
            cmb_banque.SelectedIndex = 0;
            cmb_moderemise.SelectedIndex = 0;
            cmb_valeur.SelectedIndex = 0;
            if (dgv_detail.DataSource == null)
                dgv_detail.Rows.Clear();
            else
            {
                dgv_detail.DataSource = (dgv_detail.DataSource as System.Data.DataTable).Clone();
                dgv_detail.Refresh();
            }
            btn_save.Text = "SAVE";
            btn_delete.Enabled = false;
        }
        private void btn_reset_Click(object sender, EventArgs e)
        {
            LASTIDFAC = 0;
            IDFAC = 0;
            IDFR = 0;
            txt_numop.Text = (NUMORDER).ToString();
            txt_nomfournisseur.Text = "";
            txt_icg.Text = "";
            txt_if.Text = "";
            txt_codefournisseur.Text = "";
            cmb_banque.SelectedIndex = 0;
            cmb_moderemise.SelectedIndex = 0;
            cmb_valeur.SelectedIndex = 0;
            if (dgv_detail.DataSource == null)
                dgv_detail.Rows.Clear();
            else
            {
                dgv_detail.DataSource = (dgv_detail.DataSource as System.Data.DataTable).Clone();
                dgv_detail.Refresh();
            }
            btn_save.Text = "SAVE";
            btn_delete.Enabled = false;
        }
        private void btn_delete_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Are you Sure to Delete this Record ?", "Order Paiement", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    FACTUREBL.Delete_FactureAndDetail(Convert.ToInt16(dgv_detail.Rows[0].Cells["dgv_FACTURE_ID"].Value));
                    RestCahmps();
                    dgv_listeorderpaiement.DataSource = FACTUREBL.GetALLFACTURE();
                    MessageBox.Show("Deleted Successfully");
                    dgv_impression.DataSource = FACTUREBL.GetALLFACTUREDETAIL();
                    dgv_listeorderpaiement.DataSource = FACTUREBL.GetALLFACTURE();
                }
            }
            catch (Exception)
            {

                return;
            }

        }
        private void dgv_detail_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            try
            {
                DataGridViewRow dgvRow = dgv_detail.CurrentRow;
                if (dgvRow.Cells["dgv_IDDETAIL"].Value != DBNull.Value)
                {
                    if (MessageBox.Show("Are You Sure to Delete this Record ?", "Order de Paiement", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        DETAILFACTURE_BL.DELETE_DetailFacture(Convert.ToInt16(dgvRow.Cells["dgv_IDDETAIL"].Value));
                        MessageBox.Show("Deleted Row Successfully");
                    }
                    else
                        e.Cancel = true;
                }
            }
            catch (Exception)
            {

                return;
            }

        }
        private void btn_expexcel_Click(object sender, EventArgs e)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(OrderPaiement.Classe.Connection.stringConnection))
                {
                    using (SqlDataAdapter sda = new SqlDataAdapter("select * from FOURNISSEUR", con))
                    {
                        using (System.Data.DataTable dt = new System.Data.DataTable())
                        {
                            sda.Fill(dt);
                            string folderPath = "D:\\Excel\\";
                            if (!Directory.Exists(folderPath))
                            {
                                Directory.CreateDirectory(folderPath);
                            }
                            using (XLWorkbook wb = new XLWorkbook())
                            {
                                wb.Worksheets.Add(dt, "Customers");
                                wb.SaveAs(folderPath + "Customers.xlsx");
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                using (StreamWriter sw = File.CreateText(LogFolder + "\\" + "ErrorLog_" + currentdatetime + ".log"))
                {
                    sw.WriteLine(exception.ToString());
                }
            }
        }
        private void txt_recherche_TextChanged(object sender, EventArgs e)
        {
            try
            {
                dgv_impression.DataSource = FACTUREBL.GetALLFACTUREdEATILbyValue(txt_recherche.Text);
            }
            catch (Exception)
            {

                return;
            }
        }
        private void btn_excelfile_Click(object sender, EventArgs e)
        {
            double totalHT = 0;
            double totalTTC = 0;
            double totalTVA = 0;
            double totallIGNE = 0;
            for (int i = 0; i < dgv_impression.Rows.Count - 1; i++)
            {
                totalHT += Convert.ToDouble(dgv_impression.Rows[i].Cells[12].Value);
                totalTTC += Convert.ToDouble(dgv_impression.Rows[i].Cells[14].Value);
                totalTVA += Convert.ToDouble(dgv_impression.Rows[i].Cells[15].Value);
                totallIGNE += Convert.ToDouble(dgv_impression.Rows[i].Cells[17].Value);
            }

            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // Create a new Excel workbook
                        System.Data.DataTable dt = DataGridView_To_Datatable(dgv_impression);
                        dt.exportToExcel(sfd.FileName);
                        MessageBox.Show("Data is exported!");
                        using (XLWorkbook workbook = new XLWorkbook(sfd.FileName))
                        {
                            var hoja = workbook.Worksheets.Worksheet(1);

                            string valuecelltotalHT1 = "Q" + (dgv_impression.Rows.Count + 4).ToString();
                            string valuecelltotalHT2 = "R" + (dgv_impression.Rows.Count + 4).ToString();
                            hoja.Cell(valuecelltotalHT1).Value = " SUM MONTANT HT";
                            hoja.Cell(valuecelltotalHT2).Value = totalHT;

                            hoja.Cell(valuecelltotalHT2).Style.Font.FontSize = 10;
                            hoja.Cell(valuecelltotalHT2).Style.Font.FontName = "Arial";

                            hoja.Cell(valuecelltotalHT2).Style
                            .Border.SetTopBorder(XLBorderStyleValues.Medium)
                            .Border.SetRightBorder(XLBorderStyleValues.Medium)
                            .Border.SetBottomBorder(XLBorderStyleValues.Medium)
                            .Border.SetLeftBorder(XLBorderStyleValues.Medium);
                            hoja.Cell(valuecelltotalHT1).Style
                            .Border.SetTopBorder(XLBorderStyleValues.Medium)
                            .Border.SetRightBorder(XLBorderStyleValues.Medium)
                            .Border.SetBottomBorder(XLBorderStyleValues.Medium)
                            .Border.SetLeftBorder(XLBorderStyleValues.Medium);


                            string valuecelltotalTTC1 = "Q" + (dgv_impression.Rows.Count + 5).ToString();
                            string valuecelltotalTTC2 = "R" + (dgv_impression.Rows.Count + 5).ToString();
                            hoja.Cell(valuecelltotalTTC1).Value = " SUM MONTANT TTC";
                            hoja.Cell(valuecelltotalTTC2).Value = totalTTC;

                            hoja.Cell(valuecelltotalTTC2).Style.Font.FontSize = 10;
                            hoja.Cell(valuecelltotalTTC2).Style.Font.FontName = "Arial";

                            hoja.Cell(valuecelltotalTTC2).Style
                            .Border.SetTopBorder(XLBorderStyleValues.Medium)
                            .Border.SetRightBorder(XLBorderStyleValues.Medium)
                            .Border.SetBottomBorder(XLBorderStyleValues.Medium)
                            .Border.SetLeftBorder(XLBorderStyleValues.Medium);
                            hoja.Cell(valuecelltotalTTC1).Style
                            .Border.SetTopBorder(XLBorderStyleValues.Medium)
                            .Border.SetRightBorder(XLBorderStyleValues.Medium)
                            .Border.SetBottomBorder(XLBorderStyleValues.Medium)
                            .Border.SetLeftBorder(XLBorderStyleValues.Medium);

                            string valuecelltotalTVA1 = "Q" + (dgv_impression.Rows.Count + 6).ToString();
                            string valuecelltotalTVA2 = "R" + (dgv_impression.Rows.Count + 6).ToString();
                            hoja.Cell(valuecelltotalTVA1).Value = " SUM MONTANT TVA";
                            hoja.Cell(valuecelltotalTVA2).Value = totalTVA;

                            hoja.Cell(valuecelltotalTVA2).Style.Font.FontSize = 10;
                            hoja.Cell(valuecelltotalTVA2).Style.Font.FontName = "Arial";
                            hoja.Cell(valuecelltotalTVA2).Style
                            .Border.SetTopBorder(XLBorderStyleValues.Medium)
                            .Border.SetRightBorder(XLBorderStyleValues.Medium)
                            .Border.SetBottomBorder(XLBorderStyleValues.Medium)
                            .Border.SetLeftBorder(XLBorderStyleValues.Medium);
                            hoja.Cell(valuecelltotalTVA1).Style
                            .Border.SetTopBorder(XLBorderStyleValues.Medium)
                            .Border.SetRightBorder(XLBorderStyleValues.Medium)
                            .Border.SetBottomBorder(XLBorderStyleValues.Medium)
                            .Border.SetLeftBorder(XLBorderStyleValues.Medium);

                            string valuecelltotallIGNE1 = "Q" + (dgv_impression.Rows.Count + 7).ToString();
                            string valuecelltotallIGNE2 = "R" + (dgv_impression.Rows.Count + 7).ToString();
                            hoja.Cell(valuecelltotallIGNE1).Value = " SUM LIGNE";
                            hoja.Cell(valuecelltotallIGNE2).Value = totallIGNE;

                            hoja.Cell(valuecelltotallIGNE2).Style.Font.FontSize = 10;
                            hoja.Cell(valuecelltotallIGNE2).Style.Font.FontName = "Arial";
                            hoja.Cell(valuecelltotallIGNE2).Style
                            .Border.SetTopBorder(XLBorderStyleValues.Medium)
                            .Border.SetRightBorder(XLBorderStyleValues.Medium)
                            .Border.SetBottomBorder(XLBorderStyleValues.Medium)
                            .Border.SetLeftBorder(XLBorderStyleValues.Medium);
                            hoja.Cell(valuecelltotallIGNE1).Style
                            .Border.SetTopBorder(XLBorderStyleValues.Medium)
                            .Border.SetRightBorder(XLBorderStyleValues.Medium)
                            .Border.SetBottomBorder(XLBorderStyleValues.Medium)
                            .Border.SetLeftBorder(XLBorderStyleValues.Medium);

                            hoja.Cell(valuecelltotalHT1).Style.Fill.SetBackgroundColor(XLColor.Yellow);
                            hoja.Cell(valuecelltotalTTC1).Style.Fill.SetBackgroundColor(XLColor.Yellow);
                            hoja.Cell(valuecelltotalTVA1).Style.Fill.SetBackgroundColor(XLColor.Yellow);
                            hoja.Cell(valuecelltotallIGNE1).Style.Fill.SetBackgroundColor(XLColor.Yellow);

                            hoja.Cell(valuecelltotalHT2).Style.Fill.SetBackgroundColor(XLColor.Pink);
                            hoja.Cell(valuecelltotalTTC2).Style.Fill.SetBackgroundColor(XLColor.Pink);
                            hoja.Cell(valuecelltotalTVA2).Style.Fill.SetBackgroundColor(XLColor.Pink);
                            hoja.Cell(valuecelltotallIGNE2).Style.Fill.SetBackgroundColor(XLColor.Pink);

                            hoja.Columns("Q:Q").Width = 20;
                            for (int i = 1; i < dgv_impression.Rows.Count+1; i++)
                            {
                                hoja.Rows($"{i.ToString()}:{i.ToString()}").Height = 17;
                            }
                            workbook.SaveAs(sfd.FileName);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        public static System.Data.DataTable DataGridView_To_Datatable(DataGridView dg)
        {
            System.Data.DataTable ExportDataTable = new System.Data.DataTable();
            foreach (DataGridViewColumn col in dg.Columns)
            {
                ExportDataTable.Columns.Add(col.Name);
            }
            foreach (DataGridViewRow row in dg.Rows)
            {
                DataRow dRow = ExportDataTable.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    string name = cell.Value.ToString().Replace("\r\n","");
                    dRow[cell.ColumnIndex] = name;
                }
                ExportDataTable.Rows.Add(dRow);
            }
            return ExportDataTable;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataTable dtimpression = new System.Data.DataTable();
                dtimpression = FACTUREBL.GetALLFACTUREdEATILbyDATE(Convert.ToDateTime(dte_datedebut.Text), Convert.ToDateTime(dte_datefin.Text));
                dgv_impression.DataSource = dtimpression;
            }
            catch (Exception)
            {
                return;
            }
        }
        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void btn_newfournisseur_Click(object sender, EventArgs e)
        {
            try
            {
                FOURNISSEUR fournisseur = new FOURNISSEUR();
                fournisseur.CodeFournisseur = txt_newcode.Text;
                fournisseur.NomFournisseur = txt_newnomfournisseur.Text;
                fournisseur.IGE = txt_newice.Text;
                fournisseur.IF = txt_newif.Text;
                FOURNISSEUR_BL.Insert_Fournisseur(fournisseur);
                dgv_fournisseur.DataSource = FOURNISSEUR_BL.GetALLFournisseur();
                MessageBox.Show("Fournisseur créé avec succès");
                ResetFournisseur();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void btn_newdeletefournissueur_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Are you Sure to Delete this Record ?", "Fournisseur", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    FOURNISSEUR_BL.Delete_Fournisseur(IDFr);
                    ResetFournisseur();
                    dgv_fournisseur.DataSource = FOURNISSEUR_BL.GetALLFournisseur();
                    MessageBox.Show("Deleted Successfully");
                    dgv_impression.DataSource = FACTUREBL.GetALLFACTUREDETAIL();
                    dgv_listeorderpaiement.DataSource = FACTUREBL.GetALLFACTURE();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        void ResetFournisseur()
        {
            txt_newcode.Text = "";
            txt_newnomfournisseur.Text = "";
            txt_newice.Text = "";
            txt_newif.Text = "";
        }

        private void btn_updatenewfournisseur_Click(object sender, EventArgs e)
        {
            try
            {
                //int id = Convert.ToInt16(dgv_fournisseur.Rows[0].Cells["ID_FR"].Value);
                FOURNISSEUR fournisseur = new FOURNISSEUR();
                fournisseur.ID_FR = IDFr;
                fournisseur.CodeFournisseur = txt_newcode.Text;
                fournisseur.NomFournisseur = txt_newnomfournisseur.Text;
                fournisseur.IGE = txt_newice.Text;
                fournisseur.IF = txt_newif.Text;
                FOURNISSEUR_BL.Update_Fournisseur(fournisseur);
                dgv_fournisseur.DataSource = FOURNISSEUR_BL.GetALLFournisseur();
                MessageBox.Show("Fournisseur modifié avec succès");
                ResetFournisseur();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void dgv_fournisseur_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (dgv_fournisseur.CurrentRow.Index != -1)
                {
                    DataGridViewRow _dgvCurrentRow = dgv_fournisseur.CurrentRow;
                    IDFr = Convert.ToInt32(_dgvCurrentRow.Cells[0].Value);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt = FOURNISSEUR_BL.GetFournisseurByID(IDFr);
                    FOURNISSEUR fournisseur = new FOURNISSEUR();
                    foreach (DataRow dr in dt.Rows)
                    {
                       txt_newcode.Text=dr["CODE FOURNISSEUR"].ToString();
                       txt_newnomfournisseur.Text= dr["NOM"].ToString();
                       txt_newice.Text = dr["IGE"].ToString();
                       txt_newif.Text= dr["IF"].ToString();
                    }
                   
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
    }
}

