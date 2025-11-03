using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrderPaiement.Classe
{
    public class DETAILFACTURE
    {
        public int ID_DETAIL { get; set; }
        public DateTime? DateFacture { get; set; }
        public string NumFacture { get; set; }
        public string Objet { get; set; }
        public double MontantHT { get; set; }
        public double TVA { get; set; }
        public double MontantTTC{ get; set; }
        public string TauxTAV { get; set; }
        public string Ligne { get; set; }
        public int FactureID { get; set; }
        public DateTime? DateBanque{ get; set; }

    }
}
