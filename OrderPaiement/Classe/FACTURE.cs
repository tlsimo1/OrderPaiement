using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrderPaiement.Classe
{
    public class FACTURE
    {
        public int ID_FACTURE { get; set; }
        public string NUMOP { get; set; }
        public DateTime DateSaisie { get; set; }
        public string Mode { get; set; }
        public DateTime DateEcheance { get; set; }
        public string Valeur { get; set; }
        public int Fournisseur_ID { get; set; }
        public int Banque_ID { get; set; }
    }
}
