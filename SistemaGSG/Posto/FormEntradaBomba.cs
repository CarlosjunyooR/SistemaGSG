using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SistemaGSG.Posto
{
    public partial class FormEntradaBomba : MetroFramework.Forms.MetroForm
    {
        public FormEntradaBomba()
        {
            InitializeComponent();
        }

        private void FormEntradaBomba_Load(object sender, EventArgs e)
        {
            Decimal Bomba1, Bomba2, Resultado;
            Bomba1 = Convert.ToDecimal(txtEntrBomba1.Text.Replace(".0",""));
            Bomba2 = Convert.ToDecimal(txtSaidBomba1.Text.Replace(".0", ""));
            Resultado = Bomba1 -Bomba2;
            lblQuantSaldoBomba1.Text = Resultado.ToString();
        }
        int CLICK = 0;

        private void Ocultar()
        {
            btnChange.Enabled = false;
            txtEntrBomba1.Enabled = false;
            txtEntrBomba2.Enabled = false;
            txtEntrBomba3.Enabled = false;
            txtEntrBomba4.Enabled = false;
            txtEntrBomba5.Enabled = false;
            txtEntrBomba7.Enabled = false;
            txtEntrBomba8.Enabled = false;
            btnSearch.Visible = true;
            CLICK = 1;
        }
        private void Exibir()
        {
            btnChange.Enabled = true;
            txtEntrBomba1.Enabled = true;
            txtEntrBomba2.Enabled = true;
            txtEntrBomba3.Enabled = true;
            txtEntrBomba4.Enabled = true;
            txtEntrBomba5.Enabled = true;
            txtEntrBomba7.Enabled = true;
            txtEntrBomba8.Enabled = true;
            btnSearch.Visible = false;
            CLICK = 0;
        }


        private void btnAlterarEntradaBombas_Click(object sender, EventArgs e)
        {
            if(CLICK == 0)
            {
                Ocultar();
                MessageBox.Show("Selecione uma data! \n e clique em buscar.", "Período", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (CLICK == 1)
            {
                Exibir();
            }
        }
    }
}
