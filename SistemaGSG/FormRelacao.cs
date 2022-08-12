using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Windows.Forms;

namespace SistemaGSG
{
    public partial class FormRelacao : MetroFramework.Forms.MetroForm
    {
        public FormRelacao()
        {
            InitializeComponent();
            LoadDataGrid();
        }
        private void LoadDataGrid()
        {
            MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT TipoDoc.col_desc_NFe, TbChave.*,EtapaNfe.col_desc FROM `tb_chave` AS TbChave LEFT JOIN tb_tipo_nfe AS TipoDoc ON TbChave.tpNF=TipoDoc.col_id LEFT JOIN tb_etapa_sap AS EtapaNfe ON TbChave.ACTION_REQU=EtapaNfe.ID ORDER BY `emisao` DESC", ConexaoDados.GetConnectionXML());
            DataTable SS = new DataTable();
            ADAP.Fill(SS);
            dataGridView1.DataSource = SS;
            ConexaoDados.GetConnectionXML().Close();
            dataGridView1.Columns["Column6"].DefaultCellStyle.Format = "C2";
        }
        private void LoadDataGridNaoLanc()
        {
            MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT TipoDoc.col_desc_NFe, TbChave.*,EtapaNfe.col_desc FROM `tb_chave` AS TbChave LEFT JOIN tb_tipo_nfe AS TipoDoc ON TbChave.tpNF=TipoDoc.col_id LEFT JOIN tb_etapa_sap AS EtapaNfe ON TbChave.ACTION_REQU=EtapaNfe.ID WHERE status='NÃO LANÇADA' ORDER BY `emisao` DESC", ConexaoDados.GetConnectionXML());
            DataTable SS = new DataTable();
            ADAP.Fill(SS);
            dataGridView1.DataSource = SS;
            ConexaoDados.GetConnectionXML().Close();
            dataGridView1.Columns["Column6"].DefaultCellStyle.Format = "C2";
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            var url = dataGridView1.SelectedRows[0].Cells["Column13"].Value;
            if (string.IsNullOrEmpty(url.ToString()))
            {

            }
            else
            {
                try
                {
                    var AbrirNavegador = new Navegador(url.ToString());
                    AbrirNavegador.Show();
                }
                catch (Exception Err)
                {
                    MessageBox.Show(Err.Message);
                }
            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                LoadDataGrid();
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                LoadDataGridNaoLanc();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                LoadDataGrid();
            }
            if (radioButton2.Checked)
            {
                LoadDataGridNaoLanc();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja encerrar a aplicação ?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja voltar ao Menu Inicial?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                frm_Main AbrirForm = new frm_Main();
                AbrirForm.Show();
                Close();
            }
        }
    }
}
