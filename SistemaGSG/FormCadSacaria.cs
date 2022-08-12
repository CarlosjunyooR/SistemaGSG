using MySql.Data.MySqlClient;
using System;
using System.Windows.Forms;

namespace SistemaGSG
{
    public partial class FormCadSacaria : MetroFramework.Forms.MetroForm
    {
        public FormCadSacaria()
        {
            InitializeComponent();
        }
        MySqlCommand cmd;

        private void btnInsert_Click(object sender, EventArgs e)
        {
            try
            {
                cmd = new MySqlCommand("INSERT INTO `tb_acesso` (`id`, `col_acesso`, `col_data_acesso`, `col_hora_acesso`, `col_cliente`, `col_tipo_sac`, `col_quant`, `col_ov`, `col_sacaria_inicio`, `col_sacaria_fim`, `col_safra`, `col_transport`, `col_motorista`, `col_placa_1`, `col_placa_2`, `col_placa_3`, `col_placa_4`, `Obs`) VALUES (NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '" + txtInicSacaria.Text + "', '" + txtFimSacaria.Text + "', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)", ConexaoDados.GetConnectionFaturameto());
                cmd.ExecuteNonQuery();

                MessageBox.Show("Salvo com Sucesso!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);

                ConexaoDados.GetConnectionFaturameto().Close();

                txtInicSacaria.Text = "";
                txtFimSacaria.Text = "";
            }
            catch (MySqlException Err)
            {
                MessageBox.Show(Err.Message);
            }
            catch (Exception Erro)
            {
                MessageBox.Show(Erro.Message);
            }
        }
    }
}
