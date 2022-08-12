using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Windows.Forms;

namespace SistemaGSG
{
    public partial class FormDesconto : MetroFramework.Forms.MetroForm
    {
        public string ButaoRadio { get; private set; }

        public FormDesconto()
        {
            InitializeComponent();

        }
        private void btnSave_Click_1(object sender, EventArgs e)
        {
            dataDesconto.Format = DateTimePickerFormat.Custom;
            dataDesconto.CustomFormat = "yyyy-MM-dd";
            try
            {
                if (rdSafra2020.Checked)
                {
                    ButaoRadio = "2020/2021";
                }
                if (rdSafra2021.Checked)
                {
                    ButaoRadio = "2021/2022";
                }
                MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_debitos (col_Fornecedor, col_Data, col_Desconto, col_Quantidade, col_Unidade, col_ValorDiv, col_dataImport, safra) VALUES ('" + txtCodFornecedor.Text + "', '" + this.dataDesconto.Text + "', '" + cmbDesc.Text + "', '" + txtQuant.Text + "', '" + txtUnidade.Text + "', '" + txtDivida.Text + "', NOW(), '" + ButaoRadio + "')", ConexaoDados.GetConnectionFornecedor());
                prompt_cmd.ExecuteNonQuery();
                ConexaoDados.GetConnectionFornecedor().Close();
                gridViewLoad();
                txtCodFornecedor.Text = "";
                txtDivida.Text = "";
                txtQuant.Text = "";
                txtUnidade.Text = "";
                cmbDesc.Text = "";
                txtSafra.Text = "";
            }
            catch (MySqlException)
            {
                var CadastroForm = new FormCadastroFornecedor(txtCodFornecedor.Text);
                CadastroForm.Show();
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
            }
            dataDesconto.Format = DateTimePickerFormat.Custom;
            dataDesconto.CustomFormat = "dd/MM/yyyy";
        }
        private void gridViewLoad()
        {
            try
            {
                if (rdSafra2020.Checked)
                {
                    ButaoRadio = "2020/2021";
                }
                if (rdSafra2021.Checked)
                {
                    ButaoRadio = "2021/2022";
                }
                if (rdSafra2020.Checked)
                {
                    MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT DebT.*,TBFUN.col_Nome,TIPODESC.col_Descricao FROM `tb_debitos` AS DebT LEFT JOIN tb_fornecedor AS TBFUN ON DebT.col_Fornecedor=TBFUN.id_CodFornecedor LEFT JOIN tb_tipodesconto AS TIPODESC ON DebT.col_Desconto=TIPODESC.id_desc WHERE safra='" + ButaoRadio + "'", ConexaoDados.GetConnectionFornecedor());
                    DataTable SS = new DataTable();
                    ADAP.Fill(SS);
                    dtDebitos.DataSource = SS;
                }
                if (rdSafra2021.Checked)
                {
                    MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT DebT.*,TBFUN.col_Nome,TIPODESC.col_Descricao FROM `tb_debitos` AS DebT LEFT JOIN tb_fornecedor AS TBFUN ON DebT.col_Fornecedor=TBFUN.id_CodFornecedor LEFT JOIN tb_tipodesconto AS TIPODESC ON DebT.col_Desconto=TIPODESC.id_desc WHERE safra='" + ButaoRadio + "'", ConexaoDados.GetConnectionFornecedor());
                    DataTable SS = new DataTable();
                    ADAP.Fill(SS);
                    dtDebitos.DataSource = SS;
                }
                ConexaoDados.GetConnectionFornecedor().Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void Detalhamento()
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_detalhamentodebitos WHERE col_Fornec='" + textBoxCodigo.Text + "' AND col_TipoDesc='" + textBoxDesc.Text + "' AND id_Debito='" + txtID.Text + "' ", ConexaoDados.GetConnectionFornecedor());
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dtDetalhe.DataSource = SS;
                ConexaoDados.GetConnectionFornecedor().Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void dtDebitos_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtID.Text = dtDebitos.SelectedRows[0].Cells[0].Value.ToString();
            textBoxCodigo.Text = dtDebitos.SelectedRows[0].Cells[1].Value.ToString();
            textBoxDesc.Text = dtDebitos.SelectedRows[0].Cells[3].Value.ToString();
            txtSafraOu.Text = dtDebitos.SelectedRows[0].Cells[8].Value.ToString();
            txtValorDTdesconto.Text = dtDebitos.SelectedRows[0].Cells["VALORDIV"].Value.ToString();

            Detalhamento();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Voltar?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                frm_Main back = new frm_Main();
                back.Show();
                Close();
            }
        }

        private void dtDetalhe_DoubleClick(object sender, EventArgs e)
        {
            var conexaoform = new FormDetalhe(textBoxCodigo.Text, textBoxDesc.Text, txtSafraOu.Text, txtID.Text);
            conexaoform.Show();
        }

        private void txtCodFornecedor_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                dataDesconto.Focus();
            }
        }

        private void dataDesconto_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                cmbDesc.Focus();
            }
        }

        private void cmbDesc_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtQuant.Focus();
            }
        }

        private void txtUnidade_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtDivida.Focus();
            }
        }

        private void txtQuant_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtUnidade.Focus();
            }
        }

        private void txtDivida_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtSafra.Focus();
            }
        }

        private void txtSafra_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnSave.Focus();
            }
        }

        private void DescontoForm_Load(object sender, EventArgs e)
        {
            gridViewLoad();

        }

        private void rdSafra2020_Click(object sender, EventArgs e)
        {
            MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT DebT.*,TBFUN.col_Nome,TIPODESC.col_Descricao FROM `tb_debitos` AS DebT LEFT JOIN tb_fornecedor AS TBFUN ON DebT.col_Fornecedor=TBFUN.id_CodFornecedor LEFT JOIN tb_tipodesconto AS TIPODESC ON DebT.col_Desconto=TIPODESC.id_desc WHERE safra='2020/2021'", ConexaoDados.GetConnectionFornecedor());
            DataTable SS = new DataTable();
            ADAP.Fill(SS);
            dtDebitos.DataSource = SS;
        }
        private void rdSafra2021_Click(object sender, EventArgs e)
        {
            MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT DebT.*,TBFUN.col_Nome,TIPODESC.col_Descricao FROM `tb_debitos` AS DebT LEFT JOIN tb_fornecedor AS TBFUN ON DebT.col_Fornecedor=TBFUN.id_CodFornecedor LEFT JOIN tb_tipodesconto AS TIPODESC ON DebT.col_Desconto=TIPODESC.id_desc WHERE safra='2021/2022'", ConexaoDados.GetConnectionFornecedor());
            DataTable SS = new DataTable();
            ADAP.Fill(SS);
            dtDebitos.DataSource = SS;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Deseja Excluir?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Error) == DialogResult.Yes)
                {
                    MySqlCommand prompt_cmd = new MySqlCommand("DELETE FROM `tb_debitos` WHERE id='" + txtID.Text + "'", ConexaoDados.GetConnectionFornecedor());
                    prompt_cmd.ExecuteNonQuery();
                    ConexaoDados.GetConnectionFornecedor().Close();
                    txtID.Text = "";
                    gridViewLoad();
                    Detalhamento();
                }
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
            }
        }

        private void rdSafra2020_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void dtDetalhe_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                decimal valorTotal2 = 0;
                string valor = "";
                if (dtDetalhe.CurrentRow.Cells["Column6"].Value != null)
                {
                    valor = dtDetalhe.CurrentRow.Cells["Column6"].Value.ToString();
                    if (!valor.Equals(""))
                    {
                        for (int i = 0; i <= dtDetalhe.RowCount - 1; i++)
                        {
                            if (dtDetalhe.Rows[i].Cells["Column6"].Value != null)
                                valorTotal2 += Convert.ToDecimal(dtDetalhe.Rows[i].Cells["Column6"].Value);
                        }
                        if (valorTotal2 == 0)
                        {
                            MessageBox.Show("Nenhum registro encontrado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        txtDescontado.Text = valorTotal2.ToString("C");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao Calcular, Verifique os Valores Texto_2\n'" + ex.Message + "'", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            double vlDifTotal = Convert.ToDouble(txtValorDTdesconto.Text.Replace("R$ ", "")) - Convert.ToDouble(txtDescontado.Text.Replace("R$ ", ""));
            txtSaldo.Text = vlDifTotal.ToString("C");
        }
    }
}
