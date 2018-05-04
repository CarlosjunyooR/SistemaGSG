using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using MySql.Data.MySqlClient;

namespace SistemaGSG
{
    public partial class FormRel : MetroFramework.Forms.MetroForm
    {
        public FormRel()
        {
            InitializeComponent();
        }

        private void FormRel_Load(object sender, EventArgs e)
        {
            

        }

        private const string TEXTO = "Erro ao Conectar!\nVerifique a Conexão com\nBanco de Dados!";
        string PRIMEIRADATA;
        string SEGUNDADATA;
        string PRIMEIRADATA_HOJE;
        string SEGUNDADATA_HOJE;
        string CEAL;
        string CELPE;
        


        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja encerrar a aplicação ?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Voltar?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                frm_Main frm_Main = new frm_Main();
                frm_Main.Show();
                this.Visible = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja fazer entrada de Contas?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Ceal ceal = new Ceal();
                ceal.Show();
                this.Visible = false;
            }
        }



        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas WHERE status='PAGO' ORDER BY data ASC", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
            }
            catch (Exception)
            {
                MessageBox.Show(TEXTO);
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas WHERE status='VENCIDA' ORDER BY data ASC", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
            }
            catch (Exception)
            {
                MessageBox.Show(TEXTO);
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas WHERE status='A VENCER' ORDER BY data ASC", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
            }
            catch (Exception)
            {
                MessageBox.Show(TEXTO);
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas ORDER BY data ASC", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
            }
            catch (Exception)
            {
                MessageBox.Show(TEXTO);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter seach = new MySqlDataAdapter("SELECT * FROM contas WHERE cod='" + textBox1.Text + "' ORDER BY mes ASC", CONEX);
                DataTable seachSS = new DataTable();
                seach.Fill(seachSS);
                dataGridView1.DataSource = seachSS;
            }
            catch (Exception)
            {
                MessageBox.Show(TEXTO);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            
        }
        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter updateCEAL = new MySqlDataAdapter("UPDATE contas SET cod='" + textBox3.Text + "', nome='" + textBox2.Text + "', mes='" + textBox4.Text + "', data='" + textBox5.Text + "', valor='" + textBox6.Text + "', status='" + textBox7.Text + "' WHERE id='" + textBox10.Text + "'", CONEX);
                DataTable seachUpdate = new DataTable();
                updateCEAL.Fill(seachUpdate);
                dataGridView1.DataSource = seachUpdate;

                textBox3.Text = "";
                textBox2.Text = "";
                textBox4.Text = "";
                textBox5.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
                textBox10.Text = "";

                MessageBox.Show("Alterado com Sucesso!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception)
            {
                MessageBox.Show(TEXTO);
            }
        }

        private void dataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            /*ID*/   textBox10.Text = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
            /*COD*/   textBox3.Text = dataGridView1.SelectedRows[0].Cells[1].Value.ToString();
            /*MES*/   textBox4.Text = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
            /*VALOR*/ textBox6.Text = dataGridView1.SelectedRows[0].Cells[4].Value.ToString();
            /*FAZ*/   textBox2.Text = dataGridView1.SelectedRows[0].Cells[5].Value.ToString();
            /*DATA*/  textBox5.Text = dataGridView1.SelectedRows[0].Cells[3].Value.ToString();
            /*STATUS*/textBox7.Text = dataGridView1.SelectedRows[0].Cells[6].Value.ToString();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter updateCEAL = new MySqlDataAdapter("DELETE FROM contas WHERE id='" + textBox10.Text + "'", CONEX);
                DataTable seachUpdate = new DataTable();
                updateCEAL.Fill(seachUpdate);
                dataGridView1.DataSource = seachUpdate;

                textBox3.Text = "";
                textBox2.Text = "";
                textBox4.Text = "";
                textBox5.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
                textBox10.Text = "";

                MessageBox.Show("Excluido com Sucesso!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception)
            {
                MessageBox.Show(TEXTO);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DGVPrinter printer = new DGVPrinter();
            printer.Title = "Contas de Energia - Protocolo para Manutenção";//Cabeçalho
            printer.SubTitle = string.Format("Data: {0}", DateTime.Now.Date.ToString("dd/MM/yyyy"));
            printer.SubTitleFormatFlags = StringFormatFlags.LineLimit | StringFormatFlags.NoClip;
            printer.PageNumbers = true;
            printer.PageNumberInHeader = false;
            printer.PorportionalColumns = true;
            printer.HeaderCellAlignment = StringAlignment.Near;
            printer.Footer = "Usina Serra Grande S/A - SistemaGSG";
            printer.FooterSpacing = 15;
            printer.PrintDataGridView(dataGridView1);
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            PRIMEIRADATA = textBox9.Text;
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            SEGUNDADATA = textBox8.Text;
        }

        string SELECAO;
        string SELECAO_HOJE;

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas WHERE empresa='CEAL' ORDER BY data ASC", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
            }
            catch (Exception)
            {
                MessageBox.Show(TEXTO);
            }
            CEAL = radioButton5.Text;
            SELECAO = "CEAL";
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas WHERE empresa='CELPE' ORDER BY data ASC", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
            }
            catch (Exception)
            {
                MessageBox.Show(TEXTO);
            }
            CELPE = radioButton6.Text;
            SELECAO = "CELPE";
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter seach = new MySqlDataAdapter("SELECT * FROM contas WHERE data between '" + PRIMEIRADATA + "' AND '" + SEGUNDADATA + "' AND empresa='" + SELECAO + "' ORDER BY data ASC", CONEX);
                DataTable seachSS = new DataTable();
                seach.Fill(seachSS);
                dataGridView1.DataSource = seachSS;
            }
            catch (Exception)
            {
                MessageBox.Show(TEXTO);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {

        }

        

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            //If you want to do some formating on the footer row
            //int rowIndex = dataGridView1.Rows.GetLastRow(DataGridViewElementStates.Visible);
            //if (rowIndex <= 0)
            //{
            //    return;
            //}
            //dataGridView1.Rows[rowIndex].DefaultCellStyle.BackColor = Color.Red;
            //dataGridView1.Rows[rowIndex].DefaultCellStyle.SelectionBackColor = Color.Red;
            //dataGridView1.Rows[rowIndex].DefaultCellStyle.Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold);
        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 3)
                {
                    decimal cell1 = Convert.ToDecimal(dataGridView1.CurrentRow.Cells[2].Value);
                    decimal cell2 = Convert.ToDecimal(dataGridView1.CurrentRow.Cells[3].Value);
                    if (cell1.ToString() != "" && cell2.ToString() != "")
                    {
                        dataGridView1.CurrentRow.Cells[4].Value = cell1 * cell2;
                    }
                }
                decimal valorTotal = 0;
                string valor = "";
                if (dataGridView1.CurrentRow.Cells[4].Value != null)
                {
                    valor = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                    if (!valor.Equals(""))
                    {
                        for (int i = 0; i <= dataGridView1.RowCount - 1; i++)
                        {
                            if (dataGridView1.Rows[i].Cells[4].Value != null)
                                valorTotal += Convert.ToDecimal(dataGridView1.Rows[i].Cells[4].Value);
                        }
                        if (valorTotal == 0)
                        {
                            MessageBox.Show("Nenhum registro encontrado");
                        }
                        txtTotal.Text = valorTotal.ToString("C");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao Calcular, Verifique os Valores");
            }

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas WHERE hoje = CURDATE() ORDER BY data ASC", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
            }
            catch(Exception)
            {
                MessageBox.Show(TEXTO);
            }

        }

        MySqlConnection CONEX;

        private void boxLocal_CheckedChanged(object sender, EventArgs e)
        {
            CONEX = new MySqlConnection(@"server=localhost;database=ceal1;Uid=root;Pwd=vertrigo;");
        }

        private void boxTeste_CheckedChanged(object sender, EventArgs e)
        {
            CONEX = new MySqlConnection(@"server=10.2.1.83;database=ceal1;Uid=id889153_id885499_junior19908;Pwd=2613679cfc418651;");
        }

        private void boxCont_CheckedChanged(object sender, EventArgs e)
        {
            CONEX = new MySqlConnection(@"server=10.2.1.95;database=ceal;Uid=root;Pwd=vertrigo;");
        }

        private void definirFiltro_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter seach = new MySqlDataAdapter("SELECT * FROM contas WHERE hoje between '" + PRIMEIRADATA_HOJE + "' AND '" + SEGUNDADATA_HOJE + "' AND empresa='" + SELECAO + "' ORDER BY data ASC", CONEX);
                DataTable seachSS = new DataTable();
                seach.Fill(seachSS);
                dataGridView1.DataSource = seachSS;
            }
            catch (Exception)
            {
                MessageBox.Show(TEXTO);
            }
        }

        private void textBoxDATA1_TextChanged(object sender, EventArgs e)
        {
            PRIMEIRADATA_HOJE = textBoxDATA1.Text;
        }

        private void textBoxDATA2_TextChanged(object sender, EventArgs e)
        {
            SEGUNDADATA_HOJE = textBoxDATA2.Text;
        }

        private void button9_Click(object sender, EventArgs e)
        {

        }

        //Classes de Datas
        Int32 segundos, minutos, milisegundos;
        DateTime dataHora = DateTime.Now;
        DateTime dataHora2 = DateTime.Now.AddDays(10);

        private void button8_Click_1(object sender, EventArgs e)
        {
            try
            {
                CONEX.Open();
                MySqlCommand cmd = new MySqlCommand("SELECT CURDATE()", CONEX);
                DateTime DataServidor = Convert.ToDateTime(cmd.ExecuteScalar());
                string novadata = DataServidor.AddDays(+10).ToShortDateString();

                label12.Text = Convert.ToString(dataHora);
                label13.Text = novadata;

                dataHora = DataServidor;
                minutos = dataHora.Minute;
                segundos = dataHora.Second;
                milisegundos = dataHora.Millisecond;

                MySqlCommand command = new MySqlCommand("SELECT COUNT(*) FROM contas WHERE data BETWEEN @DataServidor AND @dataFuturo AND status='VENCIDA' OR status='A VENCER'", CONEX);

                command.Parameters.AddWithValue("@dataFuturo", dataHora2);
                command.Parameters.AddWithValue("@DataServidor", dataHora);
                command.ExecuteNonQuery();

                int qtdVencer = Convert.ToInt32(command.ExecuteScalar());
                CONEX.Close();

                if (qtdVencer > 0)
                {
                    MessageBox.Show("Você Tem " + qtdVencer + " boletos pra vencer!");
                    FormRel_Load(e, e);
                }
                else
                {
                    MessageBox.Show("Não tem boletos para vencer!");
                }
            }
            catch (Exception ex)
            {
                throw new Exception(TEXTO);
            }
        }
        Int32 PAGO;
        Int32 VENCIDA;

        public void RowsColor()
        {
            for(int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                int val = Int32.Parse(dataGridView1.Rows[i].Cells[6].Value.ToString());
                if(val > PAGO)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                }
                else if(val >= PAGO&& val == VENCIDA)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Blue;
                }
            }
        }




    }
}
