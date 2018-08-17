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
using System.Diagnostics;

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
            if (MessageBox.Show("Deseja fazer entrada de contas_1?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Ceal ceal = new Ceal();
                ceal.Show();
                this.Visible = false;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter seach = new MySqlDataAdapter("SELECT * FROM contas_1 WHERE cod='" + textBox1.Text + "' ORDER BY mes ASC", CONEX);
                DataTable seachSS = new DataTable();
                seach.Fill(seachSS);
                dataGridView1.DataSource = seachSS;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
                MySqlDataAdapter updateCEAL = new MySqlDataAdapter("UPDATE contas_1 SET cod='" + textBox3.Text + "', nome='" + textBox2.Text + "', mes='" + textBox4.Text + "', data='" + this.dateTimePicker.Text + "', valor='" + textBox6.Text + "', status='" + textBox7.Text + "' WHERE id='" + textBox10.Text + "'", CONEX);
                DataTable seachUpdate = new DataTable();
                updateCEAL.Fill(seachUpdate);
                dataGridView1.DataSource = seachUpdate;

                textBox3.Text = "";
                textBox2.Text = "";
                textBox4.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
                textBox10.Text = "";

                MessageBox.Show("Alterado com Sucesso!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            /*ID*/      textBox10.Text = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
            /*COD*/     textBox3.Text = dataGridView1.SelectedRows[0].Cells[1].Value.ToString();
            /*MES*/     textBox4.Text = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
            /*VALOR*/   textBox6.Text = dataGridView1.SelectedRows[0].Cells[4].Value.ToString();
            /*FAZ*/     textBox2.Text = dataGridView1.SelectedRows[0].Cells[5].Value.ToString();
            /*DATA*/    this.dateTimePicker.Text = dataGridView1.SelectedRows[0].Cells[3].Value.ToString();
           /*STATUS*/   textBox7.Text = dataGridView1.SelectedRows[0].Cells[6].Value.ToString();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter updateCEAL = new MySqlDataAdapter("DELETE FROM contas_1 WHERE id='" + textBox10.Text + "'", CONEX);
                DataTable seachUpdate = new DataTable();
                updateCEAL.Fill(seachUpdate);
                dataGridView1.DataSource = seachUpdate;

                textBox3.Text = "";
                textBox2.Text = "";
                textBox4.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
                textBox10.Text = "";
                dateTimePicker.Text = "";

                MessageBox.Show("Excluido com Sucesso!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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


        string SELECAO;
        string SELECAO_HOJE;

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas_1 WHERE empresa='CEAL' ORDER BY data ASC", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            CEAL = radioButton5.Text;
            SELECAO = "CEAL";
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas_1 WHERE empresa='CELPE' ORDER BY data ASC", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            CELPE = radioButton6.Text;
            SELECAO = "CELPE";
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter seach = new MySqlDataAdapter("SELECT * FROM contas_1 WHERE data between '" + this.dateTimePicker2.Text + "' AND '" + this.dateTimePicker3.Text + "' AND empresa='" + SELECAO + "' ORDER BY data ASC", CONEX);
                DataTable seachSS = new DataTable();
                seach.Fill(seachSS);
                dataGridView1.DataSource = seachSS;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button8_Click(object sender, EventArgs e)
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
           // dataGridView1.Rows[rowIndex].DefaultCellStyle.Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold);
        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas_1 WHERE hoje = CURDATE() ORDER BY data ASC", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        MySqlConnection CONEX;

        private void boxLocal_CheckedChanged(object sender, EventArgs e)
        {
            CONEX = new MySqlConnection(@"server=10.2.1.95;database=sistemagsg_ceal;Uid=remoto;Pwd=MbunHhYiRffEMAtl;SslMode=none;");
        }

        private void boxTeste_CheckedChanged(object sender, EventArgs e)
        {
            CONEX = new MySqlConnection(@"server=localhost;database=sistemagsg_ceal;Uid=root;Pwd=vertrigo;SslMode=none;");
        }

        private void definirFiltro_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter seach = new MySqlDataAdapter("SELECT * FROM contas_1 WHERE hoje between '" + this.dateTimePicker4.Text + "' AND '" + this.dateTimePicker5.Text + "' AND empresa='" + SELECAO + "' ORDER BY data ASC", CONEX);
                DataTable seachSS = new DataTable();
                seach.Fill(seachSS);
                dataGridView1.DataSource = seachSS;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {

        }

        //Classes de Datas
        Int32 segundos, minutos, milisegundos;
        DateTime dataHora = DateTime.Now;

        private void txtMes_TextChanged(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter seach = new MySqlDataAdapter("SELECT * FROM contas_1 WHERE mes='" + txtMes.Text + "' ORDER BY id ASC", CONEX);
                DataTable seachSS = new DataTable();
                seach.Fill(seachSS);
                dataGridView1.DataSource = seachSS;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void txtTotal_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas_1 WHERE status='PAGO' AND data between '" + this.dateTimePicker2.Text + "' AND '" + this.dateTimePicker3.Text + "' ORDER BY data ASC", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas_1 WHERE status='VENCIDA' ORDER BY data ASC", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas_1 WHERE status='A VENCER' ORDER BY data ASC", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            frmProtocolo protColo = new frmProtocolo();
            protColo.Show();
            this.Visible = false; 
        }

        Bitmap bmp;

        private void btnPrintview_Click(object sender, EventArgs e)
        {
            int height = dataGridView1.Height;
            dataGridView1.Height = dataGridView1.RowCount * dataGridView1.RowTemplate.Height * 2;
            bmp = new Bitmap(dataGridView1.Width, dataGridView1.Height);
            dataGridView1.DrawToBitmap(bmp, new Rectangle(0, 0, dataGridView1.Width, dataGridView1.Height));
            dataGridView1.Height = height;
            printPreviewDialog1.ShowDialog();
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawImage(bmp, 0, 0);
        }

        private void dataGridView1_CellClick_1(object sender, DataGridViewCellEventArgs e)
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
                if (dataGridView1.CurrentRow.Cells[5].Value != null)
                {
                    valor = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                    if (!valor.Equals(""))
                    {
                        for (int i = 0; i <= dataGridView1.RowCount - 1; i++)
                        {
                            if (dataGridView1.Rows[i].Cells[5].Value != null)
                                valorTotal += Convert.ToDecimal(dataGridView1.Rows[i].Cells[5].Value);
                        }
                        if (valorTotal == 0)
                        {
                            MessageBox.Show("Nenhum registro encontrado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        txtTotal.Text = valorTotal.ToString("C");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao Calcular, Verifique os Valores\n'" + ex.Message + "'", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas_1 ORDER BY data ASC", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            try
            {
                CONEX.Open();
                MySqlCommand cmd = new MySqlCommand("SELECT CURDATE()", CONEX);
                DateTime DataServidor = Convert.ToDateTime(cmd.ExecuteScalar());
                string novadata = DataServidor.AddDays(+10).ToShortDateString();

                label12.Text = Convert.ToString(DataServidor);
                label13.Text = novadata;

                dataHora = DataServidor;
                minutos = dataHora.Minute;
                segundos = dataHora.Second;
                milisegundos = dataHora.Millisecond;

                MySqlCommand command = new MySqlCommand("SELECT COUNT(*) FROM contas_1 WHERE data BETWEEN @DataServidor AND @dataFuturo AND status='VENCIDA' OR status='A VENCER'", CONEX);
                MySqlCommand update_3 = new MySqlCommand("UPDATE contas_1 SET status='VENCIDA' WHERE data < CURDATE() AND status ='A VENCER'", CONEX);
                
                command.Parameters.AddWithValue("@dataFuturo", novadata);
                command.Parameters.AddWithValue("@DataServidor", dataHora);
                command.ExecuteNonQuery();
                update_3.ExecuteNonQuery();

                MySqlDataAdapter seach = new MySqlDataAdapter("SELECT * FROM contas_1 ORDER BY id DESC", CONEX);
                DataTable seachSS = new DataTable();
                seach.Fill(seachSS);
                dataGridView1.DataSource = seachSS;

                int qtdVencer = Convert.ToInt32(command.ExecuteScalar());
                CONEX.Close();
                if (qtdVencer > 0)
                {
                    MessageBox.Show("Você Tem " + qtdVencer + " boletos pra vencer!", "Aviso!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    MessageBox.Show("Não tem boletos para vencer!", "Aviso!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                    MessageBox.Show(ex.Message);
            }
        }
    }
}
