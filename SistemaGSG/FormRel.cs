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

        string PRIMEIRADATA;
        string SEGUNDADATA;
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

        MySqlConnection con = new MySqlConnection(@"server=localhost;database=ceal1;Uid=root;Pwd=vertrigo;");


        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

            MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas WHERE status='PAGO' ORDER BY data ASC", con);
            DataTable SS = new DataTable();
            ADAP.Fill(SS);
            dataGridView1.DataSource = SS;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas WHERE status='VENCIDA' ORDER BY data ASC", con);
            DataTable SS = new DataTable();
            ADAP.Fill(SS);
            dataGridView1.DataSource = SS;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas WHERE status='A VENCER' ORDER BY data ASC", con);
            DataTable SS = new DataTable();
            ADAP.Fill(SS);
            dataGridView1.DataSource = SS;
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas ORDER BY data ASC", con);
            DataTable SS = new DataTable();
            ADAP.Fill(SS);
            dataGridView1.DataSource = SS;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            MySqlDataAdapter seach = new MySqlDataAdapter("SELECT * FROM contas WHERE cod='" + textBox1.Text + "' ORDER BY mes ASC", con);
            DataTable seachSS = new DataTable();
            seach.Fill(seachSS);
            dataGridView1.DataSource = seachSS;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            
        }
        private void button5_Click(object sender, EventArgs e)
        {
            MySqlDataAdapter updateCEAL = new MySqlDataAdapter("UPDATE contas SET cod='" + textBox3.Text + "', nome='" + textBox2.Text + "', mes='" + textBox4.Text + "', data='" + textBox5.Text + "', valor='" + textBox6.Text + "', status='" + textBox7.Text + "' WHERE id='" + textBox10.Text + "'", con);
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
            MySqlDataAdapter updateCEAL = new MySqlDataAdapter("DELETE FROM contas WHERE id='" + textBox10.Text + "'", con);
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

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas WHERE empresa='CEAL' ORDER BY data ASC", con);
            DataTable SS = new DataTable();
            ADAP.Fill(SS);
            dataGridView1.DataSource = SS;

            CEAL = radioButton5.Text;

            SELECAO = "CEAL";

        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas WHERE empresa='CELPE' ORDER BY data ASC", con);
            DataTable SS = new DataTable();
            ADAP.Fill(SS);
            dataGridView1.DataSource = SS;

            CELPE = radioButton6.Text;

            SELECAO = "CELPE";
        }

        private void button7_Click(object sender, EventArgs e)
        {
            MySqlDataAdapter seach = new MySqlDataAdapter("SELECT * FROM contas WHERE data between '" + PRIMEIRADATA + "' AND '" + SEGUNDADATA + "' AND empresa='" + SELECAO + "' ORDER BY data ASC", con);
            DataTable seachSS = new DataTable();
            seach.Fill(seachSS);
            dataGridView1.DataSource = seachSS;
        }

        private void button8_Click(object sender, EventArgs e)
        {

        }

        
        Int32 segundos, minutos, milisegundos;

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
            MySqlConnection conn = new MySqlConnection(@"server=localhost;database=ceal1;Uid=root;Pwd=vertrigo;");
            string query = String.Format("SELECT * FROM contas");

            MySqlCommand comma = new MySqlCommand(query,conn);
            MySqlDataAdapter adapter = new MySqlDataAdapter(comma);

            DataTable data = new DataTable();
            adapter.Fill(data);
            dataGridView1.DataSource = data;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM contas WHERE hoje = CURDATE() ORDER BY data ASC", con);
            DataTable SS = new DataTable();
            ADAP.Fill(SS);
            dataGridView1.DataSource = SS;

        }

        DateTime dataHora;


        private void button8_Click_1(object sender, EventArgs e)
        {
            try
            {
                MySqlConnection con = new MySqlConnection(@"server=localhost;database=ceal1;Uid=root;Pwd=vertrigo;");
                con.Open();
                MySqlCommand cmd = new MySqlCommand("SELECT CURDATE()", con);
                DateTime DataServidor = Convert.ToDateTime(cmd.ExecuteScalar());
                string novadata = DataServidor.AddDays(+10).ToShortDateString();

                label12.Text = Convert.ToString(dataHora);
                label13.Text = novadata; 

                dataHora = DataServidor;
                minutos = dataHora.Minute;
                segundos = dataHora.Second;
                milisegundos = dataHora.Millisecond;

                MySqlCommand command = new MySqlCommand("SELECT COUNT(*) FROM contas WHERE data BETWEEN @DataServidor AND @dataFuturo", con);

                command.Parameters.AddWithValue("@dataFuturo", novadata);
                command.Parameters.AddWithValue("@DataServidor", dataHora);
                command.ExecuteNonQuery();

                int qtdVencer = Convert.ToInt32(command.ExecuteScalar());
                con.Close();

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
                throw new Exception(ex.Message);
            }
        }
    }
}
