using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Data.OleDb;
using MySql.Data.MySqlClient;


namespace SistemaGSG
{
    public partial class Ceal : MetroFramework.Forms.MetroForm
    {
        private const string Texto = " Duplicidade!, Este Código Único já existe no Banco de Dados.\n Por Favor\n Informe outro.";

        MySqlCommand cmd;
        MySqlConnection CONEXAO,CONEX;
        MySqlDataAdapter da;


        public Ceal()
        {
            InitializeComponent();
        }
        string STATUS;
        string EMPRESA;

        private void dbinsert()
        {
            CONEX.Open();
            cmd = new MySqlCommand("INSERT INTO contas (cod,mes,data,valor,nome,status,hoje,empresa) VALUES ('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textValor1.Text.Replace("R$ ", "") + "','" + textBox5.Text + "','" + STATUS + "', CURDATE(),'" + EMPRESA + "')", CONEX);
            cmd.ExecuteNonQuery();
            CONEX.Close();
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox5.Text = "";
            textValor1.Text = "";

            MessageBox.Show("Inserido com Sucesso!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.None);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                CONEX.Open();
                MySqlCommand prompt = new MySqlCommand("SELECT COUNT(*) FROM contas WHERE cod='" + textBox1.Text + "' AND mes='" + textBox2.Text + "'", CONEX);
                prompt.ExecuteNonQuery();
                int consultDB = Convert.ToInt32(prompt.ExecuteScalar());
                CONEX.Close();

                if (consultDB > 0)
                {
                    MessageBox.Show(Texto);
                }
                else
                {
                    dbinsert();
                }
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                throw new Exception(ex.Message);
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja encerrar a aplicação ?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Voltar?","Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                frm_Main back = new frm_Main();
                back.Show();
                this.Visible = false;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            STATUS = "PAGO";
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            STATUS = "VENCIDA";
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            STATUS = "A VENCER";
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            EMPRESA = "CEAL";
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            EMPRESA = "CELPE";
        }

        private void boxLocal_CheckedChanged(object sender, EventArgs e)
        {
            CONEX = new MySqlConnection(@"server=localhost;database=ceal1;Uid=root;Pwd=vertrigo;");
        }

        private void boxTeste_CheckedChanged(object sender, EventArgs e)
        {
            CONEX = new MySqlConnection(@"server=10.2.1.83;database=ceal1;Uid=root;Pwd=vertrigo;");
        }

        private void boxCont_CheckedChanged(object sender, EventArgs e)
        {
            CONEX = new MySqlConnection(@"server=10.2.1.95;database=ceal;Uid=id889153_id885499_junior19908;Pwd=2613679cfc418651;");
        }
    }
}
