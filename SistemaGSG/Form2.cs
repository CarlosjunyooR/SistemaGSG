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
        //SqlCommand cmd ;
        // SqlConnection con;
        //SqlDataAdapter da;
        MySqlCommand cmd;
        MySqlConnection con;
        MySqlDataAdapter da;

        public Ceal()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
           
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        string STATUS;

        private void button1_Click(object sender, EventArgs e)
            {


                MySqlConnection conn = new MySqlConnection(@"server=localhost;database=ceal1;Uid=root;Pwd=vertrigo;");
                con.Open();
                cmd = new MySqlCommand("INSERT INTO CEAL1 (cod,mes,data,valor,nome,status) VALUES ('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textValor1.Text + "','" + textBox5.Text + "','" + STATUS + "')", con);
                cmd.ExecuteNonQuery();

                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox5.Text = "";
                textValor1.Text = "";

                MessageBox.Show("Inserido com Sucesso!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.None);
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

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
