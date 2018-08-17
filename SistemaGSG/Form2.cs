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
        private const string Texto = " Duplicidade!, Este Código Único já existe no Banco de Dados.\n Por Favor, Informe outro.";
        string STATUS;
        string EMPRESA;

        MySqlCommand cmd, prompt_cmd;
        MySqlConnection CONEXAO,CONEX,cn;
        MySqlDataAdapter da;

        private void boxLocal_CheckedChanged(object sender, EventArgs e)
        {
            CONEX = new MySqlConnection(@"server=10.2.1.95;database=sistemagsg_ceal;Uid=remoto;Pwd=MbunHhYiRffEMAtl;SslMode=none;");
        }

        private void boxTeste_CheckedChanged(object sender, EventArgs e)
        {
            CONEX = new MySqlConnection(@"server=localhost;database=sistemagsg_ceal;Uid=remoto;Pwd=MbunHhYiRffEMAtl;SslMode=none;");
        }

        //CONSULTA DE DUPLICIDADE
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                CONEX.Open();//Abrir Conexão.
                MySqlCommand prompt = new MySqlCommand("SELECT COUNT(*) FROM contas_1 WHERE nf ='" + nfe.Text + "' ", CONEX);//Seleção da tabela no Banco de Dados.
                prompt.ExecuteNonQuery();//Executa o comando.
                int consultDB = Convert.ToInt32(prompt.ExecuteScalar());//Converte o resultado para números inteiros.
                CONEX.Close();//Fecha conexão.
                if (consultDB > 0)//Verifica se o resultado for maior que zero(0), a execução inicia a Menssagem de que já existe contas, caso contrario faz a inserção no Banco.
                {
                    MessageBox.Show(Texto);
                }
                else
                {
                    dbinsert();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Por favor selecione uma conexão.\nCaso já tenha selecionado e o problema ainda persistir.\nContate o Administrador do SistemaGSG!.\n'"+ex.Message+"'", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        public Ceal()
        {
            InitializeComponent();
        }

        private void dbinsert()
        {
            CONEX.Open();
            cmd = new MySqlCommand("INSERT INTO contas_multa (cod,mes,valor,empresa) VALUES ('" + textBox1.Text + "','" + mesMulta.Text + "','" + textMulta1.Text.Replace("R$ ","") + "','" + EMPRESA + "')", CONEX);
            prompt_cmd = new MySqlCommand("INSERT INTO contas_1 (cod,mes,data,valor,nome,status,hoje,empresa,pedido,migo,miro,emissao,nf,vl_icms,vl_base) VALUES ('" + textBox1.Text + "','" + textBox2.Text + "','" + this.dateTimePicker2.Text + "','" + textValor1.Text.Replace("R$ ", "") + "','" + textBox5.Text + "','" + STATUS + "', CURDATE(),'" + EMPRESA + "',NULL,NULL,NULL,'" + this.dateTimePicker1.Text + "','" + nfe.Text + "',?icms,'"+ txtBase.Text.Replace("R$ ", "") + "')", CONEX);
            prompt_cmd.Parameters.AddWithValue("?icms", preencherCBIcms.Text.Replace(" %", ""));
            cmd.ExecuteNonQuery();
            prompt_cmd.ExecuteNonQuery();
            CONEX.Close();
            //Limpar Campos apos a inserção no banco de dados.
            textBox1.Text = "";
            textBox2.Text = "";
            textBox5.Text = "";
            textValor1.Text = "";
            nfe.Text = "";
            textMulta1.Text = "";
            mesMulta.Text = "";
            txtBase.Text = "";
            MessageBox.Show("Inserido com Sucesso!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.None);
        }

        private void Ceal_Load(object sender, EventArgs e)
        {
            
        }

        private void preencherCBIcms_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                cn = new MySqlConnection(@"server=10.2.1.95;database=sistemagsg_ceal;Uid=remoto;Pwd=MbunHhYiRffEMAtl;SslMode=none;");
                cn.Open();

                MySqlCommand com = new MySqlCommand();
                com.Connection = cn;
                com.CommandText = "SELECT porcentagem FROM icms";
                MySqlDataReader dr = com.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(dr);
                preencherCBIcms.DisplayMember = "porcentagem";
                preencherCBIcms.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.None);
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

        private void textValor1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
