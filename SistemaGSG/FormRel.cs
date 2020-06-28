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
using System.Threading;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using Application = System.Windows.Forms.Application;
using Rectangle = System.Drawing.Rectangle;

namespace SistemaGSG
{
    public partial class FormRel : MetroFramework.Forms.MetroForm
    {
        public FormRel()
        {
            InitializeComponent();
        }

        struct DataParameter
        {
            public List<DataGrid> ProductList;
            public string Filename { get; private set; }
        }
        DataParameter _inputParameter;
        public string Filename { get; private set; }
        private void FormRel_Load(object sender, EventArgs e)
        {
            try
            {
                MySqlConnection cn = new MySqlConnection(@"server=usga-servidor-m;database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
                cn.Open();

                MySqlCommand com = new MySqlCommand();
                com.Connection = cn;
                com.CommandText = "SELECT mes FROM tb_mes";
                MySqlDataReader dr = com.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(dr);
                preencherCBmes.DisplayMember = "mes";
                preencherCBmes.DataSource = dt;
            }
            catch
            {
                MySqlConnection cn = new MySqlConnection(@"server=localhost;database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
                cn.Open();

                MySqlCommand com = new MySqlCommand();
                com.Connection = cn;
                com.CommandText = "SELECT mes FROM tb_mes";
                MySqlDataReader dr = com.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(dr);
                preencherCBmes.DisplayMember = "mes";
                preencherCBmes.DataSource = dt;
            }

            try
            {
                MySqlConnection cn = new MySqlConnection(@"server=usga-servidor-m;database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
                cn.Open();

                MySqlCommand com = new MySqlCommand();
                com.Connection = cn;
                com.CommandText = "SELECT ano FROM tb_mes";
                MySqlDataReader dr = com.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(dr);
                preencherCBano.DisplayMember = "ano";
                preencherCBano.DataSource = dt;
            }
            catch
            {
                MySqlConnection cn = new MySqlConnection(@"server=localhost;database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
                cn.Open();

                MySqlCommand com = new MySqlCommand();
                com.Connection = cn;
                com.CommandText = "SELECT ano FROM tb_mes";
                MySqlDataReader dr = com.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(dr);
                preencherCBano.DisplayMember = "ano";
                preencherCBano.DataSource = dt;
            }
            try
            {
                DataTable oTable = new DataTable();
                using (MySqlConnection cn = new MySqlConnection(@"server=localhost;database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;"))
                {
                    string Mysql = "SELECT * FROM tb_boleto";
                    cn.Open();
                    MySqlCommand cmd = new MySqlCommand(Mysql, cn);
                    cmd.CommandText = Mysql;
                    cmd.CommandType = CommandType.Text;
                    MySqlDataReader oDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                    oTable.Load(oDataReader);
                    dataGridView1.DataSource = oTable;
                    formataGridView();
                }
            }catch(Exception Err)
            {
                MessageBox.Show(Err.ToString());
            }
        }
        private void button14_Click(object sender, EventArgs e)
        {
            new FormPedido().Show();
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
            if (MessageBox.Show("Deseja fazer entrada de Energia?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
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
                MySqlDataAdapter seach = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE cod_unico='" + textBox1.Text + "' ORDER BY Mes_ref DESC", CONEX);
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
                MySqlDataAdapter updateCEAL = new MySqlDataAdapter("UPDATE tb_boleto SET cod_unico='" + txtCod.Text + "', vl_dif='" + txt_dif_boleto.Text.Replace(".","") + "' ,desc_item='" + txtFaz.Text + "', Mes_ref='" + txtMess.Text + "', data_venc='" + this.dtVencimento.Text + "', valor_miro='" + txtvalor.Text.Replace(".","") + "', status='" + txtstatus.Text + "', pedido='" + txtPedido.Text + "', migo='" + txtmigo.Text + "', miro='" + txtMiro.Text + "', nfe='"+txtNf.Text+"', vl_icms='"+txtICMS.Text+ "', base_calculo='" + txtVlBase.Text.Replace(".","") + "' WHERE id='" + textBox10.Text + "'", CONEX);
                DataTable seachUpdate = new DataTable();
                updateCEAL.Fill(seachUpdate);
                dataGridView1.DataSource = seachUpdate;

                txtCod.Text = "";
                txtFaz.Text = "";
                txtMess.Text = "";
                txtvalor.Text = "";
                txtstatus.Text = "";
                textBox10.Text = "";
                txtPedido.Text = "";
                txtMiro.Text = "";
                txtmigo.Text = "";
                txtNf.Text = "";
                txtVlBase.Text = "";
                txt_dif_boleto.Text = "";
                this.dtVencimento.Text = "";
                txtICMS.Text = "";

                MessageBox.Show("Alterado com Sucesso!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void dataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            textBox10.Text = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
            txtCod.Text = dataGridView1.SelectedRows[0].Cells[25].Value.ToString();
            txtMess.Text = dataGridView1.SelectedRows[0].Cells[24].Value.ToString();
            txtvalor.Text = dataGridView1.SelectedRows[0].Cells[30].Value.ToString();
            txtFaz.Text = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
            this.dtVencimento.Text = dataGridView1.SelectedRows[0].Cells[23].Value.ToString();
            txtstatus.Text = dataGridView1.SelectedRows[0].Cells[31].Value.ToString();
            txtPedido.Text = dataGridView1.SelectedRows[0].Cells[26].Value.ToString();
            txtMiro.Text = dataGridView1.SelectedRows[0].Cells[28].Value.ToString();
            txtmigo.Text = dataGridView1.SelectedRows[0].Cells[27].Value.ToString();
            txtICMS.Text = dataGridView1.SelectedRows[0].Cells[8].Value.ToString();
            txtNf.Text = dataGridView1.SelectedRows[0].Cells[19].Value.ToString();
            txtVlBase.Text = dataGridView1.SelectedRows[0].Cells[7].Value.ToString();
            txt_dif_boleto.Text = dataGridView1.SelectedRows[0].Cells[16].Value.ToString();

        }
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter updateCEAL = new MySqlDataAdapter("DELETE FROM tb_boleto WHERE id='" + textBox10.Text + "'", CONEX);
                DataTable seachUpdate = new DataTable();
                updateCEAL.Fill(seachUpdate);
                dataGridView1.DataSource = seachUpdate;

                txtCod.Text = "";
                txtFaz.Text = "";
                txtMess.Text = "";
                txtvalor.Text = "";
                txtstatus.Text = "";
                textBox10.Text = "";
                dtVencimento.Text = "";

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
        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE empresa='CEAL' ORDER BY id ASC", CONEX);
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
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE empresa='CELPE' ORDER BY id ASC", CONEX);
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
                MySqlDataAdapter seach = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE data_venc between '" + this.dateTimePicker2.Text.Replace("/", "-").ToString() + "' AND '" + this.dateTimePicker3.Text.Replace("/", "-").ToString() + "' AND empresa='" + SELECAO + "' ORDER BY id ASC", CONEX);
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

        }
        private void label9_Click(object sender, EventArgs e)
        {

        }
        MySqlConnection CONEX;
        private void definirFiltro_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter seach = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE now_date between '" + this.dateTimePicker4.Text + "' AND '" + this.dateTimePicker5.Text + "' AND empresa='" + SELECAO + "' ORDER BY id ASC", CONEX);
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

        private void txtTotal_TextChanged(object sender, EventArgs e)
        {
            
        }
        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE status='PAGO' ORDER BY Mes_ref ASC", CONEX);
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
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE status='VENCIDA' ORDER BY id ASC", CONEX);
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
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE status='A VENCER' ORDER BY id ASC", CONEX);
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

        private void formataGridView()
        {
            var grade = dataGridView1;
            grade.AutoGenerateColumns = false;
            grade.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
            grade.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            //Alterar a Cor das Linhas alternadas no Grid
            grade.RowsDefaultCellStyle.BackColor = Color.White;
            grade.AlternatingRowsDefaultCellStyle.BackColor = Color.Gray;
            //Formata as colunas valor, vencimento e pagamento
            grade.Columns[7].DefaultCellStyle.Format = "C";
            grade.Columns[8].DefaultCellStyle.Format = "C";
            //Seleciona a linha inteira
            grade.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            //Não permite multiplas seleções
            grade.MultiSelect = false;
            // exibe nulos formatados
            //grade.DefaultCellStyle.NullValue = " - ";
            //permite que o texto maior que célula não seja truncado
            grade.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            //define o alinhamento à direita
            grade.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            grade.Columns[30].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        }
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
                decimal valorTotal = 0;
                string valor = "";
                if (dataGridView1.CurrentRow.Cells[7].Value != null)
                {
                    valor = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                    if (!valor.Equals(""))
                    {
                        for (int i = 0; i <= dataGridView1.RowCount - 1; i++)
                        {
                            if (dataGridView1.Rows[i].Cells[7].Value != null)
                                valorTotal += Convert.ToDecimal(dataGridView1.Rows[i].Cells[7].Value);
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
                MessageBox.Show("Erro ao Calcular, Verifique os Valores Texto_1\n'" + ex.Message + "'", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            try
            {
                decimal valorTotal2 = 0;
                string valor = "";
                if (dataGridView1.CurrentRow.Cells[30].Value != null)
                {
                    valor = dataGridView1.CurrentRow.Cells[30].Value.ToString();
                    if (!valor.Equals(""))
                    {
                        for (int i = 0; i <= dataGridView1.RowCount - 1; i++)
                        {
                            if (dataGridView1.Rows[i].Cells[30].Value != null)
                                valorTotal2 += Convert.ToDecimal(dataGridView1.Rows[i].Cells[30].Value);
                        }
                        if (valorTotal2 == 0)
                        {
                            MessageBox.Show("Nenhum registro encontrado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        txtTotal2.Text = valorTotal2.ToString("C");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao Calcular, Verifique os Valores Texto_2\n'" + ex.Message + "'", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            //7200000000
            lbHora.Text = (DateTime.Now.ToString("dd/MM/yy HH:mm:ss"));
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            CONEX = new MySqlConnection(@"server=localhost;database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            CONEX = new MySqlConnection(@"server=usga-servidor-m;database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
        }
        private void preencherCBmes_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void preencherCBano_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }
        private void button15_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter seach = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE Mes_ref='" + preencherCBmes.Text + "/" + preencherCBano.Text + "' ORDER BY id ASC", CONEX);
                DataTable seachSS = new DataTable();
                seach.Fill(seachSS);
                dataGridView1.DataSource = seachSS;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void metroTile1_Click(object sender, EventArgs e)
        {


        }
        public class FormBoleto : Form
        {
            public FormBoleto() { }

            private void btnSaveInput_Click(object sender, EventArgs e)
            {
                FormRel form1 = new FormRel();
                form1.txtURLBOLETO.ToString(); // How do I show my values on the first form?
                form1.ShowDialog();
            }
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            List<DataGrid> list = ((DataParameter)e.Argument).ProductList;
            string filename = ((DataParameter)e.Argument).Filename;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = excel.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet ws = (Worksheet)excel.ActiveSheet;
            excel.Visible = false;
            int index = 1;
            int process = dataGridView1.Columns.Count;
            foreach(DataGrid p in list)
            {
                if (!backgroundWorker.CancellationPending)
                {
                    backgroundWorker.ReportProgress(index++ * 100 / process);

                    for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                    {
                        ws.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                    }
                    // storing Each row and column value to excel sheet  
                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            ws.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                }
            }
            ws.SaveAs(Filename, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
            excel.Quit();
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
            lbStatus.Text = string.Format("Processando...{0}", e.ProgressPercentage);
            progressBar.Update();
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if(e.Error == null)
            {
                Thread.Sleep(100);
                lbStatus.Text = "Excel exportado com sucesso!";
            }
        }
        
        private void btnExport_Click(object sender, EventArgs e)
        {

        }

        private void txt_dif_boleto_TextChanged(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto ORDER BY Mes_ref ASC", CONEX);
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
                CONEX.Close();
                CONEX.Open();
                //Atualização de Status
                MySqlCommand update_3 = new MySqlCommand("UPDATE tb_boleto SET status='VENCIDA' WHERE data_venc < CURDATE() AND status ='A VENCER'", CONEX);
                update_3.ExecuteNonQuery();
            }
            catch(Exception err)
            {
                MessageBox.Show(err.Message);
            }

            try
            {
                CONEX.Close();
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

                MySqlCommand command = new MySqlCommand("SELECT COUNT(*) FROM tb_boleto WHERE data_venc BETWEEN @DataServidor AND @dataFuturo AND status !='PAGO'", CONEX);

                command.Parameters.AddWithValue("@dataFuturo", novadata);
                command.Parameters.AddWithValue("@DataServidor", dataHora);
                command.ExecuteNonQuery();


                int qtdVencer = Convert.ToInt32(command.ExecuteScalar());
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
