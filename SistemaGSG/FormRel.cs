using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;
using DataTable = System.Data.DataTable;
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
            public string Filename { get; private set; }
            public List<DataGrid> ProductList { get; internal set; }
        }
        public string Filename { get; private set; }
        private void FormRel_Load(object sender, EventArgs e)
        {

            if (ConexaoDados.GetConnectionEquatorial().State == ConnectionState.Open)
            {
                MySqlCommand com = new MySqlCommand();
                com.Connection = ConexaoDados.GetConnectionEquatorial();
                com.CommandText = "SELECT Id FROM tb_mes ORDER BY `tb_mes`.`ano` ASC";
                MySqlDataReader dr = com.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(dr);
                preencherCBmes.DisplayMember = "Id";
                preencherCBmes.DataSource = dt;

                MySqlCommand ComAno = new MySqlCommand();
                ComAno.Connection = ConexaoDados.GetConnectionEquatorial();
                ComAno.CommandText = "SELECT ano FROM tb_mes ORDER BY `tb_mes`.`ano` ASC";
                MySqlDataReader DatAno = ComAno.ExecuteReader();
                DataTable dataTableAno = new DataTable();
                dataTableAno.Load(DatAno);
                preencherCBano.DisplayMember = "ano";
                preencherCBano.DataSource = dataTableAno;

                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM `tb_boleto`", ConexaoDados.GetConnectionEquatorial());
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView2.DataSource = SS;

            }
            else
            {
                MessageBox.Show("Erro na Conexão!.");
            }
            ConexaoDados.GetConnectionEquatorial().Close();
        }
        private void button14_Click(object sender, EventArgs e)
        {
            var PedidoSAP = new FormPedido();
            PedidoSAP.Show();
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
                Close();
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja fazer entrada de Energia?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Ceal ceal = new Ceal();
                ceal.Show();
                Close();
            }
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter seach = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE cod_unico='" + textBox1.Text + "' ORDER BY Mes_ref DESC", ConexaoDados.GetConnectionEquatorial());
                DataTable seachSS = new DataTable();
                seach.Fill(seachSS);
                dataGridView2.DataSource = seachSS;
                ConexaoDados.GetConnectionEquatorial().Close();
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
                MySqlDataAdapter updateCEAL = new MySqlDataAdapter("UPDATE tb_boleto SET cod_unico='" + txtCod.Text + "', vl_dif='" + txt_dif_boleto.Text.Replace(".", "") + "' ,desc_item='" + txtFaz.Text + "', Mes_ref='" + txtMess.Text + "', data_venc='" + this.dtVencimento.Text + "', valor_miro='" + txtvalor.Text.Replace(".", "") + "', status='" + txtstatus.Text + "', pedido='" + txtPedido.Text + "', migo='" + txtmigo.Text + "', miro='" + txtMiro.Text + "', nfe='" + txtNf.Text + "', vl_icms='" + txtICMS.Text + "', base_calculo='" + txtVlBase.Text.Replace(".", "") + "' WHERE Id='" + textBox10.Text + "'", ConexaoDados.GetConnectionEquatorial());
                DataTable seachUpdate = new DataTable();
                updateCEAL.Fill(seachUpdate);
                dataGridView2.DataSource = seachUpdate;

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
        private void dataGridView2_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter updateCEAL = new MySqlDataAdapter("DELETE FROM tb_boleto WHERE Id='" + textBox10.Text + "'", ConexaoDados.GetConnectionEquatorial());
                DataTable seachUpdate = new DataTable();
                updateCEAL.Fill(seachUpdate);
                dataGridView2.DataSource = seachUpdate;

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
            printer.PrintDataGridView(dataGridView2);
        }
        string SELECAO;
        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE empresa='CEAL' ORDER BY id ASC", ConexaoDados.GetConnectionEquatorial());
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView2.DataSource = SS;
                ConexaoDados.GetConnectionEquatorial().Close();
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
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE empresa='CELPE' ORDER BY id ASC", ConexaoDados.GetConnectionEquatorial());
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView2.DataSource = SS;
                ConexaoDados.GetConnectionEquatorial().Close();
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
                MySqlDataAdapter seach = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE data_venc between '" + this.dateTimePicker2.Text.Replace("/", "-").ToString() + "' AND '" + this.dateTimePicker3.Text.Replace("/", "-").ToString() + "' AND empresa='" + SELECAO + "' ORDER BY id ASC", ConexaoDados.GetConnectionEquatorial());
                DataTable seachSS = new DataTable();
                seach.Fill(seachSS);
                dataGridView2.DataSource = seachSS;
                ConexaoDados.GetConnectionEquatorial().Close();
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
        private void dataGridView2_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {

        }
        private void label9_Click(object sender, EventArgs e)
        {

        }
        private void definirFiltro_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter seach = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE now_date between '" + this.dateTimePicker4.Text + "' AND '" + this.dateTimePicker5.Text + "' AND empresa='" + SELECAO + "' ORDER BY id ASC", ConexaoDados.GetConnectionEquatorial());
                DataTable seachSS = new DataTable();
                seach.Fill(seachSS);
                dataGridView2.DataSource = seachSS;
                ConexaoDados.GetConnectionEquatorial().Close();
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
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE status='PAGO' ORDER BY Mes_ref ASC", ConexaoDados.GetConnectionEquatorial());
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView2.DataSource = SS;
                ConexaoDados.GetConnectionEquatorial().Close();
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
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE status='VENCIDA' ORDER BY id ASC", ConexaoDados.GetConnectionEquatorial());
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView2.DataSource = SS;
                ConexaoDados.GetConnectionEquatorial().Close();
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
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE status='A VENCER' ORDER BY id ASC", ConexaoDados.GetConnectionEquatorial());
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView2.DataSource = SS;
                ConexaoDados.GetConnectionEquatorial().Close();
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
            Close();
        }
        Bitmap bmp;

        private void formataGridView()
        {
        }
        private void btnPrintview_Click(object sender, EventArgs e)
        {
            int height = dataGridView2.Height;
            dataGridView2.Height = dataGridView2.RowCount * dataGridView2.RowTemplate.Height * 2;
            bmp = new Bitmap(dataGridView2.Width, dataGridView2.Height);
            dataGridView2.DrawToBitmap(bmp, new Rectangle(0, 0, dataGridView2.Width, dataGridView2.Height));
            dataGridView2.Height = height;
            printPreviewDialog1.ShowDialog();
        }
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawImage(bmp, 0, 0);
        }
        private void dataGridView2_CellClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            //7200000000
            lbHora.Text = (DateTime.Now.ToString("dd/MM/yy HH:mm:ss"));
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
                MySqlDataAdapter seach = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE Mes_ref='" + preencherCBmes.Text + "/" + preencherCBano.Text + "' ORDER BY id ASC", ConexaoDados.GetConnectionEquatorial());
                DataTable seachSS = new DataTable();
                seach.Fill(seachSS);
                dataGridView2.DataSource = seachSS;
                ConexaoDados.GetConnectionEquatorial().Close();
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
        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
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
            int process = dataGridView2.Columns.Count;
            foreach (DataGrid p in list)
            {
                if (!backgroundWorker.CancellationPending)
                {
                    backgroundWorker.ReportProgress(index++ * 100 / process);

                    for (int i = 1; i < dataGridView2.Columns.Count + 1; i++)
                    {
                        ws.Cells[1, i] = dataGridView2.Columns[i - 1].HeaderText;
                    }
                    // storing Each row and column value to excel sheet  
                    for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < dataGridView2.Columns.Count; j++)
                        {
                            ws.Cells[i + 2, j + 1] = dataGridView2.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                }
            }
            ws.SaveAs(Filename, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
            excel.Quit();
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //progressBar.Value = e.ProgressPercentage;
            //lbStatus.Text = string.Format("Processando...{0}", e.ProgressPercentage);
            //progressBar.Update();
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                Thread.Sleep(100);
                //lbStatus.Text = "Excel exportado com sucesso!";
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {

        }

        private void txt_dif_boleto_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtTotal2_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView2_MouseDoubleClick_1(object sender, MouseEventArgs e)
        {
            textBox10.Text = dataGridView2.SelectedRows[0].Cells[0].Value.ToString();
            txtCod.Text = dataGridView2.SelectedRows[0].Cells[25].Value.ToString();
            txtMess.Text = dataGridView2.SelectedRows[0].Cells[24].Value.ToString();
            txtvalor.Text = dataGridView2.SelectedRows[0].Cells[30].Value.ToString();
            txtFaz.Text = dataGridView2.SelectedRows[0].Cells[2].Value.ToString();
            this.dtVencimento.Text = dataGridView2.SelectedRows[0].Cells[23].Value.ToString();
            txtstatus.Text = dataGridView2.SelectedRows[0].Cells[31].Value.ToString();
            txtPedido.Text = dataGridView2.SelectedRows[0].Cells[26].Value.ToString();
            txtMiro.Text = dataGridView2.SelectedRows[0].Cells[28].Value.ToString();
            txtmigo.Text = dataGridView2.SelectedRows[0].Cells[27].Value.ToString();
            txtICMS.Text = dataGridView2.SelectedRows[0].Cells[8].Value.ToString();
            txtNf.Text = dataGridView2.SelectedRows[0].Cells[19].Value.ToString();
            txtVlBase.Text = dataGridView2.SelectedRows[0].Cells[7].Value.ToString();
            txt_dif_boleto.Text = dataGridView2.SelectedRows[0].Cells[16].Value.ToString();

        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            var url = dataGridView2.SelectedRows[0].Cells["Column37"].Value;
            if (string.IsNullOrEmpty(url.ToString()))
            {

            }
            else
            {
                try
                {
                    Process.Start("chrome.exe", url.ToString());
                }
                catch
                {
                    Process.Start("msedge.exe", url.ToString());
                }
            }

            try
            {
                decimal valorTotal = 0;
                string valor = "";
                if (dataGridView2.CurrentRow.Cells["Column8"].Value != null)
                {
                    valor = dataGridView2.CurrentRow.Cells["Column8"].Value.ToString();
                    if (!valor.Equals(""))
                    {
                        for (int i = 0; i <= dataGridView2.RowCount - 1; i++)
                        {
                            if (dataGridView2.Rows[i].Cells["Column8"].Value != null)
                                valorTotal += Convert.ToDecimal(dataGridView2.Rows[i].Cells["Column8"].Value);
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
                if (dataGridView2.CurrentRow.Cells["Column31"].Value != null)
                {
                    valor = dataGridView2.CurrentRow.Cells["Column31"].Value.ToString();
                    if (!valor.Equals(""))
                    {
                        for (int i = 0; i <= dataGridView2.RowCount - 1; i++)
                        {
                            if (dataGridView2.Rows[i].Cells["Column31"].Value != null)
                                valorTotal2 += Convert.ToDecimal(dataGridView2.Rows[i].Cells["Column31"].Value);
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

        private void dataGridView2_Click(object sender, EventArgs e)
        {
        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto ORDER BY Mes_ref ASC", ConexaoDados.GetConnectionEquatorial());
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView2.DataSource = SS;
                ConexaoDados.GetConnectionEquatorial().Close();
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
                //Atualização de Status
                MySqlCommand update_3 = new MySqlCommand("UPDATE tb_boleto SET status='VENCIDA' WHERE data_venc < GETDATE() AND status ='A VENCER'", ConexaoDados.GetConnectionEquatorial());
                update_3.ExecuteNonQuery();
                ConexaoDados.GetConnectionEquatorial().Close();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }

            try
            {

                MySqlCommand cmd = new MySqlCommand("SELECT GETDATE()", ConexaoDados.GetConnectionEquatorial());
                DateTime DataServidor = Convert.ToDateTime(cmd.ExecuteScalar());
                string novadata = DataServidor.AddDays(+10).ToShortDateString();

                label12.Text = Convert.ToString(DataServidor);
                label13.Text = novadata;

                dataHora = DataServidor;
                minutos = dataHora.Minute;
                segundos = dataHora.Second;
                milisegundos = dataHora.Millisecond;

                MySqlCommand command = new MySqlCommand("SELECT COUNT(*) FROM tb_boleto WHERE data_venc BETWEEN @DataServidor AND @dataFuturo AND status !='PAGO'", ConexaoDados.GetConnectionEquatorial());

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
            ConexaoDados.GetConnectionEquatorial().Close();
        }
    }
}
