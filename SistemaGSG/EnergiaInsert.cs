using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using iTextSharp.text.pdf;
using MetroFramework;

namespace SistemaGSG
{
    public partial class Ceal : MetroFramework.Forms.MetroForm
    {
        private const string Texto = " Duplicidade!, Este Código Único já existe no Banco de Dados.\n Por Favor, Informe outro.";
        private const string Clear = " Limpo com Sucesso!.\n Por Favor, Prossiga sua Digitação.";
        string STATUS;
        string EMPRESA;
        string CUSTO;
        MySqlCommand cmd, prompt_cmd, prompt_notif;
        MySqlConnection CONEX,cn;
        string usuarioLogado = System.Environment.UserName;
        private void boxLocal_CheckedChanged(object sender, EventArgs e)
        {
        }
        private void boxTeste_CheckedChanged(object sender, EventArgs e)
        {
        }
        private void LimparTexts()
        {
                        //Limpar Campos apos a inserção no banco de dados.
                        cod_unico.Text = "";
                        mes_nf.Text = "";
                        txtFaz.Text = "";
                        vl_boleto.Text = "";
                        nfe.Text = "";
                        vl_multa.Text = "";
                        mesMulta.Text = "";
                        vl_base.Text = "";
                        vl_fecoep.Text = "";
                        txtMesdupl.Text = "";
                        txtValordupl.Text = "";
                        textBase1.Text = "";
                        cod_unico.Focus();
        }
        private void ConsultDuplicidate()
        {
            if (string.IsNullOrWhiteSpace(nfe.Text))
            {
                MessageBox.Show("Para prosseguir, insira o nº da Nf.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                try
                {
                    CONEX.Open();//Abrir Conexão.
                    MySqlCommand prompt = new MySqlCommand("SELECT COUNT(*) FROM tb_boleto WHERE nfe ='" + nfe.Text + "' ", CONEX);//Seleção da tabela no Banco de Dados.
                    prompt.ExecuteNonQuery();//Executa o comando.
                    int consultDB = Convert.ToInt32(prompt.ExecuteScalar());//Converte o resultado para números inteiros.
                    CONEX.Close();
                    if (consultDB > 0)//Verifica se o resultado for maior que zero(0), a execução inicia a Menssagem de que já existe contas, caso contrario faz a inserção no Banco.
                    {
                        LimparTexts();
                        MessageBox.Show(Texto);
                    }
                    else
                    {
                        try
                        {
                            dbinsert();
                        }
                        catch(Exception Err)
                        {
                            MessageBox.Show(Err.Message);
                        }
                    }
                }
                catch (NullReferenceException)
                {
                    MessageBox.Show("Olá Srº(a), " + usuarioLogado + " selecione uma conexão abaixo, para iniciar a\naplicação!.");
                }
                finally
                {

                }
            }
        }
        private void ItensPedido()
        {
            cn.Open();
            MySqlCommand com = new MySqlCommand();
            com.Connection = cn;
            com.CommandText = "SELECT * FROM tb_boleto WHERE err > 0";
            MySqlDataReader dr = com.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(dr);
            CountTXT.Text = dt.Rows.Count.ToString();
            int Contagem = Convert.ToInt32(CountTXT.Text);
            Contagem++;
            CountTXT.Text = Contagem.ToString();
            cn.Close();
        }
        private void ConsultNFE()
        {
            cn.Open();
            MySqlCommand MyCommand = new MySqlCommand();
            MyCommand.Connection = cn;
            MyCommand.CommandText = "SELECT * FROM tb_boleto ORDER BY id DESC";
            MySqlDataReader dreader = MyCommand.ExecuteReader();
            while (dreader.Read())
            {
                txtUltNFE.Text = dreader[19].ToString();
                break;
            }
            cn.Close();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            ConsultDuplicidate();
            ConsultNFE();
        }
        public Ceal()
        {
            InitializeComponent();
        }
        private void dbinsert()
        {
            /*****Modifica a data para a inserção no Banco de Dados********************************/
            dataemissao.Format = DateTimePickerFormat.Custom;
            dataemissao.CustomFormat = "yyyy-MM-dd";
            datavencimento.Format = DateTimePickerFormat.Custom;
            datavencimento.CustomFormat = "yyyy-MM-dd";
            /*************************************************************************************/

            /************************Converter para valor INT*************************************/
            int ValorIcms = Convert.ToInt32(preencherCBIcms.Text.Replace(" %", ""));
            /*************************************************************************************/
            CONEX.Open();
            //Verifica se o campo Valor do Boleto esta Preenchido
                if (string.IsNullOrEmpty(textBase1.Text))
                {
                    if (string.IsNullOrEmpty(vl_base.Text))
                    {
                        try
                        {
                            if (ValorIcms == 17)
                            {
                                prompt_cmd = new MySqlCommand("INSERT INTO `tb_boleto` (`id`, `material`, `desc_item`, `qtd`, `centro`, `custo`, `cod_imp`, `base_calculo`, `vl_icms`, `txt_pedido`, `material_dif`, `desc_item_dif`, `qtd_dif`, `centro_dif`, `custo_dif`, `cod_imp_dif`, `vl_dif`, `iva_dif`, `emissao`, `nfe`, `err`, `err_col`, `txt_miro`, `data_venc`, `Mes_ref`, `cod_unico`, `pedido`, `migo`, `miro`, `fecoep`, `valor_miro`, `status`, `empresa`, `mes_dupl`, `vl_dupl`, `now_date`) VALUES (NULL, '270743', '" + txtFaz.Text + "', '1', 'USGA', '" + CUSTO + "','PH', '" + vl_boleto.Text.Replace("R$ ","") + "', '0', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', NULL, NULL, NULL , NULL, NULL, NULL, NULL, '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '"+ CountTXT.Text +"',  '1', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ","") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", CONEX);

                                prompt_notif = new MySqlCommand("INSERT INTO notifica_vencimento (cod,data,status) VALUES ('" + cod_unico.Text + "',?date,NULL)", CONEX);
                                prompt_notif.Parameters.AddWithValue("?date", dataemissao.Value.AddDays(30));
                                prompt_cmd.Parameters.AddWithValue("?icms", preencherCBIcms.Text.Replace(" %", ""));
                                prompt_cmd.ExecuteNonQuery();
                                prompt_notif.ExecuteNonQuery();
                            }
                            if (ValorIcms == 18)
                            {
                                prompt_cmd = new MySqlCommand("INSERT INTO `tb_boleto` (`id`, `material`, `desc_item`, `qtd`, `centro`, `custo`, `cod_imp`, `base_calculo`, `vl_icms`, `txt_pedido`, `material_dif`, `desc_item_dif`, `qtd_dif`, `centro_dif`, `custo_dif`, `cod_imp_dif`, `vl_dif`, `iva_dif`, `emissao`, `nfe`, `err`, `err_col`, `txt_miro`, `data_venc`, `Mes_ref`, `cod_unico`, `pedido`, `migo`, `miro`, `fecoep`, `valor_miro`, `status`, `empresa`, `mes_dupl`, `vl_dupl`, `now_date`) VALUES (NULL, '272920',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','PH', '" + vl_boleto.Text.Replace("R$ ", "") + "',  '0', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', NULL, NULL, NULL, NULL, NULL, NULL, NULL, '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '"+ CountTXT.Text +"', '1', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", CONEX);

                                prompt_notif = new MySqlCommand("INSERT INTO notifica_vencimento (cod,data,status) VALUES ('" + cod_unico.Text + "',?date,NULL)", CONEX);
                                prompt_notif.Parameters.AddWithValue("?date", dataemissao.Value.AddDays(30));
                                prompt_cmd.Parameters.AddWithValue("?icms", preencherCBIcms.Text.Replace(" %", ""));
                                prompt_cmd.ExecuteNonQuery();
                                prompt_notif.ExecuteNonQuery();
                            }
                            if (ValorIcms == 25)
                            {
                                prompt_cmd = new MySqlCommand("INSERT INTO `tb_boleto` (`id`, `material`, `desc_item`, `qtd`, `centro`, `custo`, `cod_imp`, `base_calculo`, `vl_icms`, `txt_pedido`, `material_dif`, `desc_item_dif`, `qtd_dif`, `centro_dif`, `custo_dif`, `cod_imp_dif`, `vl_dif`, `iva_dif`, `emissao`, `nfe`, `err`, `err_col`, `txt_miro`, `data_venc`, `Mes_ref`, `cod_unico`, `pedido`, `migo`, `miro`, `fecoep`, `valor_miro`, `status`, `empresa`, `mes_dupl`, `vl_dupl`, `now_date`) VALUES (NULL, '271199',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','PH', '" + vl_boleto.Text.Replace("R$ ", "") + "',  '0', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', NULL, NULL, NULL, NULL, NULL, NULL, NULL, '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '"+ CountTXT.Text +"',  '1', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", CONEX);

                                prompt_notif = new MySqlCommand("INSERT INTO notifica_vencimento (cod,data,status) VALUES ('" + cod_unico.Text + "',?date,NULL)", CONEX);
                                prompt_notif.Parameters.AddWithValue("?date", dataemissao.Value.AddDays(30));
                                prompt_cmd.Parameters.AddWithValue("?icms", preencherCBIcms.Text.Replace(" %", ""));
                                prompt_cmd.ExecuteNonQuery();
                                prompt_notif.ExecuteNonQuery();
                            }
                            if (ValorIcms == 27)
                            {
                                prompt_cmd = new MySqlCommand("INSERT INTO `tb_boleto` (`id`, `material`, `desc_item`, `qtd`, `centro`, `custo`, `cod_imp`, `base_calculo`, `vl_icms`, `txt_pedido`, `material_dif`, `desc_item_dif`, `qtd_dif`, `centro_dif`, `custo_dif`, `cod_imp_dif`, `vl_dif`, `iva_dif`, `emissao`, `nfe`, `err`, `err_col`, `txt_miro`, `data_venc`, `Mes_ref`, `cod_unico`, `pedido`, `migo`, `miro`, `fecoep`, `valor_miro`, `status`, `empresa`, `mes_dupl`, `vl_dupl`, `now_date`) VALUES (NULL, '271229',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','PH', '" + vl_boleto.Text.Replace("R$ ", "") + "',  '0', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', NULL, NULL, NULL, NULL, NULL, NULL, NULL, '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '"+ CountTXT.Text +"', '1', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", CONEX);

                                prompt_notif = new MySqlCommand("INSERT INTO notifica_vencimento (cod,data,status) VALUES ('" + cod_unico.Text + "',?date,NULL)", CONEX);
                                prompt_notif.Parameters.AddWithValue("?date", dataemissao.Value.AddDays(30));
                                prompt_cmd.Parameters.AddWithValue("?icms", preencherCBIcms.Text.Replace(" %", ""));
                                prompt_cmd.ExecuteNonQuery();
                                prompt_notif.ExecuteNonQuery();
                            }
                            if (ValorIcms == 0)
                            {
                                prompt_cmd = new MySqlCommand("INSERT INTO `tb_boleto` (`id`, `material`, `desc_item`, `qtd`, `centro`, `custo`, `cod_imp`, `base_calculo`, `vl_icms`, `txt_pedido`, `material_dif`, `desc_item_dif`, `qtd_dif`, `centro_dif`, `custo_dif`, `cod_imp_dif`, `vl_dif`, `iva_dif`, `emissao`, `nfe`, `err`, `err_col`, `txt_miro`, `data_venc`, `Mes_ref`, `cod_unico`, `pedido`, `migo`, `miro`, `fecoep`, `valor_miro`, `status`, `empresa`, `mes_dupl`, `vl_dupl`, `now_date`) VALUES (NULL, '270743',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','PH', '" + vl_boleto.Text.Replace("R$ ", "") + "',  '0', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', NULL, NULL, NULL, NULL, NULL, NULL, '0', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '"+ CountTXT.Text +"',  '1', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", CONEX);

                                prompt_notif = new MySqlCommand("INSERT INTO notifica_vencimento (cod,data,status) VALUES ('" + cod_unico.Text + "',?date,NULL)", CONEX);
                                prompt_notif.Parameters.AddWithValue("?date", dataemissao.Value.AddDays(30));
                                prompt_cmd.Parameters.AddWithValue("?icms", preencherCBIcms.Text.Replace(" %", ""));
                                prompt_cmd.ExecuteNonQuery();
                                prompt_notif.ExecuteNonQuery();
                            }   
                        }
                        catch (Exception err)
                        {
                            MessageBox.Show(err.Message);
                        }
                    }
                    else
                    {
                        try
                        {
                        double vlDifTotal = Convert.ToDouble(vl_boleto.Text.Replace("R$ ", "")) - Convert.ToDouble(vl_base.Text.Replace("R$ ", ""));
                        if (ValorIcms == 17)
                            {
                                prompt_cmd = new MySqlCommand("INSERT INTO `tb_boleto` (`id`, `material`, `desc_item`, `qtd`, `centro`, `custo`, `cod_imp`, `base_calculo`, `vl_icms`, `txt_pedido`, `material_dif`, `desc_item_dif`, `qtd_dif`, `centro_dif`, `custo_dif`, `cod_imp_dif`, `vl_dif`, `iva_dif`, `emissao`, `nfe`, `err`, `err_col`, `txt_miro`, `data_venc`, `Mes_ref`, `cod_unico`, `pedido`, `migo`, `miro`, `fecoep`, `valor_miro`, `status`, `empresa`, `mes_dupl`, `vl_dupl`, `now_date`) VALUES (NULL, '270743', '" + txtFaz.Text + "', '1', 'USGA', '" + CUSTO + "','P1', '" + vl_base.Text.Replace("R$ ","") + "', '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 270743, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', 'PH', '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '"+ CountTXT.Text +"',  '2', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", CONEX);

                                prompt_notif = new MySqlCommand("INSERT INTO notifica_vencimento (cod,data,status) VALUES ('" + cod_unico.Text + "',?date,NULL)", CONEX);
                                prompt_notif.Parameters.AddWithValue("?date", dataemissao.Value.AddDays(30));
                                prompt_cmd.Parameters.AddWithValue("?icms", preencherCBIcms.Text.Replace(" %", ""));
                                prompt_cmd.ExecuteNonQuery();
                                prompt_notif.ExecuteNonQuery();
                            }
                            if (ValorIcms == 18)
                            {

                                prompt_cmd = new MySqlCommand("INSERT INTO `tb_boleto` (`id`, `material`, `desc_item`, `qtd`, `centro`, `custo`, `cod_imp`, `base_calculo`, `vl_icms`, `txt_pedido`, `material_dif`, `desc_item_dif`, `qtd_dif`, `centro_dif`, `custo_dif`, `cod_imp_dif`, `vl_dif`, `iva_dif`, `emissao`, `nfe`, `err`, `err_col`, `txt_miro`, `data_venc`, `Mes_ref`, `cod_unico`, `pedido`, `migo`, `miro`, `fecoep`, `valor_miro`, `status`, `empresa`, `mes_dupl`, `vl_dupl`, `now_date`) VALUES (NULL, '272920',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','P1', '" + vl_base.Text.Replace("R$ ", "") + "',  '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 272920, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', 'PH', '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '"+ CountTXT.Text +"', '2', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", CONEX);

                                prompt_notif = new MySqlCommand("INSERT INTO notifica_vencimento (cod,data,status) VALUES ('" + cod_unico.Text + "',?date,NULL)", CONEX);
                                prompt_notif.Parameters.AddWithValue("?date", dataemissao.Value.AddDays(30));
                                prompt_cmd.Parameters.AddWithValue("?icms", preencherCBIcms.Text.Replace(" %", ""));
                                prompt_cmd.ExecuteNonQuery();
                                prompt_notif.ExecuteNonQuery();
                            }
                            if (ValorIcms == 25)
                            {
                                prompt_cmd = new MySqlCommand("INSERT INTO `tb_boleto` (`id`, `material`, `desc_item`, `qtd`, `centro`, `custo`, `cod_imp`, `base_calculo`, `vl_icms`, `txt_pedido`, `material_dif`, `desc_item_dif`, `qtd_dif`, `centro_dif`, `custo_dif`, `cod_imp_dif`, `vl_dif`, `iva_dif`, `emissao`, `nfe`, `err`, `err_col`, `txt_miro`, `data_venc`, `Mes_ref`, `cod_unico`, `pedido`, `migo`, `miro`, `fecoep`, `valor_miro`, `status`, `empresa`, `mes_dupl`, `vl_dupl`, `now_date`) VALUES (NULL, '271199',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','P1', '" + vl_base.Text.Replace("R$ ", "") + "',  '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 271199, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', 'PH', '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '"+ CountTXT.Text +"',  '2', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", CONEX);

                                prompt_notif = new MySqlCommand("INSERT INTO notifica_vencimento (cod,data,status) VALUES ('" + cod_unico.Text + "',?date,NULL)", CONEX);
                                prompt_notif.Parameters.AddWithValue("?date", dataemissao.Value.AddDays(30));
                                prompt_cmd.Parameters.AddWithValue("?icms", preencherCBIcms.Text.Replace(" %", ""));
                                prompt_cmd.ExecuteNonQuery();
                                prompt_notif.ExecuteNonQuery();
                            }
                            if (ValorIcms == 27)
                            {
                                prompt_cmd = new MySqlCommand("INSERT INTO `tb_boleto` (`id`, `material`, `desc_item`, `qtd`, `centro`, `custo`, `cod_imp`, `base_calculo`, `vl_icms`, `txt_pedido`, `material_dif`, `desc_item_dif`, `qtd_dif`, `centro_dif`, `custo_dif`, `cod_imp_dif`, `vl_dif`, `iva_dif`, `emissao`, `nfe`, `err`, `err_col`, `txt_miro`, `data_venc`, `Mes_ref`, `cod_unico`, `pedido`, `migo`, `miro`, `fecoep`, `valor_miro`, `status`, `empresa`, `mes_dupl`, `vl_dupl`, `now_date`) VALUES (NULL, '271229',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','P1', '" + vl_base.Text.Replace("R$ ", "") + "',  '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 271229, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', 'PH', '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '"+ CountTXT.Text +"',  '2', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", CONEX);

                                prompt_notif = new MySqlCommand("INSERT INTO notifica_vencimento (cod,data,status) VALUES ('" + cod_unico.Text + "',?date,NULL)", CONEX);
                                prompt_notif.Parameters.AddWithValue("?date", dataemissao.Value.AddDays(30));
                                prompt_cmd.Parameters.AddWithValue("?icms", preencherCBIcms.Text.Replace(" %", ""));
                                prompt_cmd.ExecuteNonQuery();
                                prompt_notif.ExecuteNonQuery();
                            }
                            if (ValorIcms == 0)
                            {
                                prompt_cmd = new MySqlCommand("INSERT INTO `tb_boleto` (`id`, `material`, `desc_item`, `qtd`, `centro`, `custo`, `cod_imp`, `base_calculo`, `vl_icms`, `txt_pedido`, `material_dif`, `desc_item_dif`, `qtd_dif`, `centro_dif`, `custo_dif`, `cod_imp_dif`, `vl_dif`, `iva_dif`, `emissao`, `nfe`, `err`, `err_col`, `txt_miro`, `data_venc`, `Mes_ref`, `cod_unico`, `pedido`, `migo`, `miro`, `fecoep`, `valor_miro`, `status`, `empresa`, `mes_dupl`, `vl_dupl`, `now_date`) VALUES (NULL, '270743',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','PH', '" + vl_base.Text.Replace("R$ ", "") + "',  '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 270743, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', NULL, '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '"+ CountTXT.Text +"',  '2', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", CONEX);

                                prompt_notif = new MySqlCommand("INSERT INTO notifica_vencimento (cod,data,status) VALUES ('" + cod_unico.Text + "',?date,NULL)", CONEX);
                                prompt_notif.Parameters.AddWithValue("?date", dataemissao.Value.AddDays(30));
                                prompt_cmd.Parameters.AddWithValue("?icms", preencherCBIcms.Text.Replace(" %", ""));
                                prompt_cmd.ExecuteNonQuery();
                                prompt_notif.ExecuteNonQuery();
                            }
                        }
                        catch (Exception err)
                        {
                            MessageBox.Show(err.Message);
                        }
                    }
                }else{
                    try
                    {
                    double vlDifTotal = Convert.ToDouble(vl_boleto.Text.Replace("R$ ", "")) - Convert.ToDouble(vl_base.Text.Replace("R$ ", ""));
                    if (ValorIcms == 17)
                        {
                            prompt_cmd = new MySqlCommand("INSERT INTO `tb_boleto` (`id`, `material`, `desc_item`, `qtd`, `centro`, `custo`, `cod_imp`, `base_calculo`, `vl_icms`, `txt_pedido`, `material_dif`, `desc_item_dif`, `qtd_dif`, `centro_dif`, `custo_dif`, `cod_imp_dif`, `vl_dif`, `iva_dif`, `emissao`, `nfe`, `err`, `err_col`, `txt_miro`, `data_venc`, `Mes_ref`, `cod_unico`, `pedido`, `migo`, `miro`, `fecoep`, `valor_miro`, `status`, `empresa`, `mes_dupl`, `vl_dupl`, `now_date`) VALUES (NULL, '270743', '" + txtFaz.Text + "', '1', 'USGA', '" + CUSTO + "','P1', '" + vl_base.Text.Replace("R$ ", "") + "', '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 270743, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', 'PH', '" + vlDifTotal + "', 0, '" + this.dataemissao.Text + "', '" + nfe.Text + "', '"+ CountTXT.Text +"', NULL , '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, '" + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", CONEX);

                            prompt_notif = new MySqlCommand("INSERT INTO notifica_vencimento (cod,data,status) VALUES ('" + cod_unico.Text + "',?date,NULL)", CONEX);
                            prompt_notif.Parameters.AddWithValue("?date", dataemissao.Value.AddDays(30));
                            prompt_cmd.Parameters.AddWithValue("?icms", preencherCBIcms.Text.Replace(" %", ""));
                            prompt_cmd.ExecuteNonQuery();
                            prompt_notif.ExecuteNonQuery();
                        }
                        if (ValorIcms == 18)
                        {
                            prompt_cmd = new MySqlCommand("INSERT INTO `tb_boleto` (`id`, `material`, `desc_item`, `qtd`, `centro`, `custo`, `cod_imp`, `base_calculo`, `vl_icms`, `txt_pedido`, `material_dif`, `desc_item_dif`, `qtd_dif`, `centro_dif`, `custo_dif`, `cod_imp_dif`, `vl_dif`, `iva_dif`, `emissao`, `nfe`, `err`, `err_col`, `txt_miro`, `data_venc`, `Mes_ref`, `cod_unico`, `pedido`, `migo`, `miro`, `fecoep`, `valor_miro`, `status`, `empresa`, `mes_dupl`, `vl_dupl`, `now_date`) VALUES (NULL, '272920',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','P1', '" + vl_base.Text.Replace("R$ ", "") + "',  '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 272920, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', 'PH', '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '"+ CountTXT.Text +"', NULL, '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, '" + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", CONEX);

                            prompt_notif = new MySqlCommand("INSERT INTO notifica_vencimento (cod,data,status) VALUES ('" + cod_unico.Text + "',?date,NULL)", CONEX);
                            prompt_notif.Parameters.AddWithValue("?date", dataemissao.Value.AddDays(30));
                            prompt_cmd.Parameters.AddWithValue("?icms", preencherCBIcms.Text.Replace(" %", ""));
                            prompt_cmd.ExecuteNonQuery();
                            prompt_notif.ExecuteNonQuery();
                        }
                        if (ValorIcms == 25)
                        {
                            prompt_cmd = new MySqlCommand("INSERT INTO `tb_boleto` (`id`, `material`, `desc_item`, `qtd`, `centro`, `custo`, `cod_imp`, `base_calculo`, `vl_icms`, `txt_pedido`, `material_dif`, `desc_item_dif`, `qtd_dif`, `centro_dif`, `custo_dif`, `cod_imp_dif`, `vl_dif`, `iva_dif`, `emissao`, `nfe`, `err`, `err_col`, `txt_miro`, `data_venc`, `Mes_ref`, `cod_unico`, `pedido`, `migo`, `miro`, `fecoep`, `valor_miro`, `status`, `empresa`, `mes_dupl`, `vl_dupl`, `now_date`) VALUES (NULL, '271199',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','P1', '" + vl_base.Text.Replace("R$ ", "") + "',  '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 271199, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', 'PH', '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '"+ CountTXT.Text +"',  NULL, '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, '" + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", CONEX);

                            prompt_notif = new MySqlCommand("INSERT INTO notifica_vencimento (cod,data,status) VALUES ('" + cod_unico.Text + "',?date,NULL)", CONEX);
                            prompt_notif.Parameters.AddWithValue("?date", dataemissao.Value.AddDays(30));
                            prompt_cmd.Parameters.AddWithValue("?icms", preencherCBIcms.Text.Replace(" %", ""));
                            prompt_cmd.ExecuteNonQuery();
                            prompt_notif.ExecuteNonQuery();
                        }
                        if (ValorIcms == 27)
                        {
                            prompt_cmd = new MySqlCommand("INSERT INTO `tb_boleto` (`id`, `material`, `desc_item`, `qtd`, `centro`, `custo`, `cod_imp`, `base_calculo`, `vl_icms`, `txt_pedido`, `material_dif`, `desc_item_dif`, `qtd_dif`, `centro_dif`, `custo_dif`, `cod_imp_dif`, `vl_dif`, `iva_dif`, `emissao`, `nfe`, `err`, `err_col`, `txt_miro`, `data_venc`, `Mes_ref`, `cod_unico`, `pedido`, `migo`, `miro`, `fecoep`, `valor_miro`, `status`, `empresa`, `mes_dupl`, `vl_dupl`, `now_date`) VALUES (NULL, '271229',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','P1', '" + vl_base.Text.Replace("R$ ", "") + "',  '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 271229, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', 'PH', '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '"+ CountTXT.Text +"',  NULL, '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, '" + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", CONEX);

                            prompt_notif = new MySqlCommand("INSERT INTO notifica_vencimento (cod,data,status) VALUES ('" + cod_unico.Text + "',?date,NULL)", CONEX);
                            prompt_notif.Parameters.AddWithValue("?date", dataemissao.Value.AddDays(30));
                            prompt_cmd.Parameters.AddWithValue("?icms", preencherCBIcms.Text.Replace(" %", ""));
                            prompt_cmd.ExecuteNonQuery();
                            prompt_notif.ExecuteNonQuery();
                        }
                        if (ValorIcms == 0)
                        {
                            prompt_cmd = new MySqlCommand("INSERT INTO `tb_boleto` (`id`, `material`, `desc_item`, `qtd`, `centro`, `custo`, `cod_imp`, `base_calculo`, `vl_icms`, `txt_pedido`, `material_dif`, `desc_item_dif`, `qtd_dif`, `centro_dif`, `custo_dif`, `cod_imp_dif`, `vl_dif`, `iva_dif`, `emissao`, `nfe`, `err`, `err_col`, `txt_miro`, `data_venc`, `Mes_ref`, `cod_unico`, `pedido`, `migo`, `miro`, `fecoep`, `valor_miro`, `status`, `empresa`, `mes_dupl`, `vl_dupl`, `now_date`) VALUES (NULL, '270743',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','PH', '" + vl_base.Text.Replace("R$ ", "") + "',  '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 270743, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', NULL, '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '"+ CountTXT.Text +"', NULL, '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, '" + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", CONEX);

                            prompt_notif = new MySqlCommand("INSERT INTO notifica_vencimento (cod,data,status) VALUES ('" + cod_unico.Text + "',?date,NULL)", CONEX);
                            prompt_notif.Parameters.AddWithValue("?date", dataemissao.Value.AddDays(30));
                            prompt_cmd.Parameters.AddWithValue("?icms", preencherCBIcms.Text.Replace(" %", ""));
                            prompt_cmd.ExecuteNonQuery();
                            prompt_notif.ExecuteNonQuery();
                        }
                    }
                    catch (Exception err)
                    {
                        MessageBox.Show(err.Message);
                    }
                }

            if (string.IsNullOrWhiteSpace(vl_multa.Text))
            {

            }
            else
            {
                cmd = new MySqlCommand("INSERT INTO contas_multa (cod,mes,valor,empresa) VALUES ('" + cod_unico.Text + "','" + mesMulta.Text + "','" + vl_multa.Text.Replace("R$ ", "") + "','" + EMPRESA + "')", CONEX);
                cmd.ExecuteNonQuery();
            }
                
            //Fechar Conexão
            CONEX.Close();

            //Limpar Campos apos a inserção no banco de dados.
            LimparTexts();

            //Adiciona o Número do item
            ItensPedido();
            /*****Retorna o resultado dos campos após a inserção no Banco de Dados****************/
            dataemissao.Format = DateTimePickerFormat.Custom;
            dataemissao.CustomFormat = "dd/MM/yyyy";
            datavencimento.Format = DateTimePickerFormat.Custom;
            datavencimento.CustomFormat = "dd/MM/yyyy";
            /*************************************************************************************/

            MessageBox.Show("Inserido com Sucesso!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.None);
        }
        //Mascara
        private void Ceal_Load(object sender, EventArgs e)
        {
            txtUltNFE.Enabled = false;
            try
            {
                cn = new MySqlConnection(@"server=localhost;database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
                cn.Open();
                MySqlCommand com = new MySqlCommand();
                com.Connection = cn;
                com.CommandText = "SELECT * FROM tb_boleto WHERE err > 0";
                MySqlDataReader dr = com.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(dr);
                CountTXT.Text = dt.Rows.Count.ToString();
                int Contagem = Convert.ToInt32(CountTXT.Text);
                Contagem++;
                CountTXT.Text = Contagem.ToString();
                cn.Close();
            }catch(Exception Err)
            {
                MessageBox.Show(Err.Message);
            }

            try
            {
                cn = new MySqlConnection(@"server=usga-servidor-m;database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
                cn.Open();

                MySqlCommand com = new MySqlCommand();
                com.Connection = cn;
                com.CommandText = "SELECT porcentagem FROM icms";
                MySqlDataReader dr = com.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(dr);
                preencherCBIcms.DisplayMember = "porcentagem";
                preencherCBIcms.DataSource = dt;
                cn.Close();
            }
            catch
            {
                cn = new MySqlConnection(@"server=localhost;database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
                cn.Open();

                MySqlCommand com = new MySqlCommand();
                com.Connection = cn;
                com.CommandText = "SELECT porcentagem FROM icms";
                MySqlDataReader dr = com.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(dr);
                preencherCBIcms.DisplayMember = "porcentagem";
                preencherCBIcms.DataSource = dt;
                cn.Close();
            }
                vl_boleto.Enabled = true;
                vl_multa.Enabled = true;
                mesMulta.Enabled = true;
                vl_base.Enabled = true;
                preencherCBIcms.Enabled = true;
                txtMesdupl.Enabled = false;
                txtValordupl.Enabled = false;
                textBase1.Enabled = false;

            if (rdDupl.Checked)
            {
                vl_boleto.Enabled = false;
                vl_multa.Enabled = false;
                mesMulta.Enabled = false;
                vl_base.Enabled = false;
                preencherCBIcms.Enabled = false;
                txtMesdupl.Enabled = true;
                txtValordupl.Enabled = true;
            }
            else
            {
                vl_boleto.Enabled = true;
                vl_multa.Enabled = true;
                mesMulta.Enabled = true;
                vl_base.Enabled = true;
                preencherCBIcms.Enabled = true;
                txtMesdupl.Enabled = false;
                txtValordupl.Enabled = false;
            }
        }
        private void preencherCBIcms_SelectedIndexChanged(object sender, EventArgs e)
        {
            
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
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            CONEX = new MySqlConnection(@"server=usga-servidor-m;database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
        }
        private void metroRadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            STATUS = "PAGO";
        }
        private void rdVenc_CheckedChanged(object sender, EventArgs e)
        {
            STATUS = "VENCIDA";
        }
        private void rdAven_CheckedChanged(object sender, EventArgs e)
        {
            STATUS = "A VENCER";
        }
        private void metroRadioButton1_CheckedChanged_1(object sender, EventArgs e)
        {
            CUSTO = "SG01040201";
        }
        private void metroRadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            CUSTO = "SG01040101";
        }
        private void label5_Click(object sender, EventArgs e)
        {

        }
        private void btnClear_Click(object sender, EventArgs e)
        {
            //Limpar Campos apos a inserção no banco de dados.
            LimparTexts();
            MessageBox.Show(Clear);
        }
        private void txtUltNFE_MouseDoubleClick(object sender, MouseEventArgs e)
        {
        }
        private void btnView_Click(object sender, EventArgs e)
        {
            var PDFReader = new ReadPDF();
            PDFReader.Show();
        }
        private void rdDupl_CheckedChanged(object sender, EventArgs e)
        {
            STATUS = "DUPLICIDADE";

            if (rdDupl.Checked)
            {
                vl_boleto.Enabled = false;
                vl_multa.Enabled = false;
                mesMulta.Enabled = false;
                vl_base.Enabled = false;
                textBase1.Enabled = true;
                preencherCBIcms.Enabled = false;
                txtMesdupl.Enabled = true;
                txtValordupl.Enabled = true;
            }
            else
            {
                vl_boleto.Enabled = true;
                vl_multa.Enabled = true;
                mesMulta.Enabled = true;
                vl_base.Enabled = true;
                preencherCBIcms.Enabled = true;
                textBase1.Enabled = false;
                txtMesdupl.Enabled = false;
                txtValordupl.Enabled = false;
            }
        }
        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            CONEX = new MySqlConnection(@"server=localhost;database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
        }

        private void txtPedido_DoubleClick(object sender, EventArgs e)
        {
            MetroMessageBox.Show(this, "Your message here.", "Title Here",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }

        private void txtMigo_DoubleClick(object sender, EventArgs e)
        {
            MetroMessageBox.Show(this, "Your message here.", "Title Here", MessageBoxButtons.OK, MessageBoxIcon.Stop);
        }

        private void txtMiro_DoubleClick(object sender, EventArgs e)
        {
            MetroMessageBox.Show(this, "Your message here.", "Title Here", MessageBoxButtons.OK, MessageBoxIcon.Question);
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
