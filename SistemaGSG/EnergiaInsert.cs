using MetroFramework;
using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Windows.Forms;

namespace SistemaGSG
{
    public partial class Ceal : MetroFramework.Forms.MetroForm
    {
        private const string Texto = " Duplicidade!, Este Código Único já existe no Banco de Dados.\n Por Favor, Informe outro.";
        private const string Clear = " Limpo com Sucesso!.\n Por Favor, Prossiga sua Digitação.";
        string STATUS;
        string EMPRESA;
        string CUSTO;
        string usuarioLogado = dados.Usuario;
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
            if (string.IsNullOrWhiteSpace(nfe.Text.Replace("-", "")))
            {
                MessageBox.Show("Para prosseguir, insira o nº da Nf.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                try
                {
                    //Abrir Conexão.
                    MySqlCommand prompt = new MySqlCommand("SELECT COUNT(*) FROM tb_boleto WHERE nfe ='" + nfe.Text + "' ", ConexaoDados.GetConnectionEquatorial());//Seleção da tabela no Banco de Dados.
                    prompt.ExecuteNonQuery();//Executa o comando.
                    int consultDB = Convert.ToInt32(prompt.ExecuteScalar());//Converte o resultado para números inteiros.
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
                        catch (Exception Err)
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
                    ConexaoDados.GetConnectionEquatorial().Close();
                }
            }
        }
        private void ItensPedido()
        {
            try
            {
                MySqlCommand com = new MySqlCommand();
                com.Connection = ConexaoDados.GetConnectionEquatorial();
                com.CommandText = "SELECT * FROM tb_boleto WHERE err > 0";
                MySqlDataReader dr = com.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(dr);
                //Contagem de Linhas
                CountTXT.Text = dt.Rows.Count.ToString();
                int Contagem = Convert.ToInt32(CountTXT.Text);
                //Adiciona mais um Item
                Contagem++;
                CountTXT.Text = Contagem.ToString();
                //if(Contagem > 10)
                //{
                //   var CriarPedido = new FormPedido(txthost.Text);
                // CriarPedido.Show();
                //}
            }
            catch (Exception Error)
            {
                MessageBox.Show(Error.Message);
            }
            finally
            {
                ConexaoDados.GetConnectionEquatorial().Close();
            }
        }
        private void ConsultNFE()
        {
            try
            {
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = ConexaoDados.GetConnectionEquatorial();
                MyCommand.CommandText = "SELECT * FROM tb_boleto ORDER BY id DESC";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    txtUltNFE.Text = dreader[19].ToString();
                    break;
                }
            }
            catch (Exception Error)
            {
                MessageBox.Show(Error.Message);
            }
            finally
            {
                ConexaoDados.GetConnectionEquatorial().Close();
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                ConsultDuplicidate();
                ConsultNFE();
            }
            catch (Exception Errorrr)
            {
                MessageBox.Show(Errorrr.Message);
            }
            finally
            {
                ConexaoDados.GetConnectionEquatorial().Close();
            }
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

            //Verifica se o campo Valor do Boleto esta Preenchido
            if (string.IsNullOrEmpty(textBase1.Text))
            {
                if (string.IsNullOrEmpty(vl_base.Text))
                {
                    try
                    {
                        if (ValorIcms == 17)
                        {
                            MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_boleto (material, desc_item, qtd, centro, custo, cod_imp, base_calculo, vl_icms, txt_pedido, material_dif, desc_item_dif, qtd_dif, centro_dif, custo_dif, cod_imp_dif, vl_dif, iva_dif, emissao, nfe, err, err_col, txt_miro, data_venc, Mes_ref, cod_unico, pedido, migo, miro, fecoep, valor_miro, status, empresa, mes_dupl, vl_dupl, now_date) VALUES ('270743', '" + txtFaz.Text + "', '1', 'USGA', '" + CUSTO + "','PH', '" + vl_boleto.Text.Replace("R$ ", "") + "', '0', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', NULL, NULL, NULL , NULL, NULL, NULL, NULL, '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '" + CountTXT.Text + "',  '1', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", ConexaoDados.GetConnectionEquatorial());
                            prompt_cmd.ExecuteNonQuery();
                        }
                        if (ValorIcms == 18)
                        {
                            MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_boleto (material, desc_item, qtd, centro, custo, cod_imp, base_calculo, vl_icms, txt_pedido, material_dif, desc_item_dif, qtd_dif, centro_dif, custo_dif, cod_imp_dif, vl_dif, iva_dif, emissao, nfe, err, err_col, txt_miro, data_venc, Mes_ref, cod_unico, pedido, migo, miro, fecoep, valor_miro, status, empresa, mes_dupl, vl_dupl, now_date) VALUES ('272920',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','PH', '" + vl_boleto.Text.Replace("R$ ", "") + "',  '0', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', NULL, NULL, NULL, NULL, NULL, NULL, NULL, '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '" + CountTXT.Text + "', '1', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", ConexaoDados.GetConnectionEquatorial());
                            prompt_cmd.ExecuteNonQuery();
                        }
                        if (ValorIcms == 25)
                        {
                            MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_boleto(material, desc_item, qtd, centro, custo, cod_imp, base_calculo, vl_icms, txt_pedido, material_dif, desc_item_dif, qtd_dif, centro_dif, custo_dif, cod_imp_dif, vl_dif, iva_dif, emissao, nfe, err, err_col, txt_miro, data_venc, Mes_ref, cod_unico, pedido, migo, miro, fecoep, valor_miro, status, empresa, mes_dupl, vl_dupl, now_date) VALUES ('271199',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','PH', '" + vl_boleto.Text.Replace("R$ ", "") + "',  '0', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', NULL, NULL, NULL, NULL, NULL, NULL, NULL, '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '" + CountTXT.Text + "',  '1', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", ConexaoDados.GetConnectionEquatorial());
                            prompt_cmd.ExecuteNonQuery();
                        }
                        if (ValorIcms == 27)
                        {
                            MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_boleto(material, desc_item, qtd, centro, custo, cod_imp, base_calculo, vl_icms, txt_pedido, material_dif, desc_item_dif, qtd_dif, centro_dif, custo_dif, cod_imp_dif, vl_dif, iva_dif, emissao, nfe, err, err_col, txt_miro, data_venc, Mes_ref, cod_unico, pedido, migo, miro, fecoep, valor_miro, status, empresa, mes_dupl, vl_dupl, now_date) VALUES ('271229',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','PH', '" + vl_boleto.Text.Replace("R$ ", "") + "',  '0', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', NULL, NULL, NULL, NULL, NULL, NULL, NULL, '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '" + CountTXT.Text + "', '1', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", ConexaoDados.GetConnectionEquatorial());
                            prompt_cmd.ExecuteNonQuery();
                        }
                        if (ValorIcms == 0)
                        {
                            MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_boleto(material, desc_item, qtd, centro, custo, cod_imp, base_calculo, vl_icms, txt_pedido, material_dif, desc_item_dif, qtd_dif, centro_dif, custo_dif, cod_imp_dif, vl_dif, iva_dif, emissao, nfe, err, err_col, txt_miro, data_venc, Mes_ref, cod_unico, pedido, migo, miro, fecoep, valor_miro, status, empresa, mes_dupl, vl_dupl, now_date) VALUES ('270743',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','PH', '" + vl_boleto.Text.Replace("R$ ", "") + "',  '0', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', NULL, NULL, NULL, NULL, NULL, NULL, '0', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '" + CountTXT.Text + "',  '1', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", ConexaoDados.GetConnectionEquatorial());
                            prompt_cmd.ExecuteNonQuery();
                        }
                    }
                    catch (Exception err)
                    {
                        MessageBox.Show(err.Message);
                    }
                    finally
                    {
                        ConexaoDados.GetConnectionEquatorial().Close();
                    }
                }
                else
                {
                    try
                    {
                        double vlDifTotal = Convert.ToDouble(vl_boleto.Text.Replace("R$ ", "")) - Convert.ToDouble(vl_base.Text.Replace("R$ ", ""));
                        if (ValorIcms == 17)
                        {
                            MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_boleto(material, desc_item, qtd, centro, custo, cod_imp, base_calculo, vl_icms, txt_pedido, material_dif, desc_item_dif, qtd_dif, centro_dif, custo_dif, cod_imp_dif, vl_dif, iva_dif, emissao, nfe, err, err_col, txt_miro, data_venc, Mes_ref, cod_unico, pedido, migo, miro, fecoep, valor_miro, status, empresa, mes_dupl, vl_dupl, now_date) VALUES ('270743', '" + txtFaz.Text + "', '1', 'USGA', '" + CUSTO + "','P1', '" + vl_base.Text.Replace("R$ ", "") + "', '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 270743, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', 'PH', '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '" + CountTXT.Text + "',  '2', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", ConexaoDados.GetConnectionEquatorial());
                            prompt_cmd.ExecuteNonQuery();
                        }
                        if (ValorIcms == 18)
                        {
                            MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_boleto(material, desc_item, qtd, centro, custo, cod_imp, base_calculo, vl_icms, txt_pedido, material_dif, desc_item_dif, qtd_dif, centro_dif, custo_dif, cod_imp_dif, vl_dif, iva_dif, emissao, nfe, err, err_col, txt_miro, data_venc, Mes_ref, cod_unico, pedido, migo, miro, fecoep, valor_miro, status, empresa, mes_dupl, vl_dupl, now_date) VALUES ('272920',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','P1', '" + vl_base.Text.Replace("R$ ", "") + "',  '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 272920, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', 'PH', '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '" + CountTXT.Text + "', '2', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", ConexaoDados.GetConnectionEquatorial());
                            prompt_cmd.ExecuteNonQuery();
                        }
                        if (ValorIcms == 25)
                        {
                            MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_boleto(material, desc_item, qtd, centro, custo, cod_imp, base_calculo, vl_icms, txt_pedido, material_dif, desc_item_dif, qtd_dif, centro_dif, custo_dif, cod_imp_dif, vl_dif, iva_dif, emissao, nfe, err, err_col, txt_miro, data_venc, Mes_ref, cod_unico, pedido, migo, miro, fecoep, valor_miro, status, empresa, mes_dupl, vl_dupl, now_date) VALUES ('271199',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','P1', '" + vl_base.Text.Replace("R$ ", "") + "',  '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 271199, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', 'PH', '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '" + CountTXT.Text + "',  '2', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", ConexaoDados.GetConnectionEquatorial());
                            prompt_cmd.ExecuteNonQuery();
                        }
                        if (ValorIcms == 27)
                        {
                            MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_boleto(material, desc_item, qtd, centro, custo, cod_imp, base_calculo, vl_icms, txt_pedido, material_dif, desc_item_dif, qtd_dif, centro_dif, custo_dif, cod_imp_dif, vl_dif, iva_dif, emissao, nfe, err, err_col, txt_miro, data_venc, Mes_ref, cod_unico, pedido, migo, miro, fecoep, valor_miro, status, empresa, mes_dupl, vl_dupl, now_date) VALUES ('271229',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','P1', '" + vl_base.Text.Replace("R$ ", "") + "',  '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 271229, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', 'PH', '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '" + CountTXT.Text + "',  '2', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", ConexaoDados.GetConnectionEquatorial());
                            prompt_cmd.ExecuteNonQuery();
                        }
                        if (ValorIcms == 0)
                        {
                            MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_boleto(material, desc_item, qtd, centro, custo, cod_imp, base_calculo, vl_icms, txt_pedido, material_dif, desc_item_dif, qtd_dif, centro_dif, custo_dif, cod_imp_dif, vl_dif, iva_dif, emissao, nfe, err, err_col, txt_miro, data_venc, Mes_ref, cod_unico, pedido, migo, miro, fecoep, valor_miro, status, empresa, mes_dupl, vl_dupl, now_date) VALUES ('270743',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','PH', '" + vl_base.Text.Replace("R$ ", "") + "',  '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 270743, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', NULL, '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '" + CountTXT.Text + "',  '2', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", ConexaoDados.GetConnectionEquatorial());
                            prompt_cmd.ExecuteNonQuery();
                        }
                    }
                    catch (Exception err)
                    {
                        MessageBox.Show(err.Message);
                    }
                    finally
                    {
                        ConexaoDados.GetConnectionEquatorial().Close();
                    }
                }
            }
            else
            {
                try
                {
                    double vlDifTotal = Convert.ToDouble(textBase1.Text.Replace("R$ ", ""));
                    if (ValorIcms == 17)
                    {
                        try
                        {
                            MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_boleto(material, desc_item, qtd, centro, custo, cod_imp, base_calculo, vl_icms, txt_pedido, material_dif, desc_item_dif, qtd_dif, centro_dif, custo_dif, cod_imp_dif, vl_dif, iva_dif, emissao, nfe, err, err_col, txt_miro, data_venc, Mes_ref, cod_unico, pedido, migo, miro, fecoep, valor_miro, status, empresa, mes_dupl, vl_dupl, now_date) VALUES ('270743', '" + txtFaz.Text + "', '1', 'USGA', '" + CUSTO + "','P1', '" + vl_base.Text.Replace("R$ ", "") + "', '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 270743, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', 'PH', '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '" + CountTXT.Text + "',  '2', '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, 'Fecoep.: " + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", ConexaoDados.GetConnectionEquatorial());
                            prompt_cmd.ExecuteNonQuery();
                        }
                        catch (MySqlException Erro)
                        {
                            MessageBox.Show(Erro.Message);
                        }
                    }
                    if (ValorIcms == 18)
                    {
                        MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_boleto(material, desc_item, qtd, centro, custo, cod_imp, base_calculo, vl_icms, txt_pedido, material_dif, desc_item_dif, qtd_dif, centro_dif, custo_dif, cod_imp_dif, vl_dif, iva_dif, emissao, nfe, err, err_col, txt_miro, data_venc, Mes_ref, cod_unico, pedido, migo, miro, fecoep, valor_miro, status, empresa, mes_dupl, vl_dupl, now_date) VALUES (NULL, '272920',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','P1', '" + vl_base.Text.Replace("R$ ", "") + "',  '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 272920, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', 'PH', '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '" + CountTXT.Text + "', NULL, '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, '" + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", ConexaoDados.GetConnectionEquatorial());
                        prompt_cmd.ExecuteNonQuery();
                    }
                    if (ValorIcms == 25)
                    {
                        MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_boleto(material, desc_item, qtd, centro, custo, cod_imp, base_calculo, vl_icms, txt_pedido, material_dif, desc_item_dif, qtd_dif, centro_dif, custo_dif, cod_imp_dif, vl_dif, iva_dif, emissao, nfe, err, err_col, txt_miro, data_venc, Mes_ref, cod_unico, pedido, migo, miro, fecoep, valor_miro, status, empresa, mes_dupl, vl_dupl, now_date) VALUES (NULL, '271199',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','P1', '" + vl_base.Text.Replace("R$ ", "") + "',  '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 271199, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', 'PH', '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '" + CountTXT.Text + "',  NULL, '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, '" + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", ConexaoDados.GetConnectionEquatorial());
                        prompt_cmd.ExecuteNonQuery();
                    }
                    if (ValorIcms == 27)
                    {
                        MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_boleto(material, desc_item, qtd, centro, custo, cod_imp, base_calculo, vl_icms, txt_pedido, material_dif, desc_item_dif, qtd_dif, centro_dif, custo_dif, cod_imp_dif, vl_dif, iva_dif, emissao, nfe, err, err_col, txt_miro, data_venc, Mes_ref, cod_unico, pedido, migo, miro, fecoep, valor_miro, status, empresa, mes_dupl, vl_dupl, now_date) VALUES (NULL, '271229',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','P1', '" + vl_base.Text.Replace("R$ ", "") + "',  '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 271229, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', 'PH', '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '" + CountTXT.Text + "',  NULL, '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, '" + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", ConexaoDados.GetConnectionEquatorial());
                        prompt_cmd.ExecuteNonQuery();
                    }
                    if (ValorIcms == 0)
                    {
                        MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_boleto(material, desc_item, qtd, centro, custo, cod_imp, base_calculo, vl_icms, txt_pedido, material_dif, desc_item_dif, qtd_dif, centro_dif, custo_dif, cod_imp_dif, vl_dif, iva_dif, emissao, nfe, err, err_col, txt_miro, data_venc, Mes_ref, cod_unico, pedido, migo, miro, fecoep, valor_miro, status, empresa, mes_dupl, vl_dupl, now_date) VALUES (NULL, '270743',  '" + txtFaz.Text + "', '1', 'USGA',  '" + CUSTO + "','PH', '" + vl_base.Text.Replace("R$ ", "") + "',  '" + ValorIcms + "', 'Ref. Nota Fiscal Nº:" + nfe.Text + " de " + this.dataemissao.Text + "', 270743, '" + txtFaz.Text + " - Dif. ICMS', '1', 'USGA', '" + CUSTO + "', NULL, '" + vlDifTotal + "', '0', '" + this.dataemissao.Text + "', '" + nfe.Text + "', '" + CountTXT.Text + "', NULL, '" + txtFaz.Text + "', '" + this.datavencimento.Text + "', '" + mes_nf.Text + "', '" + cod_unico.Text + "', NULL, NULL, NULL, '" + vl_fecoep.Text + "', '" + vl_boleto.Text.Replace("R$ ", "") + "', '" + STATUS + "', '" + EMPRESA + "', '" + txtMesdupl.Text + "', '" + txtValordupl.Text.Replace("R$ ", "") + "', NOW())", ConexaoDados.GetConnectionEquatorial());
                        prompt_cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message);
                }
                finally
                {
                    ConexaoDados.GetConnectionEquatorial().Close();
                }
            }
            if (string.IsNullOrWhiteSpace(vl_multa.Text))
            {

            }
            else
            {
                MySqlCommand cmd = new MySqlCommand("INSERT INTO TB_MULTA (cod,mes,valor,empresa) VALUES ('" + cod_unico.Text + "','" + mesMulta.Text + "','" + vl_multa.Text.Replace("R$ ", "") + "','" + EMPRESA + "')", ConexaoDados.GetConnectionEquatorial());
                cmd.ExecuteNonQuery();
            }

            //Fechar Conexão
            ConexaoDados.GetConnectionEquatorial().Close();

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
        private void Ceal_Load(object sender, EventArgs e)
        {
            txtUltNFE.Enabled = false;
            try
            {

                MySqlCommand com = new MySqlCommand();
                com.Connection = ConexaoDados.GetConnectionEquatorial();
                com.CommandText = "SELECT * FROM tb_boleto WHERE err > 0";
                MySqlDataReader dr = com.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(dr);
                CountTXT.Text = dt.Rows.Count.ToString();
                int Contagem = Convert.ToInt32(CountTXT.Text);
                Contagem++;
                CountTXT.Text = Contagem.ToString();
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
            }
            finally
            {
                ConexaoDados.GetConnectionEquatorial().Close();
            }

            try
            {


                MySqlCommand com = new MySqlCommand();
                com.Connection = ConexaoDados.GetConnectionEquatorial();
                com.CommandText = "SELECT porcentagem FROM TB_ICMS";
                MySqlDataReader dr = com.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(dr);
                preencherCBIcms.DisplayMember = "porcentagem";
                preencherCBIcms.DataSource = dt;
                ConexaoDados.GetConnectionEquatorial().Close();
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
            }
            finally
            {
                ConexaoDados.GetConnectionEquatorial().Close();
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
            if (MessageBox.Show("Deseja Voltar?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                frm_Main back = new frm_Main();
                back.Show();
                Close();
            }
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
        private void txtPedido_DoubleClick(object sender, EventArgs e)
        {
            MetroMessageBox.Show(this, "Your message here.", "Title Here", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void txtMigo_DoubleClick(object sender, EventArgs e)
        {
            MetroMessageBox.Show(this, "Your message here.", "Title Here", MessageBoxButtons.OK, MessageBoxIcon.Stop);
        }
        private void txtMiro_DoubleClick(object sender, EventArgs e)
        {
            MetroMessageBox.Show(this, "Your message here.", "Title Here", MessageBoxButtons.OK, MessageBoxIcon.Question);
        }

        private void txTexte_Click(object sender, EventArgs e)
        {
            try
            {
                ItensPedido();
            }
            catch (Exception erado)
            {
                MessageBox.Show(erado.Message);
            }
            finally
            {
                ConexaoDados.GetConnectionEquatorial().Close();
            }
        }

        private void cod_unico_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    //Consulta Itens no Banco de Dados
                    MySqlCommand MyCommand = new MySqlCommand();
                    MyCommand.Connection = ConexaoDados.GetConnectionEquatorial();
                    MyCommand.CommandText = "SELECT * FROM tb_boleto WHERE cod_unico='" + cod_unico.Text + "'";
                    MySqlDataReader dreader = MyCommand.ExecuteReader();
                    while (dreader.Read())
                    {
                        txtFaz.Text = dreader[2].ToString();
                    }
                }
                catch (Exception Err)
                {
                    txtFaz.Focus();
                }
                finally
                {
                    ConexaoDados.GetConnectionEquatorial().Close();
                }
            }
        }

        private void txtFaz_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                nfe.Focus();
            }
        }

        private void nfe_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                mes_nf.Focus();
            }
        }

        private void mes_nf_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                dataemissao.Focus();
            }
        }

        private void dataemissao_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                datavencimento.Focus();
            }
        }

        private void datavencimento_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                vl_fecoep.Focus();
            }
        }

        private void vl_fecoep_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                vl_boleto.Focus();
            }
        }

        private void vl_boleto_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                vl_base.Focus();
            }
        }

        private void vl_base_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnGravar.Focus();
            }
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

        private void cod_unico_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
