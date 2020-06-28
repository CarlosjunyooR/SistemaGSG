﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using Interop.SAPFEWSELib;
using Interop.SapROTWr;
using java.util.concurrent;

namespace SistemaGSG
{
    public partial class FormNotaFiscal : MetroFramework.Forms.MetroForm
    {
        MySqlConnection cn;
        MySqlCommand cmd_item1;
        public FormNotaFiscal()
        {
            InitializeComponent();
        }
        public void Conexao()
        {
            cn = new MySqlConnection(@"server='" + txtHost.Text + "';database='" + txtDataBase.Text + "';Uid='" + txtUser.Text + "';Pwd='" + txtPass.Text + "';SslMode=none;");
            cn.Open();
        }
        private void btnAddItem_Click(object sender, EventArgs e)
        {

        }
        private void label6_Click(object sender, EventArgs e)
        {

        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkItem1.Checked)
            {
                //Ativa o Proximo Item
                lblItem2.Visible = true;
                Item2.Visible = true;
                DescItem2.Visible = true;
                checkItem2.Visible = true;

                //Ativa o Foco para o proximo item
                Item2.Focus();

                //Verifica Cadastro do Item
                if (Item1.Text == "334")
                {
                    DescItem1.Text = "ARROZ TIO VIERA PARBORIZADO 1KG";
                }
                else
                {
                    DescItem1.Text = "Item Não Cadastrado No Sistema!";
                }
            }
            else
            {
                //Desativa o Item Anterior
                lblItem2.Visible = false;
                Item2.Visible = false;
                DescItem2.Visible = false;
                checkItem2.Visible = false;

                //Limpa os Campos
                DescItem1.Text = "";
                Item1.Text = "";
            }
        }
        private void checkItem2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkItem2.Checked)
            {
                //Bloqueia para alteração
                Item1.Enabled = false;
                checkItem1.Enabled = false;

                //Ativa o Foco para o proximo item
                Item3.Focus();

                //Ativa o Proximo Item
                lblItem3.Visible = true;
                Item3.Visible = true;
                DescItem3.Visible = true;
                checkItem3.Visible = true;

                //Verifica Cadastro do Item
                if (Item2.Text == "68")
                {
                    DescItem2.Text = "MACARRAO ESPAGUETE MAURICEA 500G";
                }
                else
                {
                    DescItem2.Text = "Item Não Cadastrado No Sistema!";
                }

            }
            else
            {
                //Ativa para alteração
                Item1.Enabled = true;
                checkItem1.Enabled = true;

                //Desativa o Item Anterior
                lblItem3.Visible = false;
                Item3.Visible = false;
                DescItem3.Visible = false;
                checkItem3.Visible = false;

                //Limpa os Campos
                DescItem2.Text = "";
                Item2.Text = "";
            }
        }
        private void checkItem3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkItem3.Checked)
            {
                //Bloqueia para alteração
                Item2.Enabled = false;
                checkItem2.Enabled = false;

                //Ativa o Foco para o proximo item
                Item4.Focus();

                //Ativa o Proximo Item
                lblItem4.Visible = true;
                Item4.Visible = true;
                DescItem4.Visible = true;
                checkItem4.Visible = true;

                //Verifica Cadastro do Item
                if (Item3.Text == "260")
                {
                    DescItem3.Text = "SOJA";
                }
                else
                {
                    DescItem3.Text = "Item Não Cadastrado No Sistema!";
                }

            }
            else
            {
                //Ativa para alteração
                Item2.Enabled = true;
                checkItem2.Enabled = true;

                //Desativa o Item Anterior
                lblItem4.Visible = false;
                Item4.Visible = false;
                DescItem4.Visible = false;
                checkItem4.Visible = false;

                //Limpa os Campos
                DescItem3.Text = "";
                Item3.Text = "";
            }
        }
        private void checkItem4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkItem4.Checked)
            {
                //Bloqueia para alteração
                Item3.Enabled = false;
                checkItem3.Enabled = false;

                //Ativa o Foco para o proximo item
                Item5.Focus();

                //Ativa o Proximo Item
                lblItem5.Visible = true;
                Item5.Visible = true;
                DescItem5.Visible = true;
                checkItem5.Visible = true;

                //Verifica Cadastro do Item
                if (Item4.Text == "261")
                {
                    DescItem4.Text = "CHARQUE PA BANDEIRANTE";
                }
                else
                {
                    DescItem4.Text = "Item Não Cadastrado No Sistema!";
                }

            }
            else
            {
                //Ativa para alteração
                Item3.Enabled = true;
                checkItem3.Enabled = true;

                //Desativa o Item Anterior
                lblItem5.Visible = false;
                Item5.Visible = false;
                DescItem5.Visible = false;
                checkItem5.Visible = false;

                //Limpa os Campos
                DescItem4.Text = "";
                Item4.Text = "";
            }
        }
        private void checkItem5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkItem5.Checked)
            {
                //Bloqueia para alteração
                Item4.Enabled = false;
                checkItem4.Enabled = false;

                //Ativa o Foco para o proximo item
                Item6.Focus();

                //Ativa o Proximo Item
                lblItem6.Visible = true;
                Item6.Visible = true;
                DescItem6.Visible = true;
                checkItem6.Visible = true;

                //Verifica Cadastro do Item
                if (Item5.Text == "264")
                {
                    DescItem5.Text = "FARINHA MANDIOCA";
                }
                else
                {
                    DescItem5.Text = "Item Não Cadastrado No Sistema!";
                }

            }
            else
            {
                //Ativa para alteração
                Item4.Enabled = true;
                checkItem4.Enabled = true;

                //Desativa o Item Anterior
                lblItem6.Visible = false;
                Item6.Visible = false;
                DescItem6.Visible = false;
                checkItem6.Visible = false;

                //Limpa os Campos
                DescItem5.Text = "";
                Item5.Text = "";
            }
        }
        private void checkItem6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkItem6.Checked)
            {
                //Bloqueia para alteração
                Item5.Enabled = false;
                checkItem5.Enabled = false;

                //Ativa o Foco para o proximo item
                Item7.Focus();

                //Ativa o Proximo Item
                lblItem7.Visible = true;
                Item7.Visible = true;
                DescItem7.Visible = true;
                checkItem7.Visible = true;

                //Verifica Cadastro do Item
                if (Item6.Text == "266")
                {
                    DescItem6.Text = "MARGARINA 500G";
                }
                else
                {
                    DescItem6.Text = "Item Não Cadastrado No Sistema!";
                }

            }
            else
            {
                //Ativa para alteração
                Item5.Enabled = true;
                checkItem5.Enabled = true;

                //Desativa o Item Anterior
                lblItem7.Visible = false;
                Item7.Visible = false;
                DescItem7.Visible = false;
                checkItem7.Visible = false;

                //Limpa os Campos
                DescItem6.Text = "";
                Item6.Text = "";
            }
        }
        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
        private void checkItem7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkItem7.Checked)
            {
                //Bloqueia para alteração
                Item6.Enabled = false;
                checkItem6.Enabled = false;

                //Ativa o Foco para o proximo item
                Item8.Focus();

                //Ativa o Proximo Item
                lblItem8.Visible = true;
                Item8.Visible = true;
                DescItem8.Visible = true;
                checkItem8.Visible = true;

                //Verifica Cadastro do Item
                if (Item7.Text == "66")
                {
                    DescItem7.Text = "OLEO SOYA";
                }
                else
                {
                    DescItem7.Text = "Item Não Cadastrado No Sistema!";
                }

            }
            else
            {
                //Ativa para alteração
                Item6.Enabled = true;
                checkItem6.Enabled = true;

                //Desativa o Item Anterior
                lblItem8.Visible = false;
                Item8.Visible = false;
                DescItem8.Visible = false;
                checkItem8.Visible = false;

                //Limpa os Campos
                DescItem7.Text = "";
                Item7.Text = "";
            }
        }
        private void checkItem8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkItem8.Checked)
            {
                //Bloqueia para alteração
                Item7.Enabled = false;
                checkItem7.Enabled = false;

                //Ativa o Foco para o proximo item
                Item9.Focus();

                //Ativa o Proximo Item
                lblItem9.Visible = true;
                Item9.Visible = true;
                DescItem9.Visible = true;
                checkItem9.Visible = true;

                //Verifica Cadastro do Item
                if (Item8.Text == "271")
                {
                    DescItem8.Text = "OVOS";
                }
                else
                {
                    DescItem8.Text = "Item Não Cadastrado No Sistema!";
                }

            }
            else
            {
                //Ativa para alteração
                Item7.Enabled = true;
                checkItem7.Enabled = true;

                //Desativa o Item Anterior
                lblItem9.Visible = false;
                Item9.Visible = false;
                DescItem9.Visible = false;
                checkItem9.Visible = false;

                //Limpa os Campos
                DescItem8.Text = "";
                Item8.Text = "";
            }
        }
        private void checkItem9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkItem9.Checked)
            {
                //Bloqueia para alteração
                Item8.Enabled = false;
                checkItem8.Enabled = false;

                //Ativa o Foco para o proximo item
                Item10.Focus();

                //Ativa o Proximo Item
                lblItem10.Visible = true;
                Item10.Visible = true;
                DescItem10.Visible = true;
                checkItem10.Visible = true;

                //Verifica Cadastro do Item
                if (Item9.Text == "258")
                {
                    DescItem9.Text = "TOMATE";
                }
                else
                {
                    DescItem9.Text = "Item Não Cadastrado No Sistema!";
                }

            }
            else
            {
                //Ativa para alteração
                Item8.Enabled = true;
                checkItem8.Enabled = true;

                //Desativa o Item Anterior
                lblItem10.Visible = false;
                Item10.Visible = false;
                DescItem10.Visible = false;
                checkItem10.Visible = false;

                //Limpa os Campos
                DescItem9.Text = "";
                Item9.Text = "";
            }
        }
        private void checkItem10_CheckedChanged(object sender, EventArgs e)
        {
            if (checkItem10.Checked)
            {
                //Bloqueia para alteração
                Item9.Enabled = false;
                checkItem9.Enabled = false;

                //Ativa o Foco para o proximo item
                Item11.Focus();

                //Ativa o Proximo Item
                lblItem11.Visible = true;
                Item11.Visible = true;
                DescItem11.Visible = true;
                checkItem11.Visible = true;

                //Verifica Cadastro do Item
                if (Item10.Text == "259")
                {
                    DescItem10.Text = "CEBOLA";
                }
                else
                {
                    DescItem10.Text = "Item Não Cadastrado No Sistema!";
                }

            }
            else
            {
                //Ativa para alteração
                Item9.Enabled = true;
                checkItem9.Enabled = true;

                //Desativa o Item Anterior
                lblItem11.Visible = false;
                Item11.Visible = false;
                DescItem11.Visible = false;
                checkItem11.Visible = false;

                //Limpa os Campos
                DescItem10.Text = "";
                Item10.Text = "";
            }
        }
        private void btnConsult_Click(object sender, EventArgs e)
        {
            cn = new MySqlConnection(@"server='" + txtHost.Text + "';database='" + txtDataBase.Text + "';Uid='" + txtUser.Text + "';Pwd='" + txtPass.Text + "';SslMode=none;");

            if (string.IsNullOrEmpty(Item1.Text))
            {
                DescItem1.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item1.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem1.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item2.Text))
            {
                DescItem2.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item2.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem2.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item3.Text))
            {
                DescItem3.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item3.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem3.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item4.Text))
            {
                DescItem4.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item4.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem4.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item5.Text))
            {
                DescItem5.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item5.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem5.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item6.Text))
            {
                DescItem6.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item6.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem6.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item7.Text))
            {
                DescItem7.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item7.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem7.Text = dreader[1].ToString();
                    break;
                }
                cn.Close(); 
            }
            if (string.IsNullOrEmpty(Item8.Text))
            {
                DescItem8.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item8.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem8.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item9.Text))
            {
                DescItem9.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item9.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem9.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item10.Text))
            {
                DescItem10.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item10.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem10.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item11.Text))
            {
                DescItem11.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item11.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem11.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item12.Text))
            {
                DescItem12.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item12.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem12.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item13.Text))
            {
                DescItem13.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item13.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem13.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item14.Text))
            {
                DescItem14.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item14.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem14.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item15.Text))
            {
                DescItem15.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item15.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem15.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item16.Text))
            {
                DescItem16.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item16.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem16.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item17.Text))
            {
                DescItem17.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item17.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem17.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item18.Text))
            {
                DescItem18.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item18.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem18.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item19.Text))
            {
                DescItem19.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item19.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem19.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item20.Text))
            {
                DescItem20.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item20.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem20.Text = dreader[1].ToString();
                    break;
                }
                cn.Close(); 
            }
            if (string.IsNullOrEmpty(Item21.Text))
            {
                DescItem21.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item21.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem21.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item22.Text))
            {
                DescItem22.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item22.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem22.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item23.Text))
            {
                DescItem23.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item23.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem23.Text = dreader[1].ToString();
                    break;
                }
                cn.Close(); 
            }
            if (string.IsNullOrEmpty(Item24.Text))
            {
                DescItem24.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item24.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem24.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item25.Text))
            {
                DescItem25.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item25.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem25.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item26.Text))
            {
                DescItem26.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item26.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem26.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item27.Text))
            {
                DescItem27.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item27.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem27.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item28.Text))
            {
                DescItem28.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item28.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem28.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item29.Text))
            {
                DescItem29.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item29.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem29.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item30.Text))
            {
                DescItem30.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item30.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem30.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item31.Text))
            {
                DescItem31.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item31.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem31.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item32.Text))
            {
                DescItem32.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item32.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem32.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();  
            }
            if (string.IsNullOrEmpty(Item33.Text))
            {
                DescItem33.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item33.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem33.Text = dreader[1].ToString();
                    break;
                }
                cn.Close(); 
            }
            if (string.IsNullOrEmpty(Item34.Text))
            {
                DescItem34.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item34.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem34.Text = dreader[1].ToString();
                    break;
                }
                cn.Close(); 
            }
            if (string.IsNullOrEmpty(Item35.Text))
            {
                DescItem35.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item35.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem35.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();  
            }
            if (string.IsNullOrEmpty(Item36.Text))
            {
                DescItem36.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item36.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem36.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item37.Text))
            {
                DescItem37.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item37.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem37.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item38.Text))
            {
                DescItem38.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item38.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem38.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item39.Text))
            {
                DescItem39.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item39.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem39.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
            if (string.IsNullOrEmpty(Item40.Text))
            {
                DescItem40.Text = "";
            }
            else
            {
                cn.Open();
                //Consulta Itens no Banco de Dados
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = cn;
                MyCommand.CommandText = "SELECT * FROM tb_produtos WHERE CD_PRODUTO='" + Item40.Text + "'";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    DescItem40.Text = dreader[1].ToString();
                    break;
                }
                cn.Close();
            }
        }
        private void checkItem11_CheckedChanged(object sender, EventArgs e)
        {
            if (checkItem11.Checked)
            {
                //Bloqueia para alteração
                Item10.Enabled = false;
                checkItem10.Enabled = false;

                //Ativa o Foco para o proximo item
                Item12.Focus();

                //Ativa o Proximo Item
                lblItem12.Visible = true;
                Item12.Visible = true;
                DescItem12.Visible = true;
                checkItem12.Visible = true;

                //Verifica Cadastro do Item
                if (Item11.Text == "262")
                {
                    DescItem11.Text = "ALHO KG";
                }
                else
                {
                    DescItem11.Text = "Item Não Cadastrado No Sistema!";
                }

            }
            else
            {
                //Ativa para alteração
                Item10.Enabled = true;
                checkItem10.Enabled = true;

                //Desativa o Item Anterior
                lblItem12.Visible = false;
                Item12.Visible = false;
                DescItem12.Visible = false;
                checkItem12.Visible = false;

                //Limpa os Campos
                DescItem11.Text = "";
                Item11.Text = "";
            }
        }
        private void checkItem12_CheckedChanged(object sender, EventArgs e)
        {
            if (checkItem12.Checked)
            {
                //Bloqueia para alteração
                Item11.Enabled = false;
                checkItem11.Enabled = false;

                //Ativa o Foco para o proximo item
                Item11.Focus();

                //Ativa o Proximo Item
                lblItem13.Visible = true;
                Item13.Visible = true;
                DescItem13.Visible = true;
                checkItem13.Visible = true;

                //Verifica Cadastro do Item
                if (Item12.Text == "694")
                {
                    DescItem12.Text = "FARINHA DE TRIGO C/ FERMENTO ROSA BRANCA 1KG";
                }
                else
                {
                    DescItem12.Text = "Item Não Cadastrado No Sistema!";
                }

            }
            else
            {
                //Ativa para alteração
                Item11.Enabled = true;
                checkItem11.Enabled = true;

                //Desativa o Item Anterior
                lblItem13.Visible = false;
                Item13.Visible = false;
                DescItem13.Visible = false;
                checkItem13.Visible = false;

                //Limpa os Campos
                DescItem12.Text = "";
                Item12.Text = "";
            }
        }
        private void btnAddItem_Click_1(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtQTDItensNF.Text))
            {
                MessageBox.Show("Nenhum Item!");
            }
            if (txtQTDItensNF.Text == "1")
            {
                MessageBox.Show("Box Já Adicionado!");
            }
            if (txtQTDItensNF.Text == "2")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;
            }
            if (txtQTDItensNF.Text == "3")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;
            }
            if (txtQTDItensNF.Text == "4")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;
            }
            if (txtQTDItensNF.Text == "5")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;
            }
            if (txtQTDItensNF.Text == "6")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;
            }
            if (txtQTDItensNF.Text == "7")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;
            }
            if (txtQTDItensNF.Text == "8")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;
            }
            if (txtQTDItensNF.Text == "9")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;
            }
            if (txtQTDItensNF.Text == "10")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;
            }
            if (txtQTDItensNF.Text == "11")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;
            }
            if (txtQTDItensNF.Text == "12")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;
            }
            if (txtQTDItensNF.Text == "13")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;
            }
            if (txtQTDItensNF.Text == "14")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;
            }
            if (txtQTDItensNF.Text == "15")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;
            }
            if (txtQTDItensNF.Text == "16")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;
            }
            if (txtQTDItensNF.Text == "17")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;
            }
            if (txtQTDItensNF.Text == "18")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;
            }
            if (txtQTDItensNF.Text == "19")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;
            }
            if (txtQTDItensNF.Text == "20")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;
            }
            if (txtQTDItensNF.Text == "21")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;
            }
            if (txtQTDItensNF.Text == "22")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;
            }
            if (txtQTDItensNF.Text == "23")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;

                lblItem23.Visible = true;
                Item23.Visible = true;
                qdt23.Visible = true;
                preco23.Visible = true;
                DescItem23.Visible = true;
            }
            if (txtQTDItensNF.Text == "24")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;

                lblItem23.Visible = true;
                Item23.Visible = true;
                qdt23.Visible = true;
                preco23.Visible = true;
                DescItem23.Visible = true;

                lblItem24.Visible = true;
                Item24.Visible = true;
                qdt24.Visible = true;
                preco24.Visible = true;
                DescItem24.Visible = true;
            }
            if (txtQTDItensNF.Text == "25")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;

                lblItem23.Visible = true;
                Item23.Visible = true;
                qdt23.Visible = true;
                preco23.Visible = true;
                DescItem23.Visible = true;

                lblItem24.Visible = true;
                Item24.Visible = true;
                qdt24.Visible = true;
                preco24.Visible = true;
                DescItem24.Visible = true;

                lblItem25.Visible = true;
                Item25.Visible = true;
                qdt25.Visible = true;
                DescItem25.Visible = true;
            }
            if (txtQTDItensNF.Text == "26")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;

                lblItem23.Visible = true;
                Item23.Visible = true;
                qdt23.Visible = true;
                preco23.Visible = true;
                DescItem23.Visible = true;

                lblItem24.Visible = true;
                Item24.Visible = true;
                qdt24.Visible = true;
                preco24.Visible = true;
                DescItem24.Visible = true;

                lblItem25.Visible = true;
                Item25.Visible = true;
                qdt25.Visible = true;
                DescItem25.Visible = true;

                lblItem26.Visible = true;
                Item26.Visible = true;
                qdt26.Visible = true;
                preco26.Visible = true;
                DescItem26.Visible = true;
            }
            if (txtQTDItensNF.Text == "27")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;

                lblItem23.Visible = true;
                Item23.Visible = true;
                qdt23.Visible = true;
                preco23.Visible = true;
                DescItem23.Visible = true;

                lblItem24.Visible = true;
                Item24.Visible = true;
                qdt24.Visible = true;
                preco24.Visible = true;
                DescItem24.Visible = true;

                lblItem25.Visible = true;
                Item25.Visible = true;
                qdt25.Visible = true;
                DescItem25.Visible = true;

                lblItem26.Visible = true;
                Item26.Visible = true;
                qdt26.Visible = true;
                preco26.Visible = true;
                DescItem26.Visible = true;

                lblItem27.Visible = true;
                Item27.Visible = true;
                qdt27.Visible = true;
                preco27.Visible = true;
                DescItem27.Visible = true;
            }
            if (txtQTDItensNF.Text == "28")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;

                lblItem23.Visible = true;
                Item23.Visible = true;
                qdt23.Visible = true;
                preco23.Visible = true;
                DescItem23.Visible = true;

                lblItem24.Visible = true;
                Item24.Visible = true;
                qdt24.Visible = true;
                preco24.Visible = true;
                DescItem24.Visible = true;

                lblItem25.Visible = true;
                Item25.Visible = true;
                qdt25.Visible = true;
                DescItem25.Visible = true;

                lblItem26.Visible = true;
                Item26.Visible = true;
                qdt26.Visible = true;
                preco26.Visible = true;
                DescItem26.Visible = true;

                lblItem27.Visible = true;
                Item27.Visible = true;
                qdt27.Visible = true;
                preco27.Visible = true;
                DescItem27.Visible = true;

                lblItem28.Visible = true;
                Item28.Visible = true;
                qdt28.Visible = true;
                preco28.Visible = true;
                DescItem28.Visible = true;
            }
            if (txtQTDItensNF.Text == "29")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;

                lblItem23.Visible = true;
                Item23.Visible = true;
                qdt23.Visible = true;
                preco23.Visible = true;
                DescItem23.Visible = true;

                lblItem24.Visible = true;
                Item24.Visible = true;
                qdt24.Visible = true;
                preco24.Visible = true;
                DescItem24.Visible = true;

                lblItem25.Visible = true;
                Item25.Visible = true;
                qdt25.Visible = true;
                DescItem25.Visible = true;

                lblItem26.Visible = true;
                Item26.Visible = true;
                qdt26.Visible = true;
                preco26.Visible = true;
                DescItem26.Visible = true;

                lblItem27.Visible = true;
                Item27.Visible = true;
                qdt27.Visible = true;
                preco27.Visible = true;
                DescItem27.Visible = true;

                lblItem28.Visible = true;
                Item28.Visible = true;
                qdt28.Visible = true;
                preco28.Visible = true;
                DescItem28.Visible = true;

                lblItem29.Visible = true;
                Item29.Visible = true;
                qdt29.Visible = true;
                preco29.Visible = true;
                DescItem29.Visible = true;
            }
            if (txtQTDItensNF.Text == "30")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;

                lblItem23.Visible = true;
                Item23.Visible = true;
                qdt23.Visible = true;
                preco23.Visible = true;
                DescItem23.Visible = true;

                lblItem24.Visible = true;
                Item24.Visible = true;
                qdt24.Visible = true;
                preco24.Visible = true;
                DescItem24.Visible = true;

                lblItem25.Visible = true;
                Item25.Visible = true;
                qdt25.Visible = true;
                DescItem25.Visible = true;

                lblItem26.Visible = true;
                Item26.Visible = true;
                qdt26.Visible = true;
                preco26.Visible = true;
                DescItem26.Visible = true;

                lblItem27.Visible = true;
                Item27.Visible = true;
                qdt27.Visible = true;
                preco27.Visible = true;
                DescItem27.Visible = true;

                lblItem28.Visible = true;
                Item28.Visible = true;
                qdt28.Visible = true;
                preco28.Visible = true;
                DescItem28.Visible = true;

                lblItem29.Visible = true;
                Item29.Visible = true;
                qdt29.Visible = true;
                preco29.Visible = true;
                DescItem29.Visible = true;

                lblItem30.Visible = true;
                Item30.Visible = true;
                qdt30.Visible = true;
                preco30.Visible = true;
                DescItem30.Visible = true;
            }
            if (txtQTDItensNF.Text == "31")
            {

                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;

                lblItem23.Visible = true;
                Item23.Visible = true;
                qdt23.Visible = true;
                preco23.Visible = true;
                DescItem23.Visible = true;

                lblItem24.Visible = true;
                Item24.Visible = true;
                qdt24.Visible = true;
                preco24.Visible = true;
                DescItem24.Visible = true;

                lblItem25.Visible = true;
                Item25.Visible = true;
                qdt25.Visible = true;
                DescItem25.Visible = true;

                lblItem26.Visible = true;
                Item26.Visible = true;
                qdt26.Visible = true;
                preco26.Visible = true;
                DescItem26.Visible = true;

                lblItem27.Visible = true;
                Item27.Visible = true;
                qdt27.Visible = true;
                preco27.Visible = true;
                DescItem27.Visible = true;

                lblItem28.Visible = true;
                Item28.Visible = true;
                qdt28.Visible = true;
                preco28.Visible = true;
                DescItem28.Visible = true;

                lblItem29.Visible = true;
                Item29.Visible = true;
                qdt29.Visible = true;
                preco29.Visible = true;
                DescItem29.Visible = true;

                lblItem30.Visible = true;
                Item30.Visible = true;
                qdt30.Visible = true;
                preco30.Visible = true;
                DescItem30.Visible = true;

                lblItem31.Visible = true;
                Item31.Visible = true;
                qdt31.Visible = true;
                preco31.Visible = true;
                DescItem31.Visible = true;
            }
            if (txtQTDItensNF.Text == "32")
            {

                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;

                lblItem23.Visible = true;
                Item23.Visible = true;
                qdt23.Visible = true;
                preco23.Visible = true;
                DescItem23.Visible = true;

                lblItem24.Visible = true;
                Item24.Visible = true;
                qdt24.Visible = true;
                preco24.Visible = true;
                DescItem24.Visible = true;

                lblItem25.Visible = true;
                Item25.Visible = true;
                qdt25.Visible = true;
                DescItem25.Visible = true;

                lblItem26.Visible = true;
                Item26.Visible = true;
                qdt26.Visible = true;
                preco26.Visible = true;
                DescItem26.Visible = true;

                lblItem27.Visible = true;
                Item27.Visible = true;
                qdt27.Visible = true;
                preco27.Visible = true;
                DescItem27.Visible = true;

                lblItem28.Visible = true;
                Item28.Visible = true;
                qdt28.Visible = true;
                preco28.Visible = true;
                DescItem28.Visible = true;

                lblItem29.Visible = true;
                Item29.Visible = true;
                qdt29.Visible = true;
                preco29.Visible = true;
                DescItem29.Visible = true;

                lblItem30.Visible = true;
                Item30.Visible = true;
                qdt30.Visible = true;
                preco30.Visible = true;
                DescItem30.Visible = true;

                lblItem31.Visible = true;
                Item31.Visible = true;
                qdt31.Visible = true;
                preco31.Visible = true;
                DescItem31.Visible = true;

                lblItem32.Visible = true;
                Item32.Visible = true;
                qdt32.Visible = true;
                preco32.Visible = true;
                DescItem32.Visible = true;
            }
            if (txtQTDItensNF.Text == "33")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;

                lblItem23.Visible = true;
                Item23.Visible = true;
                qdt23.Visible = true;
                preco23.Visible = true;
                DescItem23.Visible = true;

                lblItem24.Visible = true;
                Item24.Visible = true;
                qdt24.Visible = true;
                preco24.Visible = true;
                DescItem24.Visible = true;

                lblItem25.Visible = true;
                Item25.Visible = true;
                qdt25.Visible = true;
                DescItem25.Visible = true;

                lblItem26.Visible = true;
                Item26.Visible = true;
                qdt26.Visible = true;
                preco26.Visible = true;
                DescItem26.Visible = true;

                lblItem27.Visible = true;
                Item27.Visible = true;
                qdt27.Visible = true;
                preco27.Visible = true;
                DescItem27.Visible = true;

                lblItem28.Visible = true;
                Item28.Visible = true;
                qdt28.Visible = true;
                preco28.Visible = true;
                DescItem28.Visible = true;

                lblItem29.Visible = true;
                Item29.Visible = true;
                qdt29.Visible = true;
                preco29.Visible = true;
                DescItem29.Visible = true;

                lblItem30.Visible = true;
                Item30.Visible = true;
                qdt30.Visible = true;
                preco30.Visible = true;
                DescItem30.Visible = true;

                lblItem31.Visible = true;
                Item31.Visible = true;
                qdt31.Visible = true;
                preco31.Visible = true;
                DescItem31.Visible = true;

                lblItem32.Visible = true;
                Item32.Visible = true;
                qdt32.Visible = true;
                preco32.Visible = true;
                DescItem32.Visible = true;

                lblItem33.Visible = true;
                Item33.Visible = true;
                qdt33.Visible = true;
                preco33.Visible = true;
                DescItem33.Visible = true;
            }
            if (txtQTDItensNF.Text == "34")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;

                lblItem23.Visible = true;
                Item23.Visible = true;
                qdt23.Visible = true;
                preco23.Visible = true;
                DescItem23.Visible = true;

                lblItem24.Visible = true;
                Item24.Visible = true;
                qdt24.Visible = true;
                preco24.Visible = true;
                DescItem24.Visible = true;

                lblItem25.Visible = true;
                Item25.Visible = true;
                qdt25.Visible = true;
                DescItem25.Visible = true;

                lblItem26.Visible = true;
                Item26.Visible = true;
                qdt26.Visible = true;
                preco26.Visible = true;
                DescItem26.Visible = true;

                lblItem27.Visible = true;
                Item27.Visible = true;
                qdt27.Visible = true;
                preco27.Visible = true;
                DescItem27.Visible = true;

                lblItem28.Visible = true;
                Item28.Visible = true;
                qdt28.Visible = true;
                preco28.Visible = true;
                DescItem28.Visible = true;

                lblItem29.Visible = true;
                Item29.Visible = true;
                qdt29.Visible = true;
                preco29.Visible = true;
                DescItem29.Visible = true;

                lblItem30.Visible = true;
                Item30.Visible = true;
                qdt30.Visible = true;
                preco30.Visible = true;
                DescItem30.Visible = true;

                lblItem31.Visible = true;
                Item31.Visible = true;
                qdt31.Visible = true;
                preco31.Visible = true;
                DescItem31.Visible = true;

                lblItem32.Visible = true;
                Item32.Visible = true;
                qdt32.Visible = true;
                preco32.Visible = true;
                DescItem32.Visible = true;

                lblItem33.Visible = true;
                Item33.Visible = true;
                qdt33.Visible = true;
                preco33.Visible = true;
                DescItem33.Visible = true;

                lblItem34.Visible = true;
                Item34.Visible = true;
                qdt34.Visible = true;
                preco34.Visible = true;
                DescItem34.Visible = true;
            }
            if (txtQTDItensNF.Text == "35")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;

                lblItem23.Visible = true;
                Item23.Visible = true;
                qdt23.Visible = true;
                preco23.Visible = true;
                DescItem23.Visible = true;

                lblItem24.Visible = true;
                Item24.Visible = true;
                qdt24.Visible = true;
                preco24.Visible = true;
                DescItem24.Visible = true;

                lblItem25.Visible = true;
                Item25.Visible = true;
                qdt25.Visible = true;
                DescItem25.Visible = true;

                lblItem26.Visible = true;
                Item26.Visible = true;
                qdt26.Visible = true;
                preco26.Visible = true;
                DescItem26.Visible = true;

                lblItem27.Visible = true;
                Item27.Visible = true;
                qdt27.Visible = true;
                preco27.Visible = true;
                DescItem27.Visible = true;

                lblItem28.Visible = true;
                Item28.Visible = true;
                qdt28.Visible = true;
                preco28.Visible = true;
                DescItem28.Visible = true;

                lblItem29.Visible = true;
                Item29.Visible = true;
                qdt29.Visible = true;
                preco29.Visible = true;
                DescItem29.Visible = true;

                lblItem30.Visible = true;
                Item30.Visible = true;
                qdt30.Visible = true;
                preco30.Visible = true;
                DescItem30.Visible = true;

                lblItem31.Visible = true;
                Item31.Visible = true;
                qdt31.Visible = true;
                preco31.Visible = true;
                DescItem31.Visible = true;

                lblItem32.Visible = true;
                Item32.Visible = true;
                qdt32.Visible = true;
                preco32.Visible = true;
                DescItem32.Visible = true;

                lblItem33.Visible = true;
                Item33.Visible = true;
                qdt33.Visible = true;
                preco33.Visible = true;
                DescItem33.Visible = true;

                lblItem34.Visible = true;
                Item34.Visible = true;
                qdt34.Visible = true;
                preco34.Visible = true;
                DescItem34.Visible = true;

                lblItem35.Visible = true;
                Item35.Visible = true;
                qdt35.Visible = true;
                preco35.Visible = true;
                DescItem35.Visible = true;
            }
            if (txtQTDItensNF.Text == "36")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;

                lblItem23.Visible = true;
                Item23.Visible = true;
                qdt23.Visible = true;
                preco23.Visible = true;
                DescItem23.Visible = true;

                lblItem24.Visible = true;
                Item24.Visible = true;
                qdt24.Visible = true;
                preco24.Visible = true;
                DescItem24.Visible = true;

                lblItem25.Visible = true;
                Item25.Visible = true;
                qdt25.Visible = true;
                DescItem25.Visible = true;

                lblItem26.Visible = true;
                Item26.Visible = true;
                qdt26.Visible = true;
                preco26.Visible = true;
                DescItem26.Visible = true;

                lblItem27.Visible = true;
                Item27.Visible = true;
                qdt27.Visible = true;
                preco27.Visible = true;
                DescItem27.Visible = true;

                lblItem28.Visible = true;
                Item28.Visible = true;
                qdt28.Visible = true;
                preco28.Visible = true;
                DescItem28.Visible = true;

                lblItem29.Visible = true;
                Item29.Visible = true;
                qdt29.Visible = true;
                preco29.Visible = true;
                DescItem29.Visible = true;

                lblItem30.Visible = true;
                Item30.Visible = true;
                qdt30.Visible = true;
                preco30.Visible = true;
                DescItem30.Visible = true;

                lblItem31.Visible = true;
                Item31.Visible = true;
                qdt31.Visible = true;
                preco31.Visible = true;
                DescItem31.Visible = true;

                lblItem32.Visible = true;
                Item32.Visible = true;
                qdt32.Visible = true;
                preco32.Visible = true;
                DescItem32.Visible = true;

                lblItem33.Visible = true;
                Item33.Visible = true;
                qdt33.Visible = true;
                preco33.Visible = true;
                DescItem33.Visible = true;

                lblItem34.Visible = true;
                Item34.Visible = true;
                qdt34.Visible = true;
                preco34.Visible = true;
                DescItem34.Visible = true;

                lblItem35.Visible = true;
                Item35.Visible = true;
                qdt35.Visible = true;
                preco35.Visible = true;
                DescItem35.Visible = true;

                lblItem36.Visible = true;
                Item36.Visible = true;
                qdt36.Visible = true;
                preco36.Visible = true;
                DescItem36.Visible = true;
            }
            if (txtQTDItensNF.Text == "37")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;

                lblItem23.Visible = true;
                Item23.Visible = true;
                qdt23.Visible = true;
                preco23.Visible = true;
                DescItem23.Visible = true;

                lblItem24.Visible = true;
                Item24.Visible = true;
                qdt24.Visible = true;
                preco24.Visible = true;
                DescItem24.Visible = true;

                lblItem25.Visible = true;
                Item25.Visible = true;
                qdt25.Visible = true;
                DescItem25.Visible = true;

                lblItem26.Visible = true;
                Item26.Visible = true;
                qdt26.Visible = true;
                preco26.Visible = true;
                DescItem26.Visible = true;

                lblItem27.Visible = true;
                Item27.Visible = true;
                qdt27.Visible = true;
                preco27.Visible = true;
                DescItem27.Visible = true;

                lblItem28.Visible = true;
                Item28.Visible = true;
                qdt28.Visible = true;
                preco28.Visible = true;
                DescItem28.Visible = true;

                lblItem29.Visible = true;
                Item29.Visible = true;
                qdt29.Visible = true;
                preco29.Visible = true;
                DescItem29.Visible = true;

                lblItem30.Visible = true;
                Item30.Visible = true;
                qdt30.Visible = true;
                preco30.Visible = true;
                DescItem30.Visible = true;

                lblItem31.Visible = true;
                Item31.Visible = true;
                qdt31.Visible = true;
                preco31.Visible = true;
                DescItem31.Visible = true;

                lblItem32.Visible = true;
                Item32.Visible = true;
                qdt32.Visible = true;
                preco32.Visible = true;
                DescItem32.Visible = true;

                lblItem33.Visible = true;
                Item33.Visible = true;
                qdt33.Visible = true;
                preco33.Visible = true;
                DescItem33.Visible = true;

                lblItem34.Visible = true;
                Item34.Visible = true;
                qdt34.Visible = true;
                preco34.Visible = true;
                DescItem34.Visible = true;

                lblItem35.Visible = true;
                Item35.Visible = true;
                qdt35.Visible = true;
                preco35.Visible = true;
                DescItem35.Visible = true;

                lblItem36.Visible = true;
                Item36.Visible = true;
                qdt36.Visible = true;
                preco36.Visible = true;
                DescItem36.Visible = true;

                lblItem37.Visible = true;
                Item37.Visible = true;
                qdt37.Visible = true;
                preco37.Visible = true;
                DescItem37.Visible = true;
            }
            if (txtQTDItensNF.Text == "38")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;

                lblItem23.Visible = true;
                Item23.Visible = true;
                qdt23.Visible = true;
                preco23.Visible = true;
                DescItem23.Visible = true;

                lblItem24.Visible = true;
                Item24.Visible = true;
                qdt24.Visible = true;
                preco24.Visible = true;
                DescItem24.Visible = true;

                lblItem25.Visible = true;
                Item25.Visible = true;
                qdt25.Visible = true;
                DescItem25.Visible = true;

                lblItem26.Visible = true;
                Item26.Visible = true;
                qdt26.Visible = true;
                preco26.Visible = true;
                DescItem26.Visible = true;

                lblItem27.Visible = true;
                Item27.Visible = true;
                qdt27.Visible = true;
                preco27.Visible = true;
                DescItem27.Visible = true;

                lblItem28.Visible = true;
                Item28.Visible = true;
                qdt28.Visible = true;
                preco28.Visible = true;
                DescItem28.Visible = true;

                lblItem29.Visible = true;
                Item29.Visible = true;
                qdt29.Visible = true;
                preco29.Visible = true;
                DescItem29.Visible = true;

                lblItem30.Visible = true;
                Item30.Visible = true;
                qdt30.Visible = true;
                preco30.Visible = true;
                DescItem30.Visible = true;

                lblItem31.Visible = true;
                Item31.Visible = true;
                qdt31.Visible = true;
                preco31.Visible = true;
                DescItem31.Visible = true;

                lblItem32.Visible = true;
                Item32.Visible = true;
                qdt32.Visible = true;
                preco32.Visible = true;
                DescItem32.Visible = true;

                lblItem33.Visible = true;
                Item33.Visible = true;
                qdt33.Visible = true;
                preco33.Visible = true;
                DescItem33.Visible = true;

                lblItem34.Visible = true;
                Item34.Visible = true;
                qdt34.Visible = true;
                preco34.Visible = true;
                DescItem34.Visible = true;

                lblItem35.Visible = true;
                Item35.Visible = true;
                qdt35.Visible = true;
                preco35.Visible = true;
                DescItem35.Visible = true;

                lblItem36.Visible = true;
                Item36.Visible = true;
                qdt36.Visible = true;
                preco36.Visible = true;
                DescItem36.Visible = true;

                lblItem37.Visible = true;
                Item37.Visible = true;
                qdt37.Visible = true;
                preco37.Visible = true;
                DescItem37.Visible = true;

                lblItem38.Visible = true;
                Item38.Visible = true;
                qdt38.Visible = true;
                preco38.Visible = true;
                DescItem38.Visible = true;
            }
            if (txtQTDItensNF.Text == "39")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;

                lblItem23.Visible = true;
                Item23.Visible = true;
                qdt23.Visible = true;
                preco23.Visible = true;
                DescItem23.Visible = true;

                lblItem24.Visible = true;
                Item24.Visible = true;
                qdt24.Visible = true;
                preco24.Visible = true;
                DescItem24.Visible = true;

                lblItem25.Visible = true;
                Item25.Visible = true;
                qdt25.Visible = true;
                DescItem25.Visible = true;

                lblItem26.Visible = true;
                Item26.Visible = true;
                qdt26.Visible = true;
                preco26.Visible = true;
                DescItem26.Visible = true;

                lblItem27.Visible = true;
                Item27.Visible = true;
                qdt27.Visible = true;
                preco27.Visible = true;
                DescItem27.Visible = true;

                lblItem28.Visible = true;
                Item28.Visible = true;
                qdt28.Visible = true;
                preco28.Visible = true;
                DescItem28.Visible = true;

                lblItem29.Visible = true;
                Item29.Visible = true;
                qdt29.Visible = true;
                preco29.Visible = true;
                DescItem29.Visible = true;

                lblItem30.Visible = true;
                Item30.Visible = true;
                qdt30.Visible = true;
                preco30.Visible = true;
                DescItem30.Visible = true;

                lblItem31.Visible = true;
                Item31.Visible = true;
                qdt31.Visible = true;
                preco31.Visible = true;
                DescItem31.Visible = true;

                lblItem32.Visible = true;
                Item32.Visible = true;
                qdt32.Visible = true;
                preco32.Visible = true;
                DescItem32.Visible = true;

                lblItem33.Visible = true;
                Item33.Visible = true;
                qdt33.Visible = true;
                preco33.Visible = true;
                DescItem33.Visible = true;

                lblItem34.Visible = true;
                Item34.Visible = true;
                qdt34.Visible = true;
                preco34.Visible = true;
                DescItem34.Visible = true;

                lblItem35.Visible = true;
                Item35.Visible = true;
                qdt35.Visible = true;
                preco35.Visible = true;
                DescItem35.Visible = true;

                lblItem36.Visible = true;
                Item36.Visible = true;
                qdt36.Visible = true;
                preco36.Visible = true;
                DescItem36.Visible = true;

                lblItem37.Visible = true;
                Item37.Visible = true;
                qdt37.Visible = true;
                preco37.Visible = true;
                DescItem37.Visible = true;

                lblItem38.Visible = true;
                Item38.Visible = true;
                qdt38.Visible = true;
                preco38.Visible = true;
                DescItem38.Visible = true;

                lblItem39.Visible = true;
                Item39.Visible = true;
                qdt39.Visible = true;
                preco39.Visible = true;
                DescItem39.Visible = true;
            }
            if (txtQTDItensNF.Text == "40")
            {
                lblItem2.Visible = true;
                Item2.Visible = true;
                qdt2.Visible = true;
                preco2.Visible = true;
                DescItem2.Visible = true;

                lblItem3.Visible = true;
                Item3.Visible = true;
                qdt3.Visible = true;
                preco3.Visible = true;
                DescItem3.Visible = true;

                lblItem4.Visible = true;
                Item4.Visible = true;
                qdt4.Visible = true;
                preco4.Visible = true;
                DescItem4.Visible = true;

                lblItem5.Visible = true;
                Item5.Visible = true;
                qdt5.Visible = true;
                preco5.Visible = true;
                DescItem5.Visible = true;

                lblItem6.Visible = true;
                Item6.Visible = true;
                qdt6.Visible = true;
                preco6.Visible = true;
                DescItem6.Visible = true;

                lblItem7.Visible = true;
                Item7.Visible = true;
                qdt7.Visible = true;
                preco7.Visible = true;
                DescItem7.Visible = true;

                lblItem8.Visible = true;
                Item8.Visible = true;
                qdt8.Visible = true;
                preco8.Visible = true;
                DescItem8.Visible = true;

                lblItem9.Visible = true;
                Item9.Visible = true;
                qdt9.Visible = true;
                preco9.Visible = true;
                DescItem9.Visible = true;

                lblItem10.Visible = true;
                Item10.Visible = true;
                qdt10.Visible = true;
                preco10.Visible = true;
                DescItem10.Visible = true;

                lblItem11.Visible = true;
                Item11.Visible = true;
                qdt11.Visible = true;
                preco11.Visible = true;
                DescItem11.Visible = true;

                lblItem12.Visible = true;
                Item12.Visible = true;
                qdt12.Visible = true;
                preco12.Visible = true;
                DescItem12.Visible = true;

                lblItem13.Visible = true;
                Item13.Visible = true;
                qdt13.Visible = true;
                preco13.Visible = true;
                DescItem13.Visible = true;

                lblItem14.Visible = true;
                Item14.Visible = true;
                qdt14.Visible = true;
                preco14.Visible = true;
                DescItem14.Visible = true;

                lblItem15.Visible = true;
                Item15.Visible = true;
                qdt15.Visible = true;
                preco15.Visible = true;
                DescItem15.Visible = true;

                lblItem16.Visible = true;
                Item16.Visible = true;
                qdt16.Visible = true;
                preco16.Visible = true;
                DescItem16.Visible = true;

                lblItem17.Visible = true;
                Item17.Visible = true;
                qdt17.Visible = true;
                preco17.Visible = true;
                DescItem17.Visible = true;

                lblItem18.Visible = true;
                Item18.Visible = true;
                qdt18.Visible = true;
                preco18.Visible = true;
                DescItem18.Visible = true;

                lblItem19.Visible = true;
                Item19.Visible = true;
                qdt19.Visible = true;
                preco19.Visible = true;
                DescItem19.Visible = true;

                lblItem20.Visible = true;
                Item20.Visible = true;
                qdt20.Visible = true;
                preco20.Visible = true;
                DescItem20.Visible = true;

                lblItem21.Visible = true;
                Item21.Visible = true;
                qdt21.Visible = true;
                preco21.Visible = true;
                DescItem21.Visible = true;

                lblItem22.Visible = true;
                Item22.Visible = true;
                qdt22.Visible = true;
                preco22.Visible = true;
                DescItem22.Visible = true;

                lblItem23.Visible = true;
                Item23.Visible = true;
                qdt23.Visible = true;
                preco23.Visible = true;
                DescItem23.Visible = true;

                lblItem24.Visible = true;
                Item24.Visible = true;
                qdt24.Visible = true;
                preco24.Visible = true;
                DescItem24.Visible = true;

                lblItem25.Visible = true;
                Item25.Visible = true;
                qdt25.Visible = true;
                DescItem25.Visible = true;

                lblItem26.Visible = true;
                Item26.Visible = true;
                qdt26.Visible = true;
                preco26.Visible = true;
                DescItem26.Visible = true;

                lblItem27.Visible = true;
                Item27.Visible = true;
                qdt27.Visible = true;
                preco27.Visible = true;
                DescItem27.Visible = true;

                lblItem28.Visible = true;
                Item28.Visible = true;
                qdt28.Visible = true;
                preco28.Visible = true;
                DescItem28.Visible = true;

                lblItem29.Visible = true;
                Item29.Visible = true;
                qdt29.Visible = true;
                preco29.Visible = true;
                DescItem29.Visible = true;

                lblItem30.Visible = true;
                Item30.Visible = true;
                qdt30.Visible = true;
                preco30.Visible = true;
                DescItem30.Visible = true;

                lblItem31.Visible = true;
                Item31.Visible = true;
                qdt31.Visible = true;
                preco31.Visible = true;
                DescItem31.Visible = true;

                lblItem32.Visible = true;
                Item32.Visible = true;
                qdt32.Visible = true;
                preco32.Visible = true;
                DescItem32.Visible = true;

                lblItem33.Visible = true;
                Item33.Visible = true;
                qdt33.Visible = true;
                preco33.Visible = true;
                DescItem33.Visible = true;

                lblItem34.Visible = true;
                Item34.Visible = true;
                qdt34.Visible = true;
                preco34.Visible = true;
                DescItem34.Visible = true;

                lblItem35.Visible = true;
                Item35.Visible = true;
                qdt35.Visible = true;
                preco35.Visible = true;
                DescItem35.Visible = true;

                lblItem36.Visible = true;
                Item36.Visible = true;
                qdt36.Visible = true;
                preco36.Visible = true;
                DescItem36.Visible = true;

                lblItem37.Visible = true;
                Item37.Visible = true;
                qdt37.Visible = true;
                preco37.Visible = true;
                DescItem37.Visible = true;

                lblItem38.Visible = true;
                Item38.Visible = true;
                qdt38.Visible = true;
                preco38.Visible = true;
                DescItem38.Visible = true;

                lblItem39.Visible = true;
                Item39.Visible = true;
                qdt39.Visible = true;
                preco39.Visible = true;
                DescItem39.Visible = true;

                lblItem40.Visible = true;
                Item40.Visible = true;
                qdt40.Visible = true;
                preco40.Visible = true;
                DescItem40.Visible = true;
            }
        }
        public string vbCr { get; private set; }
        private void FormNotaFiscal_Load(object sender, EventArgs e)
        {

        }
        private void DescItem33_TextChanged(object sender, EventArgs e)
        {

        }
        MySqlConnection CONEXAOBD;
        MySqlCommand cmd_item2, cmd_item3, cmd_item4, cmd_item5, cmd_item6, cmd_item7, cmd_item8, cmd_item9, cmd_item10, cmd_item11, cmd_item12, cmd_item13, cmd_item14, cmd_item15, cmd_item16, cmd_item17, cmd_item18, cmd_item19, cmd_item20, cmd_item21, cmd_item22, cmd_item23, cmd_item24, cmd_item25, cmd_item26, cmd_item27, cmd_item28, cmd_item29, cmd_item30, cmd_item31, cmd_item32, cmd_item33, cmd_item34, cmd_item35, cmd_item36, cmd_item37, cmd_item38, cmd_item39, cmd_item40;

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                ProgBar.Value = 55;
            }catch(Exception Error)
            {
                MessageBox.Show(Error.Message);
            }
        }

        private void btnPedido_Click(object sender, EventArgs e)
        {
            //Pega a tela de execução do Windows
            CSapROTWrapper sapROTWrapper = new CSapROTWrapper();
            //Pega a entrada ROT para o SAP Gui para conectar-se ao COM
            object SapGuilRot = sapROTWrapper.GetROTEntry("SAPGUI");
            //Pega a referência de Scripting Engine do SAP
            object engine = SapGuilRot.GetType().InvokeMember("GetScriptingEngine", System.Reflection.BindingFlags.InvokeMethod, null, SapGuilRot, null);
            //Pega a referência da janela de aplicativos em execução no SAP
            GuiApplication GuiApp = (GuiApplication)engine;
            //Pega a primeira conexão aberta do SAP
            GuiConnection connection = (GuiConnection)GuiApp.Connections.ElementAt(0);
            //Pega a primeira sessão aberta
            GuiSession Session = (GuiSession)connection.Children.ElementAt(0);
            //Pega a referência ao "FRAME" principal para enviar comandos de chaves virtuais o SAP
            GuiFrameWindow guiWindow = Session.ActiveWindow;
            //Maximisa Janela
            guiWindow.Maximize();
            //Abre Transação
            Session.SendCommand("/NME21N");
            //Inicia a Barra de Progresso em 25%
            ProgBar.Value = 25;
            //Tecla Enter
            guiWindow.SendVKey(0);
            //Cód. Fornecedor//
            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD")).Text = "1200005362";
            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD")).SetFocus();
            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD")).CaretPosition = 10;
            //Modifica o tipo de pedido
            ((GuiComboBox)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART")).Key = "NB";
            //Tecla Enter
            guiWindow.SendVKey(0);
            // Modifica o tipo de formato na data referênte ao Banco de Dados Ex.: 2020/06/27 para 27/06/2020.
            dateTimePicker1.Text = dtGridView.Rows[0].Cells[2].Value.ToString();
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd/MM/yyyy";
            //Dados Organizacionais
            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG")).Text = "1000";
            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP")).Text = "400";
            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS")).Text = "USGA";
            //Seleciona a aba texto e adiciona a nota fiscal e data
            ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3")).Select();
            ((GuiShell)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell")).Text = "REF. NOTA FISCAL Nº:" + txtChamarNotaFiscal.Text + " DE " + this.dateTimePicker1.Text + "." + vbCr + "";
            //Conta quantas linhas(itens) tem na nota fiscal referida
            int countg = dtGridView.RowCount;
            //Condição para criar o pedido com 1 item e por diante
            if (countg == 1)
            {
                try
                {
                    ProgBar.Value = 35;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Primeiro Item
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]")).Text = dtGridView.Rows[0].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]")).Text = dtGridView.Rows[0].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).CaretPosition = 4;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ProgBar.Value = 45;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[0].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    ProgBar.Value = 55;

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).CaretPosition = 1;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    ProgBar.Value = 75;

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[0].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);


                    ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();
                    GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");

                    string resultado = statusbar.Text.Substring(6, statusbar.Text.IndexOf('2'));
                    try
                    {
                        MySqlConnection CONEX = new MySqlConnection(@"server='" + txtHost.Text + "';database='" + txtDataBase.Text + "';Uid='" + txtUser.Text + "';Pwd='" + txtPass.Text + "';SslMode=none;");
                        CONEX.Open();
                        MySqlCommand cmd = new MySqlCommand("UPDATE `tb_nota` SET `PEDIDO`='" + resultado.Split('º')[1].Replace(" ", "") + "' WHERE DANFE='" + txtChamarNotaFiscal.Text + "'", CONEX);
                        cmd.ExecuteNonQuery();
                        CONEX.Close();
                    }
                    catch (MySqlException ErroMysql)
                    {
                        MessageBox.Show(ErroMysql.Message);
                    }
                }
                catch (Exception Errorr)
                {
                    MessageBox.Show(Errorr.Message);
                }
            }
            if (countg == 2)
            {
                try
                {
                    ProgBar.Value = 35;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Primeiro Item
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]")).Text = dtGridView.Rows[0].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]")).Text = dtGridView.Rows[0].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).CaretPosition = 4;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    ProgBar.Value = 45;

                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[0].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).CaretPosition = 1;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    ProgBar.Value = 60;

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[0].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Segundo Item

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dtGridView.Rows[1].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = dtGridView.Rows[1].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).CaretPosition = 28;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ProgBar.Value = 75;

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[1].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[1].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();
                    GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");

                    string resultado = statusbar.Text.Substring(6, statusbar.Text.IndexOf('2'));
                    try
                    {
                        MySqlConnection CONEX = new MySqlConnection(@"server='" + txtHost.Text + "';database='" + txtDataBase.Text + "';Uid='" + txtUser.Text + "';Pwd='" + txtPass.Text + "';SslMode=none;");
                        CONEX.Open();
                        MySqlCommand cmd = new MySqlCommand("UPDATE `tb_nota` SET `PEDIDO`='" + resultado.Split('º')[1].Replace(" ", "") + "' WHERE DANFE='" + txtChamarNotaFiscal.Text + "'", CONEX);
                        cmd.ExecuteNonQuery();
                        CONEX.Close();
                    }
                    catch (MySqlException ErroMysql)
                    {
                        MessageBox.Show(ErroMysql.Message);
                    }
                }
                catch (Exception Errorr)
                {
                    MessageBox.Show(Errorr.Message);
                }

                MessageBox.Show("Pedido Concluido!");
            }
            if (countg == 3)
            {
                try
                {
                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Primeiro Item
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]")).Text = dtGridView.Rows[0].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]")).Text = dtGridView.Rows[0].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).CaretPosition = 4;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[0].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).CaretPosition = 1;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[0].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Segundo Item

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dtGridView.Rows[1].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = dtGridView.Rows[1].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).CaretPosition = 28;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[1].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[1].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Terceiro Item

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,2]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,2]")).Text = dtGridView.Rows[2].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,2]")).Text = dtGridView.Rows[2].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).CaretPosition = 28;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[2].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[2].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 15;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();
                    GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");

                    string resultado = statusbar.Text.Substring(6, statusbar.Text.IndexOf('2'));
                    try
                    {
                        MySqlConnection CONEX = new MySqlConnection(@"server='" + txtHost.Text + "';database='" + txtDataBase.Text + "';Uid='" + txtUser.Text + "';Pwd='" + txtPass.Text + "';SslMode=none;");
                        CONEX.Open();
                        MySqlCommand cmd = new MySqlCommand("UPDATE `tb_nota` SET `PEDIDO`='" + resultado.Split('º')[1].Replace(" ", "") + "' WHERE DANFE='" + txtChamarNotaFiscal.Text + "'", CONEX);
                        cmd.ExecuteNonQuery();
                        CONEX.Close();
                    }
                    catch (MySqlException ErroMysql)
                    {
                        MessageBox.Show(ErroMysql.Message);
                    }
                }
                catch (Exception Errorr)
                {
                    MessageBox.Show(Errorr.Message);
                }

                MessageBox.Show("Pedido Concluido!");
            }
            if (countg == 5)
            {
                try
                {
                    ProgBar.Value = 35;
                    //Tecla Enter
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]")).Text = dtGridView.Rows[0].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]")).Text = dtGridView.Rows[0].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).CaretPosition = 28;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[0].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[0].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    ((GuiButton)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON")).Press();
                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")).VerticalScrollbar.Position = 1;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dtGridView.Rows[1].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = dtGridView.Rows[1].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).CaretPosition = 28;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[1].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[1].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    //Terceiro
                    ((GuiButton)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON")).Press();
                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")).VerticalScrollbar.Position = 2;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dtGridView.Rows[2].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = dtGridView.Rows[2].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).CaretPosition = 28;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[2].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[2].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    //Quarto
                    ((GuiButton)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON")).Press();
                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")).VerticalScrollbar.Position = 3;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dtGridView.Rows[3].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = dtGridView.Rows[3].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).CaretPosition = 28;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[3].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[3].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    ///Quinto 




























                    ///Pressiona o Botão Gravar
                    ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();
                    //Pega a Barra de Status do SAP
                    GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");
                    //Me retorna apenas o número do pedido no tratamento da importação no Banco de Dados ele retira o º e os espaços.
                    string resultado = statusbar.Text.Substring(6, statusbar.Text.IndexOf('2'));
                    try
                    {
                        MySqlConnection CONEX = new MySqlConnection(@"server='" + txtHost.Text + "';database='" + txtDataBase.Text + "';Uid='" + txtUser.Text + "';Pwd='" + txtPass.Text + "';SslMode=none;");
                        CONEX.Open();
                        MySqlCommand cmd = new MySqlCommand("UPDATE `tb_nota` SET `PEDIDO`='" + resultado.Split('º')[1].Replace(" ", "") + "' WHERE DANFE='" + txtChamarNotaFiscal.Text + "'", CONEX);
                        cmd.ExecuteNonQuery();
                        CONEX.Close();
                    }
                    catch (MySqlException ErroMysql)
                    {
                        MessageBox.Show(ErroMysql.Message);
                    }
                }
                catch (Exception Erro)
                {
                    MessageBox.Show(Erro.Message);
                }
            }
            if(countg == 6)
            {
                try
                {
                    ProgBar.Value = 35;
                    //Tecla Enter
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,2]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,3]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,4]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]")).Text = "694850";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = "694850";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,2]")).Text = "694850";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,3]")).Text = "694850";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,4]")).Text = "694850";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]")).Text = "11";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = "11";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,2]")).Text = "11";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,3]")).Text = "11";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,4]")).Text = "11";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,3]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,4]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,4]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,4]")).CaretPosition = 0;

                    //Tecla Enter
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = "SG01020201";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter
                    guiWindow.SendVKey(0);


                    ProgBar.Value = 45;

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = "32";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = "SG01020201";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = "152";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ProgBar.Value = 50;

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = "SG01020201";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = "52";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ProgBar.Value = 75;

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = "SG01020201";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = "526";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ProgBar.Value = 85;

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = "SG01020201";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter
                    guiWindow.SendVKey(0);

                    ProgBar.Value = 92;

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = "856";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter
                    guiWindow.SendVKey(0);

                    ///Pressiona o Botão Gravar
                    //((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();
                    //Pega a Barra de Status do SAP
                   // GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");
                    //Me retorna apenas o número do pedido no tratamento da importação no Banco de Dados ele retira o º e os espaços.
                    //string resultado = statusbar.Text.Substring(6, statusbar.Text.IndexOf('2'));
                    //try
                    //{
                    //   MySqlConnection CONEX = new MySqlConnection(@"server='" + txtHost.Text + "';database='" + txtDataBase.Text + "';Uid='" + txtUser.Text + "';Pwd='" + txtPass.Text + "';SslMode=none;");
                    //    CONEX.Open();
                    //    MySqlCommand cmd = new MySqlCommand("UPDATE `tb_nota` SET `PEDIDO`='" + resultado.Split('º')[1].Replace(" ", "") + "' WHERE DANFE='" + txtChamarNotaFiscal.Text + "'", CONEX);
                    //    cmd.ExecuteNonQuery();
                    //    CONEX.Close();
                   // }
                    //catch (MySqlException ErroMysql)
                    //{
                     //   MessageBox.Show(ErroMysql.Message);
                    //}
                }
                catch (Exception Erro)
                {
                    MessageBox.Show(Erro.Message);
                }

            }
            if (countg == 11)
            {
                try
                {
                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Primeiro Item
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]")).Text = dtGridView.Rows[0].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]")).Text = dtGridView.Rows[0].Cells[4].Value.ToString() ;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).CaretPosition = 4;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[0].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).CaretPosition = 1;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[0].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Segundo Item

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dtGridView.Rows[1].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = dtGridView.Rows[1].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).CaretPosition = 28;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[1].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[1].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Terceiro Item

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,2]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,2]")).Text = dtGridView.Rows[2].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,2]")).Text = dtGridView.Rows[2].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).CaretPosition = 28;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[2].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[2].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 15;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Rolar Barra Vertical
                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")).VerticalScrollbar.Position = 2;

                    //Segundo Item

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dtGridView.Rows[3].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = dtGridView.Rows[3].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).CaretPosition = 28;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[3].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[3].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Rolar Barra Vertical
                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")).VerticalScrollbar.Position = 10;



                    ((GuiButton)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON")).Press();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,2]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,3]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,4]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,5]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,6]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,2]")).Text = dtGridView.Rows[4].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,3]")).Text = dtGridView.Rows[5].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,4]")).Text = dtGridView.Rows[6].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,5]")).Text = dtGridView.Rows[7].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,6]")).Text = dtGridView.Rows[8].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,2]")).Text = dtGridView.Rows[4].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,3]")).Text = dtGridView.Rows[5].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,4]")).Text = dtGridView.Rows[6].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,5]")).Text = dtGridView.Rows[7].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,6]")).Text = dtGridView.Rows[8].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,3]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,4]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,5]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,6]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,6]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,6]")).CaretPosition = 28;
                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[4].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[4].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[5].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[5].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[6].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[6].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[7].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[7].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[8].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[8].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ((GuiButton)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON")).Press();
                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")).VerticalScrollbar.Position = 18;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dtGridView.Rows[9].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = dtGridView.Rows[9].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).CaretPosition = 28;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[9].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    guiWindow.SendVKey(0);
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[9].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);

                    ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();
                    GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");

                    string resultado = statusbar.Text.Substring(6, statusbar.Text.IndexOf('2'));
                    try
                    {
                        MySqlConnection CONEX = new MySqlConnection(@"server='" + txtHost.Text + "';database='" + txtDataBase.Text + "';Uid='" + txtUser.Text + "';Pwd='" + txtPass.Text + "';SslMode=none;");
                        CONEX.Open();
                        MySqlCommand cmd = new MySqlCommand("UPDATE `tb_nota` SET `PEDIDO`='" + resultado.Split('º')[1].Replace(" ", "") + "' WHERE DANFE='" + txtChamarNotaFiscal.Text + "'", CONEX);
                        cmd.ExecuteNonQuery();
                        CONEX.Close();
                    }
                    catch(MySqlException ErroMysql)
                    {
                        MessageBox.Show(ErroMysql.Message);
                    }
                }
                catch (Exception Errorr)
                {
                    MessageBox.Show(Errorr.Message);
                }

                MessageBox.Show("Pedido Concluido!");
            }
            if (countg == 12)
            {
                try
                {
                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Primeiro Item
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]")).Text = dtGridView.Rows[0].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]")).Text = dtGridView.Rows[0].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).CaretPosition = 4;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[0].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).CaretPosition = 1;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[0].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Segundo Item

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dtGridView.Rows[1].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = dtGridView.Rows[1].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).CaretPosition = 28;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[1].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[1].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Terceiro Item

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,2]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,2]")).Text = dtGridView.Rows[2].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,2]")).Text = dtGridView.Rows[2].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).CaretPosition = 28;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[2].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[2].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 15;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Rolar Barra Vertical
                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")).VerticalScrollbar.Position = 2;

                    //Segundo Item

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dtGridView.Rows[3].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = dtGridView.Rows[3].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).CaretPosition = 28;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[3].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[3].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Rolar Barra Vertical
                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")).VerticalScrollbar.Position = 10;



                    ((GuiButton)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON")).Press();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,2]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,3]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,4]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,5]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,6]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,2]")).Text = dtGridView.Rows[4].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,3]")).Text = dtGridView.Rows[5].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,4]")).Text = dtGridView.Rows[6].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,5]")).Text = dtGridView.Rows[7].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,6]")).Text = dtGridView.Rows[8].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,2]")).Text = dtGridView.Rows[4].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,3]")).Text = dtGridView.Rows[5].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,4]")).Text = dtGridView.Rows[6].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,5]")).Text = dtGridView.Rows[7].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,6]")).Text = dtGridView.Rows[8].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,3]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,4]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,5]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,6]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,6]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,6]")).CaretPosition = 28;
                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[4].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[4].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    //guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[5].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[5].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    //guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[6].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[6].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[7].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[7].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    //guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[8].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[8].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ((GuiButton)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON")).Press();
                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")).VerticalScrollbar.Position = 30;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dtGridView.Rows[9].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = dtGridView.Rows[9].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).CaretPosition = 28;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[9].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    guiWindow.SendVKey(0);
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[9].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);

                    ((GuiButton)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON")).Press();
                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")).VerticalScrollbar.Position = 60;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dtGridView.Rows[10].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = dtGridView.Rows[10].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).CaretPosition = 28;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[10].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    guiWindow.SendVKey(0);
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[10].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);


                    ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();
                    GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");

                    string resultado = statusbar.Text.Substring(6, statusbar.Text.IndexOf('2'));
                    try
                    {
                        MySqlConnection CONEX = new MySqlConnection(@"server='" + txtHost.Text + "';database='" + txtDataBase.Text + "';Uid='" + txtUser.Text + "';Pwd='" + txtPass.Text + "';SslMode=none;");
                        CONEX.Open();
                        MySqlCommand cmd = new MySqlCommand("UPDATE `tb_nota` SET `PEDIDO`='" + resultado.Split('º')[1].Replace(" ", "") + "' WHERE DANFE='" + txtChamarNotaFiscal.Text + "'", CONEX);
                        cmd.ExecuteNonQuery();
                        CONEX.Close();
                    }
                    catch (MySqlException ErroMysql)
                    {
                        MessageBox.Show(ErroMysql.Message);
                    }
                }
                catch (Exception Errorr)
                {
                    MessageBox.Show(Errorr.Message);
                }

                MessageBox.Show("Pedido Concluido!");
            }
            if (countg == 13)
            {
                try
                {
                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Primeiro Item
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]")).Text = dtGridView.Rows[0].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]")).Text = dtGridView.Rows[0].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).CaretPosition = 4;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[0].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).CaretPosition = 1;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[0].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Segundo Item

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dtGridView.Rows[1].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = dtGridView.Rows[1].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).CaretPosition = 28;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[1].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[1].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Terceiro Item

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,2]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,2]")).Text = dtGridView.Rows[2].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,2]")).Text = dtGridView.Rows[2].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).CaretPosition = 28;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[2].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[2].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 15;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Rolar Barra Vertical
                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")).VerticalScrollbar.Position = 2;

                    //Segundo Item

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dtGridView.Rows[3].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = dtGridView.Rows[3].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).CaretPosition = 28;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[3].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[3].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;

                    //Tecla Enter//
                    guiWindow.SendVKey(0);

                    //Rolar Barra Vertical
                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")).VerticalScrollbar.Position = 10;



                    ((GuiButton)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON")).Press();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,2]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,3]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,4]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,5]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,6]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,2]")).Text = dtGridView.Rows[4].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,3]")).Text = dtGridView.Rows[5].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,4]")).Text = dtGridView.Rows[6].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,5]")).Text = dtGridView.Rows[7].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,6]")).Text = dtGridView.Rows[8].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,2]")).Text = dtGridView.Rows[4].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,3]")).Text = dtGridView.Rows[5].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,4]")).Text = dtGridView.Rows[6].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,5]")).Text = dtGridView.Rows[7].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,6]")).Text = dtGridView.Rows[8].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,3]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,4]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,5]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,6]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,6]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,6]")).CaretPosition = 28;
                    //Tecla Enter//
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[4].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[4].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    //guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[5].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[5].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[6].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[6].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[7].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[7].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    //guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[8].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[8].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);

                    ((GuiButton)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON")).Press();
                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")).VerticalScrollbar.Position = 24;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dtGridView.Rows[9].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = dtGridView.Rows[9].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).CaretPosition = 28;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[9].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "d0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    guiWindow.SendVKey(0);
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[9].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);

                    ((GuiButton)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON")).Press();
                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")).VerticalScrollbar.Position = 25;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dtGridView.Rows[10].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = dtGridView.Rows[10].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).CaretPosition = 28;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[10].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    guiWindow.SendVKey(0);
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[10].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,2]")).Text = "K";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,2]")).Text = dtGridView.Rows[11].Cells[15].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,2]")).Text = dtGridView.Rows[11].Cells[4].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,2]")).CaretPosition = 28;
                    guiWindow.SendVKey(0);
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView.Rows[11].Cells[7].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                    guiWindow.SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = "D0";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView.Rows[11].Cells[5].Value.ToString();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtRV61A-KOEIN[4,0]")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtRV61A-KOEIN[4,0]")).CaretPosition = 3;
                    guiWindow.SendVKey(0);

                    ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();
                    GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");

                    string resultado = statusbar.Text.Substring(6, statusbar.Text.IndexOf('2'));
                    try
                    {
                        MySqlConnection CONEX = new MySqlConnection(@"server='" + txtHost.Text + "';database='" + txtDataBase.Text + "';Uid='" + txtUser.Text + "';Pwd='" + txtPass.Text + "';SslMode=none;");
                       CONEX.Open();
                        MySqlCommand cmd = new MySqlCommand("UPDATE `tb_nota` SET `PEDIDO`='" + resultado.Split('º')[1].Replace(" ", "") + "' WHERE DANFE='" + txtChamarNotaFiscal.Text + "'", CONEX);
                        cmd.ExecuteNonQuery();
                        CONEX.Close();
                    }
                    catch (MySqlException ErroMysql)
                    {
                        MessageBox.Show(ErroMysql.Message);
                    }
                }
                catch (Exception Errorr)
                {
                    MessageBox.Show(Errorr.Message);
                }

                MessageBox.Show("Pedido Concluido!");
            }
            //Finaliza com 100% a Barra de Progresso
            ProgBar.Value = 100;
            //Exibe uma menssagem de conclusão
            MessageBox.Show("Pedido Concluido!");
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                CONEXAOBD = new MySqlConnection(@"server='" + txtHost.Text + "';database='" + txtDataBase.Text + "';Uid='" + txtUser.Text + "';Pwd='" + txtPass.Text + "';SslMode=none;");
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT N.*,P.DESC_PRODUTO,P.CD_SAP FROM tb_nota N JOIN tb_produtos P ON N.COD_PRODUTO=P.CD_PRODUTO WHERE N.DANFE='"+txtChamarNotaFiscal.Text+"' ORDER BY N.ID_TB ASC", CONEXAOBD);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dtGridView.DataSource = SS;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnCad_Click(object sender, EventArgs e)
        {
            var conexaoform = new FormCadastroItem(txtHost.Text);
            conexaoform.Show();
            //FormCadastroItem FRM = new FormCadastroItem();
            //FRM.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void tabPage3_Click(object sender, EventArgs e)
        {

        }
        private void txtChave_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void Ocultar()
        {
            lblItem2.Visible = false;
            Item2.Visible = false;
            qdt2.Visible = false;
            preco2.Visible = false;
            DescItem2.Visible = false;

            lblItem3.Visible = false;
            Item3.Visible = false;
            qdt3.Visible = false;
            preco3.Visible = false;
            DescItem3.Visible = false;

            lblItem4.Visible = false;
            Item4.Visible = false;
            qdt4.Visible = false;
            preco4.Visible = false;
            DescItem4.Visible = false;

            lblItem5.Visible = false;
            Item5.Visible = false;
            qdt5.Visible = false;
            preco5.Visible = false;
            DescItem5.Visible = false;

            lblItem6.Visible = false;
            Item6.Visible = false;
            qdt6.Visible = false;
            preco6.Visible = false;
            DescItem6.Visible = false;

            lblItem7.Visible = false;
            Item7.Visible = false;
            qdt7.Visible = false;
            preco7.Visible = false;
            DescItem7.Visible = false;

            lblItem8.Visible = false;
            Item8.Visible = false;
            qdt8.Visible = false;
            preco8.Visible = false;
            DescItem8.Visible = false;

            lblItem9.Visible = false;
            Item9.Visible = false;
            qdt9.Visible = false;
            preco9.Visible = false;
            DescItem9.Visible = false;

            lblItem10.Visible = false;
            Item10.Visible = false;
            qdt10.Visible = false;
            preco10.Visible = false;
            DescItem10.Visible = false;

            lblItem11.Visible = false;
            Item11.Visible = false;
            qdt11.Visible = false;
            preco11.Visible = false;
            DescItem11.Visible = false;

            lblItem12.Visible = false;
            Item12.Visible = false;
            qdt12.Visible = false;
            preco12.Visible = false;
            DescItem12.Visible = false;

            lblItem13.Visible = false;
            Item13.Visible = false;
            qdt13.Visible = false;
            preco13.Visible = false;
            DescItem13.Visible = false;

            lblItem14.Visible = false;
            Item14.Visible = false;
            qdt14.Visible = false;
            preco14.Visible = false;
            DescItem14.Visible = false;

            lblItem15.Visible = false;
            Item15.Visible = false;
            qdt15.Visible = false;
            preco15.Visible = false;
            DescItem15.Visible = false;

            lblItem16.Visible = false;
            Item16.Visible = false;
            qdt16.Visible = false;
            preco16.Visible = false;
            DescItem16.Visible = false;

            lblItem17.Visible = false;
            Item17.Visible = false;
            qdt17.Visible = false;
            preco17.Visible = false;
            DescItem17.Visible = false;

            lblItem18.Visible = false;
            Item18.Visible = false;
            qdt18.Visible = false;
            preco18.Visible = false;
            DescItem18.Visible = false;

            lblItem19.Visible = false;
            Item19.Visible = false;
            qdt19.Visible = false;
            preco19.Visible = false;
            DescItem19.Visible = false;

            lblItem20.Visible = false;
            Item20.Visible = false;
            qdt20.Visible = false;
            preco20.Visible = false;
            DescItem20.Visible = false;

            lblItem21.Visible = false;
            Item21.Visible = false;
            qdt21.Visible = false;
            preco21.Visible = false;
            DescItem21.Visible = false;

            lblItem22.Visible = false;
            Item22.Visible = false;
            qdt22.Visible = false;
            preco22.Visible = false;
            DescItem22.Visible = false;

            lblItem23.Visible = false;
            Item23.Visible = false;
            qdt23.Visible = false;
            preco23.Visible = false;
            DescItem23.Visible = false;

            lblItem24.Visible = false;
            Item24.Visible = false;
            qdt24.Visible = false;
            preco24.Visible = false;
            DescItem24.Visible = false;

            lblItem25.Visible = false;
            Item25.Visible = false;
            qdt25.Visible = false;
            DescItem25.Visible = false;

            lblItem26.Visible = false;
            Item26.Visible = false;
            qdt26.Visible = false;
            preco26.Visible = false;
            DescItem26.Visible = false;

            lblItem27.Visible = false;
            Item27.Visible = false;
            qdt27.Visible = false;
            preco27.Visible = false;
            DescItem27.Visible = false;

            lblItem28.Visible = false;
            Item28.Visible = false;
            qdt28.Visible = false;
            preco28.Visible = false;
            DescItem28.Visible = false;

            lblItem29.Visible = false;
            Item29.Visible = false;
            qdt29.Visible = false;
            preco29.Visible = false;
            DescItem29.Visible = false;

            lblItem30.Visible = false;
            Item30.Visible = false;
            qdt30.Visible = false;
            preco30.Visible = false;
            DescItem30.Visible = false;

            lblItem31.Visible = false;
            Item31.Visible = false;
            qdt31.Visible = false;
            preco31.Visible = false;
            DescItem31.Visible = false;

            lblItem32.Visible = false;
            Item32.Visible = false;
            qdt32.Visible = false;
            preco32.Visible = false;
            DescItem32.Visible = false;

            lblItem33.Visible = false;
            Item33.Visible = false;
            qdt33.Visible = false;
            preco33.Visible = false;
            DescItem33.Visible = false;

            lblItem34.Visible = false;
            Item34.Visible = false;
            qdt34.Visible = false;
            preco34.Visible = false;
            DescItem34.Visible = false;

            lblItem35.Visible = false;
            Item35.Visible = false;
            qdt35.Visible = false;
            preco35.Visible = false;
            DescItem35.Visible = false;

            lblItem36.Visible = false;
            Item36.Visible = false;
            qdt36.Visible = false;
            preco36.Visible = false;
            DescItem36.Visible = false;

            lblItem37.Visible = false;
            Item37.Visible = false;
            qdt37.Visible = false;
            preco37.Visible = false;
            DescItem37.Visible = false;

            lblItem38.Visible = false;
            Item38.Visible = false;
            qdt38.Visible = false;
            preco38.Visible = false;
            DescItem38.Visible = false;

            lblItem39.Visible = false;
            Item39.Visible = false;
            qdt39.Visible = false;
            preco39.Visible = false;
            DescItem39.Visible = false;

            lblItem40.Visible = false;
            Item40.Visible = false;
            qdt40.Visible = false;
            preco40.Visible = false;
            DescItem40.Visible = false;
        }
        private void Limpar()
        {
            txtChaveUltimo.Text = "";
            txtChave.Text = "";
            txtProtocolo.Text = "";
            txtNfe.Text = "";
            Item1.Text = "";
            qdt1.Text = "";
            preco1.Text = "";
            DescItem1.Text = "";

            Item2.Text = "";
            qdt2.Text = "";
            preco2.Text = "";
            DescItem2.Text = "";

            Item3.Text = "";
            qdt3.Text = "";
            preco3.Text = "";
            DescItem3.Text = "";

            Item4.Text = "";
            qdt4.Text = "";
            preco4.Text = "";
            DescItem4.Text = "";

            Item5.Text = "";
            qdt5.Text = "";
            preco5.Text = "";
            DescItem5.Text = "";

            Item6.Text = "";
            qdt6.Text = "";
            preco6.Text = "";
            DescItem6.Text = "";

            Item7.Text = "";
            qdt7.Text = "";
            preco7.Text = "";
            DescItem7.Text = "";

            Item8.Text = "";
            qdt8.Text = "";
            preco8.Text = "";
            DescItem8.Text = "";

            Item9.Text = "";
            qdt9.Text = "";
            preco9.Text = "";
            DescItem9.Text = "";

            Item10.Text = "";
            qdt10.Text = "";
            preco10.Text = "";
            DescItem10.Text = "";

            Item11.Text = "";
            qdt11.Text = "";
            preco11.Text = "";
            DescItem11.Text = "";

            Item12.Text = "";
            qdt12.Text = "";
            preco12.Text = "";
            DescItem12.Text = "";

            Item13.Text = "";
            qdt13.Text = "";
            preco13.Text = "";
            DescItem13.Text = "";

            Item14.Text = "";
            qdt14.Text = "";
            preco14.Text = "";
            DescItem14.Text = "";

            Item15.Text = "";
            qdt15.Text = "";
            preco15.Text = "";
            DescItem15.Text = "";

            Item16.Text = "";
            qdt16.Text = "";
            preco16.Text = "";
            DescItem16.Text = "";

            Item17.Text = "";
            qdt17.Text = "";
            preco17.Text = "";
            DescItem17.Text = "";

            Item18.Text = "";
            qdt18.Text = "";
            preco18.Text = "";
            DescItem18.Text = "";

            Item19.Text = "";
            qdt19.Text = "";
            preco19.Text = "";
            DescItem19.Text = "";

            Item20.Text = "";
            qdt20.Text = "";
            preco20.Text = "";
            DescItem20.Text = "";

            Item21.Text = "";
            qdt21.Text = "";
            preco21.Text = "";
            DescItem21.Text = "";

            Item22.Text = "";
            qdt22.Text = "";
            preco22.Text = "";
            DescItem22.Text = "";

            Item23.Text = "";
            qdt23.Text = "";
            preco23.Text = "";
            DescItem23.Text = "";

            Item24.Text = "";
            qdt24.Text = "";
            preco24.Text = "";
            DescItem24.Text = "";

            Item25.Text = "";
            qdt25.Text = "";
            DescItem25.Text = "";

            Item26.Text = "";
            qdt26.Text = "";
            preco26.Text = "";
            DescItem26.Text = "";

            Item27.Text = "";
            qdt27.Text = "";
            preco27.Text = "";
            DescItem27.Text = "";

            Item28.Text = "";
            qdt28.Text = "";
            preco28.Text = "";
            DescItem28.Text = "";

            Item29.Text = "";
            qdt29.Text = "";
            preco29.Text = "";
            DescItem29.Text = "";

            Item30.Text = "";
            qdt30.Text = "";
            preco30.Text = "";
            DescItem30.Text = "";

            Item31.Text = "";
            qdt31.Text = "";
            preco31.Text = "";
            DescItem31.Text = "";

            Item32.Text = "";
            qdt32.Text = "";
            preco32.Text = "";
            DescItem32.Text = "";

            Item33.Text = "";
            qdt33.Text = "";
            preco33.Text = "";
            DescItem33.Text = "";

            Item34.Text = "";
            qdt34.Text = "";
            preco34.Text = "";
            DescItem34.Text = "";

            Item35.Text = "";
            qdt35.Text = "";
            preco35.Text = "";
            DescItem35.Text = "";

            Item36.Text = "";
            qdt36.Text = "";
            preco36.Text = "";
            DescItem36.Text = "";

            Item37.Text = "";
            qdt37.Text = "";
            preco37.Text = "";
            DescItem37.Text = "";

            Item38.Text = "";
            qdt38.Text = "";
            preco38.Text = "";
            DescItem38.Text = "";

            Item39.Text = "";
            qdt39.Text = "";
            preco39.Text = "";
            DescItem39.Text = "";

            Item40.Text = "";
            qdt40.Text = "";
            preco40.Text = "";
            DescItem40.Text = "";
        }
        private void btnGravar_Click(object sender, EventArgs e)
        {
            try
            {
                try
                {
                    CONEXAOBD = new MySqlConnection(@"server='"+txtHost.Text+"';database='"+txtDataBase.Text+"';Uid='"+txtUser.Text+"';Pwd='"+txtPass.Text+"';SslMode=none;");
                    CONEXAOBD.Open();

                    if (string.IsNullOrEmpty(Item1.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem1 = Convert.ToDouble(qdt1.Text) * Convert.ToDouble(preco1.Text);
                        cmd_item1 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item1.Text + "', '" + qdt1.Text + "', '" + preco1.Text + "', '" + vlTotalItem1 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item1.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item2.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem2 = Convert.ToDouble(qdt2.Text) * Convert.ToDouble(preco2.Text);
                        cmd_item2 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item2.Text + "', '" + qdt2.Text + "', '" + preco2.Text + "', '" + vlTotalItem2 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item2.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item3.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem3 = Convert.ToDouble(qdt3.Text) * Convert.ToDouble(preco3.Text);
                        cmd_item3 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item3.Text + "', '" + qdt3.Text + "', '" + preco3.Text + "', '" + vlTotalItem3 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item3.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item4.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem4 = Convert.ToDouble(qdt4.Text) * Convert.ToDouble(preco4.Text);
                        
                        cmd_item4 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item4.Text + "', '" + qdt4.Text + "', '" + preco4.Text + "', '" + vlTotalItem4 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item4.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item5.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem5 = Convert.ToDouble(qdt5.Text) * Convert.ToDouble(preco5.Text);
                        
                        cmd_item5 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item5.Text + "', '" + qdt5.Text + "', '" + preco5.Text + "', '" + vlTotalItem5 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item5.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item6.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem6 = Convert.ToDouble(qdt6.Text) * Convert.ToDouble(preco6.Text);
                        

                        cmd_item6 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item6.Text + "', '" + qdt6.Text + "', '" + preco6.Text + "', '" + vlTotalItem6 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item6.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item7.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem7 = Convert.ToDouble(qdt7.Text) * Convert.ToDouble(preco7.Text);
                        
                        cmd_item7 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item7.Text + "', '" + qdt7.Text + "', '" + preco7.Text + "', '" + vlTotalItem7 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item7.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item8.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem8 = Convert.ToDouble(qdt8.Text) * Convert.ToDouble(preco8.Text);
                        
                        cmd_item8 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item8.Text + "', '" + qdt8.Text + "', '" + preco8.Text + "', '" + vlTotalItem8 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item8.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item9.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem9 = Convert.ToDouble(qdt9.Text) * Convert.ToDouble(preco9.Text);
                        
                        cmd_item9 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item9.Text + "', '" + qdt9.Text + "', '" + preco9.Text + "', '" + vlTotalItem9 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item9.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item10.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem10 = Convert.ToDouble(qdt10.Text) * Convert.ToDouble(preco10.Text);
                        
                        cmd_item10 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item10.Text + "', '" + qdt10.Text + "', '" + preco10.Text + "', '" + vlTotalItem10 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item10.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item11.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem11 = Convert.ToDouble(qdt11.Text) * Convert.ToDouble(preco11.Text);
                        
                        cmd_item11 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item11.Text + "', '" + qdt11.Text + "', '" + preco11.Text + "', '" + vlTotalItem11 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item11.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item12.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem12 = Convert.ToDouble(qdt12.Text) * Convert.ToDouble(preco12.Text);
                        
                        cmd_item12 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item12.Text + "', '" + qdt12.Text + "', '" + preco12.Text + "', '" + vlTotalItem12 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item12.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item13.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem13 = Convert.ToDouble(qdt13.Text) * Convert.ToDouble(preco13.Text);
                        
                        cmd_item13 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item13.Text + "', '" + qdt13.Text + "', '" + preco13.Text + "', '" + vlTotalItem13 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item13.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item14.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem14 = Convert.ToDouble(qdt14.Text) * Convert.ToDouble(preco14.Text);
                        
                        cmd_item14 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item14.Text + "', '" + qdt14.Text + "', '" + preco14.Text + "', '" + vlTotalItem14 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item14.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item15.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem15 = Convert.ToDouble(qdt15.Text) * Convert.ToDouble(preco15.Text);
                        
                        cmd_item15 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item15.Text + "', '" + qdt15.Text + "', '" + preco15.Text + "', '" + vlTotalItem15 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item15.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item16.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem16 = Convert.ToDouble(qdt16.Text) * Convert.ToDouble(preco16.Text);
                        
                        cmd_item16 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item16.Text + "', '" + qdt16.Text + "', '" + preco16.Text + "', '" + vlTotalItem16 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item16.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item17.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem17 = Convert.ToDouble(qdt17.Text) * Convert.ToDouble(preco17.Text);
                        
                        cmd_item17 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item17.Text + "', '" + qdt17.Text + "', '" + preco17.Text + "', '" + vlTotalItem17 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item17.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item18.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem18 = Convert.ToDouble(qdt18.Text) * Convert.ToDouble(preco18.Text);
                        
                        cmd_item18 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item18.Text + "', '" + qdt18.Text + "', '" + preco18.Text + "', '" + vlTotalItem18 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item18.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item19.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem19 = Convert.ToDouble(qdt19.Text) * Convert.ToDouble(preco19.Text);
                        
                        cmd_item19 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item19.Text + "', '" + qdt19.Text + "', '" + preco19.Text + "', '" + vlTotalItem19 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item19.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item20.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem20 = Convert.ToDouble(qdt20.Text) * Convert.ToDouble(preco20.Text);

                        cmd_item20 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item20.Text + "', '" + qdt20.Text + "', '" + preco20.Text + "', '" + vlTotalItem20 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item20.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item21.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem21 = Convert.ToDouble(qdt21.Text) * Convert.ToDouble(preco21.Text);
                        
                        cmd_item21 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item21.Text + "', '" + qdt21.Text + "', '" + preco21.Text + "', '" + vlTotalItem21 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item21.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item22.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem22 = Convert.ToDouble(qdt22.Text) * Convert.ToDouble(preco22.Text);
                        

                        cmd_item22 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item22.Text + "', '" + qdt22.Text + "', '" + preco22.Text + "', '" + vlTotalItem22 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item22.ExecuteNonQuery();

                    }
                    if (string.IsNullOrEmpty(Item23.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem23 = Convert.ToDouble(qdt23.Text) * Convert.ToDouble(preco23.Text);
                        
                        cmd_item23 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item23.Text + "', '" + qdt23.Text + "', '" + preco23.Text + "', '" + vlTotalItem23 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item23.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item24.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem24 = Convert.ToDouble(qdt24.Text) * Convert.ToDouble(preco24.Text);
                        
                        cmd_item24 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item24.Text + "', '" + qdt24.Text + "', '" + preco24.Text + "', '" + vlTotalItem24 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item24.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item25.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem25 = Convert.ToDouble(qdt25.Text) * Convert.ToDouble(preco25.Text);
                        

                        cmd_item25 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item25.Text + "', '" + qdt25.Text + "', '" + preco25.Text + "', '" + vlTotalItem25 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item25.ExecuteNonQuery();

                    }
                    if (string.IsNullOrEmpty(Item26.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem26 = Convert.ToDouble(qdt26.Text) * Convert.ToDouble(preco26.Text);
                        
                        cmd_item26 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item26.Text + "', '" + qdt26.Text + "', '" + preco26.Text + "', '" + vlTotalItem26 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item26.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item27.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem27 = Convert.ToDouble(qdt27.Text) * Convert.ToDouble(preco27.Text);
                        
                        cmd_item27 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item27.Text + "', '" + qdt27.Text + "', '" + preco27.Text + "', '" + vlTotalItem27 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item27.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item28.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem28 = Convert.ToDouble(qdt28.Text) * Convert.ToDouble(preco28.Text);
                        
                        cmd_item28 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item28.Text + "', '" + qdt28.Text + "', '" + preco28.Text + "', '" + vlTotalItem28 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item28.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item29.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem29 = Convert.ToDouble(qdt29.Text) * Convert.ToDouble(preco29.Text);
                        
                        cmd_item29 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item29.Text + "', '" + qdt29.Text + "', '" + preco29.Text + "', '" + vlTotalItem29 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item29.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item30.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem30 = Convert.ToDouble(qdt30.Text) * Convert.ToDouble(preco30.Text);
                        
                        cmd_item30 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item30.Text + "', '" + qdt30.Text + "', '" + preco30.Text + "', '" + vlTotalItem30 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item30.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item31.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem31 = Convert.ToDouble(qdt31.Text) * Convert.ToDouble(preco31.Text);
                        
                        cmd_item31 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item31.Text + "', '" + qdt31.Text + "', '" + preco31.Text + "', '" + vlTotalItem31 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item31.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item32.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem32 = Convert.ToDouble(qdt32.Text) * Convert.ToDouble(preco32.Text);
                        
                        cmd_item32 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item32.Text + "', '" + qdt32.Text + "', '" + preco32.Text + "', '" + vlTotalItem32 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item32.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item33.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem33 = Convert.ToDouble(qdt33.Text) * Convert.ToDouble(preco33.Text);
                        
                        cmd_item33 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item33.Text + "', '" + qdt33.Text + "', '" + preco33.Text + "', '" + vlTotalItem33 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item33.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item34.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem34 = Convert.ToDouble(qdt34.Text) * Convert.ToDouble(preco34.Text);
                        
                        cmd_item34 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item34.Text + "', '" + qdt34.Text + "', '" + preco34.Text + "', '" + vlTotalItem34 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item34.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item35.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem35 = Convert.ToDouble(qdt35.Text) * Convert.ToDouble(preco35.Text);
                        
                        cmd_item35 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item35.Text + "', '" + qdt35.Text + "', '" + preco35.Text + "', '" + vlTotalItem35 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item35.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item36.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem36 = Convert.ToDouble(qdt36.Text) * Convert.ToDouble(preco36.Text);
                        
                        cmd_item36 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item36.Text + "', '" + qdt36.Text + "', '" + preco36.Text + "', '" + vlTotalItem36 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item36.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item37.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem37 = Convert.ToDouble(qdt37.Text) * Convert.ToDouble(preco37.Text);
                        
                        cmd_item37 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item37.Text + "', '" + qdt37.Text + "', '" + preco37.Text + "', '" + vlTotalItem37 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item37.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item38.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem38 = Convert.ToDouble(qdt38.Text) * Convert.ToDouble(preco38.Text);
                        
                        cmd_item38 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item38.Text + "', '" + qdt38.Text + "', '" + preco38.Text + "', '" + vlTotalItem38 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item38.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item39.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem39 = Convert.ToDouble(qdt39.Text) * Convert.ToDouble(preco39.Text);
                        
                        cmd_item39 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item39.Text + "', '" + qdt39.Text + "', '" + preco39.Text + "', '" + vlTotalItem39 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item39.ExecuteNonQuery();
                    }
                    if (string.IsNullOrEmpty(Item40.Text))
                    {

                    }
                    else
                    {
                        double vlTotalItem40 = Convert.ToDouble(qdt40.Text) * Convert.ToDouble(preco40.Text);
                        cmd_item40 = new MySqlCommand("INSERT INTO `tb_nota` (`ID_TB`, `DANFE`, `DT_EMISSAO`, `COD_PRODUTO`, `QDT`, `VL_UNITARIO`, `VL_TOTAL`, `CENTRO_CUSTO`, `PEDIDO`, `MIGO`, `MIRO`,  `CHAVE_ACESSO`,  `CHAVE_ACESS8`,  `PROTOCOLO`) VALUES (NULL, '" + txtNfe.Text + "-001', '" + this.dtEmissao.Text + "', '" + Item40.Text + "', '" + qdt40.Text + "', '" + preco40.Text + "', '" + vlTotalItem40 + "', NULL, NULL, NULL, NULL, '" + txtChave.Text + "', '" + txtChaveUltimo.Text + "', '" + txtProtocolo.Text + "')", CONEXAOBD);
                        cmd_item40.ExecuteNonQuery();
                    }

                    CONEXAOBD.Close();

                    MessageBox.Show("Salvo com Sucesso!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    Ocultar();
                    Limpar();
                   
                }
                catch (MySqlException Err)
                {
                    MessageBox.Show(Err.Message);
                }
            }
            catch(Exception Erro)
            {
                MessageBox.Show(Erro.Message);
            }
        }
    }
}
