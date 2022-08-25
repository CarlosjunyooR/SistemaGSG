using MySql.Data.MySqlClient;
using System.Data.OleDb;
using System;

namespace SistemaGSG
{
    public class ConexaoDados
    {
        public static MySqlConnection GetConnectionXML()
        {
            MySqlConnection CONEX = new MySqlConnection();
            CONEX.ConnectionString = @"server = 'localhost'; database='sistemagsgxml';Uid='xml';Pwd='02984646#Lua';SslMode=none;";
            CONEX.Open();

            return CONEX;
        }
        public static string ACESSO()
        {
            String ACESSO = "http://localhost/sistemagsgv2.0/template/dashboard/pages/relatorios/faturamento/acesso/RelatorioAcesso.php";

            return ACESSO;
        }
        public static string CHECKLIST()
        {
            String CHECKLIST = "http://localhost/sistemagsgv2.0/template/dashboard/pages/relatorios/faturamento/acesso/RelatorioCheckList.php";

            return CHECKLIST;
        }

        public static MySqlConnection GetConnectionFaturameto()
        {
            MySqlConnection CONEX = new MySqlConnection();
            CONEX.ConnectionString = @"server = 'localhost'; database='sistemagsgfaturamento';Uid='faturamento';Pwd='02984646#Lua';SslMode=none;";
            CONEX.Open();

            return CONEX;
        }

        public static MySqlConnection GetConnectionFornecedor()
        {
            MySqlConnection CONEX = new MySqlConnection();
            CONEX.ConnectionString = @"server = 'localhost'; database='sistemagsgfornecedor';Uid='fornecedor';Pwd='02984646#Lua';SslMode=none;";
            CONEX.Open();

            return CONEX;
        }
        public static MySqlConnection GetConnectionEquatorial()
        {
            MySqlConnection CONEX = new MySqlConnection();
            CONEX.ConnectionString = @"server = 'localhost'; database='sistemagsgequatorial';Uid='energia';Pwd='02984646#Lua';SslMode=none;";
            CONEX.Open();

            return CONEX;
        }
        public static MySqlConnection GetConnectionPosto()
        {
            MySqlConnection CONEX = new MySqlConnection();
            CONEX.ConnectionString = @"server = 'localhost'; database='sistemagsgposto';Uid='posto';Pwd='02984646#Lua';SslMode=none;";
            CONEX.Open();

            return CONEX;
        }
    }
}
