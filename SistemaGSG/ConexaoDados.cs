using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace SistemaGSG
{
    public class ConexaoDados
    {
        String Server = "localhost";
        public static MySqlConnection GetConnectionXML()
        {

            MySqlConnection CONEX = new MySqlConnection();
            CONEX.ConnectionString = @"server = 'localhost'; database='sistemagsgxml';Uid='energia';Pwd='02984646#Lua';SslMode=none;";
            CONEX.Open();

            return CONEX;
        }
        public static MySqlConnection GetConnectionFaturameto()
        {
            MySqlConnection CONEX = new MySqlConnection();
            CONEX.ConnectionString = @"server = 'localhost'; database='sistemagsgfaturamento';Uid='energia';Pwd='02984646#Lua';SslMode=none;";
            CONEX.Open();

            return CONEX;
        }
        public static MySqlConnection GetConnectionFornecedor()
        {
            MySqlConnection CONEX = new MySqlConnection();
            CONEX.ConnectionString = @"server = 'localhost'; database='sistemagsgfornecedor';Uid='energia';Pwd='02984646#Lua';SslMode=none;";
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
    }
}
