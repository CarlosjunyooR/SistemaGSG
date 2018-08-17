using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data.Odbc;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using MySql.Data.MySqlClient;
using System.Windows.Forms;
using System.Data;

namespace SistemaGSG
{
    public class Cls_Mysql
    {
        #region atributos
        // atributos //
        private string string_conexao = "Data Source=x.x.x.x;Initial Catalog=db_base_clientes;User ID=base_clientes;Password=xxxxx";
        public string query_string = "";
        #endregion
        #region metodos
        // metodos //
        public MySqlDataReader mysql_data_reader()
        {
            MySqlConnection conexao = new MySqlConnection();
            conexao.ConnectionString = this.string_conexao;
            conexao.Open();

            MySqlCommand comando = new MySqlCommand();
            comando.CommandText = query_string;
            comando.Connection = conexao;

            MySqlDataReader reader = comando.ExecuteReader();

            return reader;
        }

        public DataTable mysql_data_adapter()
        {
            DataTable dtb = new DataTable();

            MySqlConnection conexao = new MySqlConnection();
            conexao.ConnectionString = this.string_conexao;
            try
            {
                conexao.Open();
                MySqlDataAdapter adapter = new MySqlDataAdapter(query_string, conexao);

                adapter.Fill(dtb);

                conexao.Dispose();
                adapter.Dispose();
            }
            catch
            {
            }
            return dtb;
        }

        public bool execute_non_query()
        {
            try
            {
                MySqlConnection conexao = new MySqlConnection();
                conexao.ConnectionString = this.string_conexao;
                conexao.Open();

                MySqlCommand comando = new MySqlCommand();
                comando.CommandText = query_string;
                comando.Connection = conexao;
                comando.ExecuteNonQuery();

                conexao.Dispose();
                comando.Dispose();
                return true;
            }
            catch
            {
                return false;
            }
        }
        #endregion
    }
}
