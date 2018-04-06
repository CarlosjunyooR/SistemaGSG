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



namespace SistemaGSG
{
    class conexao
    {
        //Conexão com Mysql//

        private static MySqlConnection conn_1 = null;
        public static MySqlConnection obterConexao_1()
        {
            string connString = "Database=ceal1;Data Source=localhost;User Id=junior;Password=vertrigo";
            conn_1 = new MySqlConnection(connString);
            try
            {
                conn_1.Open();
            }
            catch(MySqlException sqle)
            {
                conn_1 = null;
                System.Windows.Forms.MessageBox.Show("Erro ao conectar ao banco de dados.\n" + sqle, "Mensagem", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            return conn_1;
        }

        public static void fechaConexao_1()
        {
            if(conn_1 != null)
            {
                conn_1.Close();
            }
        }
        //Conexão com Access





        //Conexão com SQL//
        //static public string conex = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\junio\OneDrive\RECIBOS\repos\SistemaGSG\SistemaGSG\bin\Debug\SistemaGSG.mdf;Integrated Security=True;Connect Timeout=30";
    }
}
