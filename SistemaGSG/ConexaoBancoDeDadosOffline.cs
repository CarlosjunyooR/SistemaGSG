using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SistemaGSG
{
    internal class ConexaoBancoDeDadosOffline
    {
        public static OleDbConnection DBSGSG_Conex()
        {
            string sql = System.Configuration.ConfigurationManager.ConnectionStrings["DBSGSG"].ConnectionString;
            OleDbConnection CONEXOFF = new OleDbConnection(sql);
            CONEXOFF.Open();
            return CONEXOFF;
        }
    }
}
