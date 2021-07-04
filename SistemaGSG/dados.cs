using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace SistemaGSG
{
    class dados
    {
        internal static string completo;
        static public string usuario { get; set; }
        static public string tema { get; set; }
        static public string senha { get; set; }
        static public Int32 nivel { get; set; }
    }
}
