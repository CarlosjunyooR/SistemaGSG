/* 
 * Classe acesso MySql
 * 
 * Visite nossa página http://www.codigoexpresso.com.br
 * 
 * by Antonio Azevedo
 *  
 * Instanciando a Classe
 *    Conexao(TipoConexao.Conexao.WebConfig) -> Lendo string conexao do arquivo Web.Config
 *    Conexao(TipoConexao.Conexao.Classe)    -> Lendo dados da conexao de nossa classe
 *
 * Configurando nossa Classe
 *    ConexaoWebConfig - > Nome da connectionStrings de nosso arquivo Web.Config
 *    Server           - > Nome do servidor, se estiver usando conexao local utilize ´localhost´
 *    Database         - > Nome do Banco de Dados
 *    Usuario          - > Nome do usuário de Acesso ao Banco de Dados
 *    Senha            - > Senha de Acesso ao Banco de Dados
 * 
 */
namespace System.Web
{
    internal class Configuration
    {
        public static object WebConfigurationManager { get; internal set; }
    }
}