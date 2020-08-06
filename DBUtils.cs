using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.DataAccess.Client;

namespace GHIAProj
{
    class DBUtils
    {
        public static OracleConnection GetDBConnection()
        {
            string host = "****";
            int port = 1521;
            string sid = "ORCL";
            string user = "****";
            string password = "****";

            return DBOracleUtils.GetDBConnection(host, port, sid, user, password);
        }
    }
}
