using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Common;
using Oracle.DataAccess.Client;
using System.Data;

namespace GHIAProj
{
   public enum Operations
    {
        HIGHPOLLUTION,
        EXTREMEPOLLUTION,
        BORDER_HIGHPOLLUTION,
        BORDER_EXTREMEPOLLUTION
    }
   public class TableCreater
    {
        public DataTable GetHight(Operations op, int _minKod, int _maxKod, DateTime _startDate, DateTime _endDate)
        {
            OracleConnection conn = DBUtils.GetDBConnection();
            conn.Open();
            OracleDataAdapter da = new OracleDataAdapter();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            cmd.CommandText = op.ToString();
            cmd.CommandType= CommandType.StoredProcedure;
            cmd.Parameters.Add("minKod", OracleDbType.Int32).Value = _minKod;
            cmd.Parameters.Add("maxKod", OracleDbType.Int32).Value = _maxKod;           
            cmd.Parameters.Add("startDate", OracleDbType.Date).Value = _startDate;
            cmd.Parameters.Add("endDate", OracleDbType.Date).Value = _endDate;
            cmd.Parameters.Add("c1", OracleDbType.RefCursor).Direction = ParameterDirection.Output;            
            da.SelectCommand = cmd;
            DataTable dt = new DataTable();
            da.Fill(dt);
            conn.Close();
            return dt;
        }
        public void CreateDataColumn(ref DataTable dt)
        {
            dt.Columns.Add(new DataColumn("Источники загрязнения", typeof(string)));
        } 
    }
}

