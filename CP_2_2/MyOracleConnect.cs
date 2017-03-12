using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Oracle.DataAccess.Client;
using System.Windows.Forms;

namespace CP_2_2
{
    class MyOracleConnect
    {
        private static string connection = "Data Source=(DESCRIPTION="
                    + "(ADDRESS_LIST=" + "(ADDRESS=" + "(PROTOCOL=TCP)"
                    + "(HOST=localhost)" + "(PORT=1521)" + ")" + ")"
                    + "(CONNECT_DATA=" + "(SERVER=DEDICATED)"
                    + "(SERVICE_NAME=XE)" + ")" + ");" + "User Id=SYSTEM;Password=1;";

        public static DataSet Select(string query)
        {
            OracleConnection conn = new OracleConnection(connection);
            DataSet result = new DataSet();
            try
            {
                conn.Open();
                OracleCommand command = new OracleCommand(query, conn);
                OracleDataAdapter adapter = new OracleDataAdapter(command);
                adapter.Fill(result);
            }
            catch (OracleException se)
            {
                MessageBox.Show(se.Message);
            }
            finally
            {
                if (conn.State != ConnectionState.Closed)
                    conn.Close();
            }
            return result;
        }

        public static bool Insert(string query)
        {
            OracleConnection conn = new OracleConnection(connection);
            bool result = new bool();
            DataTable res = new DataTable();
            try
            {
                conn.Open();
                OracleCommand command = new OracleCommand(query, conn);
                OracleDataAdapter adapter = new OracleDataAdapter(command);
                adapter.Fill(res);
                result = true;
            }
            catch (OracleException se)
            {
                MessageBox.Show(se.Message);
                result = false;
            }
            finally
            {
                if (conn.State != ConnectionState.Closed)
                    conn.Close();
            }
            return result;
        }

        public static bool Delete(string query)
        {
            OracleConnection conn = new OracleConnection(connection);
            bool result = new bool();
            DataTable res = new DataTable();
            try
            {
                conn.Open();
                OracleCommand command = new OracleCommand(query, conn);
                OracleDataAdapter adapter = new OracleDataAdapter(command);
                adapter.Fill(res);
                result = true;
            }
            catch (OracleException se)
            {
                MessageBox.Show(se.Message);
                result = false;
            }
            finally
            {
                if (conn.State != ConnectionState.Closed)
                    conn.Close();
            }
            return result;
        }

        public static bool Update(string query)
        {
            OracleConnection conn = new OracleConnection(connection);
            bool result = new bool();
            DataTable res = new DataTable();
            try
            {
                conn.Open();
                OracleCommand command = new OracleCommand(query, conn);
                OracleDataAdapter adapter = new OracleDataAdapter(command);
                adapter.Fill(res);
                result = true;
            }
            catch (OracleException se)
            {
                MessageBox.Show(se.Message);
                result = false;
            }
            finally
            {
                if (conn.State != ConnectionState.Closed)
                    conn.Close();
            }
            return result;
        }
    }
}