using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CP_2_2
{
    class Clients
    {
        private	int ID_client;
	    private string fio;
	    private int clientTelephone;
        
        public Clients(int _ID_client, string _fio, int _clientTelephone)
	    {
		    ID_client = _ID_client;
		    fio = _fio;
		    clientTelephone = _clientTelephone;
	    }

        public void addClient()
        {
            string query = "insert into store_clients values("
                         + ID_client + ",'"
                         + fio + "',"
                         + clientTelephone + ")";
            if (MyOracleConnect.Insert(query))
            {
                System.Windows.Forms.MessageBox.Show("Запись о клиента добавлена.");
                MyOracleConnect.Update("commit");
            }
        }

        public static void printDataGridTitle(System.Windows.Forms.DataGridView dgv)
        {
            dgv.Columns[0].HeaderCell.Value = "ИД клиента";
            dgv.Columns[1].HeaderCell.Value = "ФИО";
            dgv.Columns[2].HeaderCell.Value = "Номер телефона";
        }

        public static void demo_printDataGridTitle(System.Windows.Forms.DataGridView dgv)
        {
            dgv.Columns[0].HeaderCell.Value = "ИД клиента";
            dgv.Columns[1].HeaderCell.Value = "ФИО";
        }
    }
}
