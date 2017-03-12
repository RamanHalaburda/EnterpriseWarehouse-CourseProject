using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CP_2_2
{
    class Distributors
    {
        private int ID_distributor;
        private string distributorName;
	    private int distributorTelephone;

        public Distributors(int _ID_distributor, string _distributorName, int _distributorTelephone)
	    {
            ID_distributor = _ID_distributor;
            distributorName = _distributorName;
		    distributorTelephone = _distributorTelephone;
	    }
    
        public void addDistributor()
        {
            string query = "insert into store_distributors values(" 
                         + ID_distributor + ",'"
                         + distributorName + "',"
                         + distributorTelephone + ")";
            if (MyOracleConnect.Insert(query))
            {
                System.Windows.Forms.MessageBox.Show("Запись о цехе добавлена.");
                MyOracleConnect.Update("commit");
            }
        }

        public static void printDataGridTitle(System.Windows.Forms.DataGridView dgv)
        {
            dgv.Columns[0].HeaderCell.Value = "ИД цеха";
            dgv.Columns[1].HeaderCell.Value = "Название";
            dgv.Columns[2].HeaderCell.Value = "Номер телефона";
        }
    }
}
