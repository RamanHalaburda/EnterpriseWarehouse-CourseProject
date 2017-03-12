using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;

namespace CP_2_2
{
    class Bookings 
    {
        private int ID_booking;
	    private int ID_client;
	    private int ID_product;
        private int sum;
        private int quantity;
        private DateTime date;

        public Bookings(int _ID_booking, int _ID_client, int _ID_product, int _sum, int _quantity, DateTime _date)
	    {
		    ID_booking = _ID_booking;
		    ID_client = _ID_client;
		    ID_product = _ID_product;
		    sum = _sum;
            quantity = _quantity;
            date = _date;
	    }

        public void insertBooking()
        {
            string query = "INSERT INTO STORE_BOOKINGS VALUES(" 
                         + ID_booking + "," 
                         + ID_client + "," 
                         + ID_product + "," 
                         + sum + "," 
                         + quantity + ",TO_DATE('" 
                         + date.Year.ToString() + "." 
                         + date.Month.ToString() + "." 
                         + date.Day.ToString() + "', 'yyyy/mm/dd hh24:mi:ss'))";
            if (MyOracleConnect.Insert(query))
            {
                System.Windows.Forms.MessageBox.Show("Запись о продаже добавлена.");
                MyOracleConnect.Update("commit");
            }
        }

        public static void printDataGridTitle(System.Windows.Forms.DataGridView dgv)
        {
            dgv.Columns[0].HeaderCell.Value = "ИД продажи";
            dgv.Columns[1].HeaderCell.Value = "ИД товара";
            dgv.Columns[2].HeaderCell.Value = "ИД клиента";
            dgv.Columns[3].HeaderCell.Value = "Сумма";
            dgv.Columns[4].HeaderCell.Value = "Количество";
            dgv.Columns[5].HeaderCell.Value = "Дата продажи";
        }

        public static void demo_printDataGridTitle(System.Windows.Forms.DataGridView dgv)
        {
            dgv.Columns[0].HeaderCell.Value = "ИД продажи";
            dgv.Columns[1].HeaderCell.Value = "ИД товара"; 
            dgv.Columns[2].HeaderCell.Value = "ИД клиента";
        }

        public void printCheck(string _name_product, string _fio)
        {
            int columns = 2;
            int rows = 8;

            Microsoft.Office.Interop.Word.Application applictaion = new Microsoft.Office.Interop.Word.Application();
            Object missing = Type.Missing;
            applictaion.Documents.Add(ref missing, ref missing, ref missing, ref missing);

            Microsoft.Office.Interop.Word.Document document = applictaion.ActiveDocument;

            Paragraph p = document.Content.Paragraphs.Add(ref missing);
            p.Range.Text = "Накладная поступления товара:";
            p.Range.InsertParagraphAfter();
            Table table = document.Tables.Add(p.Range, rows, columns, ref missing, ref missing);

            table.Borders.Enable = 1;
            table.Cell(1, 1).Range.Text = "ИД поставки";
            table.Cell(2, 1).Range.Text = "ИД товара";
            table.Cell(3, 1).Range.Text = "Название товара";
            table.Cell(4, 1).Range.Text = "ИД клиента";
            table.Cell(5, 1).Range.Text = "ФИО клиента";
            table.Cell(6, 1).Range.Text = "Общая сумма";
            table.Cell(7, 1).Range.Text = "Продано(шт.)";
            table.Cell(8, 1).Range.Text = "Дата продажи";
            table.Cell(1, 2).Range.Text = ID_booking.ToString();
            table.Cell(2, 2).Range.Text = ID_product.ToString();
            table.Cell(3, 2).Range.Text = _name_product.ToString();
            table.Cell(4, 2).Range.Text = ID_client.ToString();
            table.Cell(5, 2).Range.Text = _fio.ToString();
            table.Cell(6, 2).Range.Text = sum.ToString();
            table.Cell(7, 2).Range.Text = quantity.ToString();
            table.Cell(8, 2).Range.Text = date.ToString();

            applictaion.Visible = true;
            try
            {
                object fileName = "e:\\study\\(1) Учёба\\БарГУ Инж.Фак\\"
                                + "Курсовое проектирование\\CP_2_2\\CP_2_2\\Checks\\check_"
                                + ID_booking.ToString();
                document.SaveAs2(fileName);
            }
            catch (Exception) { }
        }
    };
}
