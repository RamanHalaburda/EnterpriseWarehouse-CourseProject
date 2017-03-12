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
    class Products
    {
        private int ID_product;
	    private int ID_distrib;
	    private string name_pr;
	    private int cost;
	    private DateTime dateDistribution;
	    private int storage;
        private int existence;

        public Products(int _ID_product, int _ID_distrib, string _name_pr, int _cost, DateTime _dateDistribution, int _storage ,int _existence)
	    {
		    ID_product = _ID_product;
		    ID_distrib = _ID_distrib;
            name_pr = _name_pr;
		    cost = _cost;
		    dateDistribution = _dateDistribution;
            storage = _storage;
            existence = _existence;
	    }

        public void addEmptyProduct()
        {
            string query = "INSERT INTO STORE_PRODUCTS VALUES(" 
                         + ID_product + ","
                         + ID_distrib + ",'"
                         + name_pr + "',"
                         + cost + "," + "TO_DATE('"
                         + dateDistribution.Year.ToString() + "." 
                         + dateDistribution.Month.ToString() + "." 
                         + dateDistribution.Day.ToString()
                         + "', 'yyyy/mm/dd hh24:mi:ss'),"
                         + storage + ","
                         + existence + ")";
            if (MyOracleConnect.Insert(query))
            {
                System.Windows.Forms.MessageBox.Show("Запись о товаре добавлена.");
                MyOracleConnect.Update("commit");
            }
        }

        public void fillProduct(int _quantityAdded)
        {
            string query = "update store_products set date_pr = TO_DATE('"
                         + dateDistribution.Year.ToString() + "."
                         + dateDistribution.Month.ToString() + "."
                         + dateDistribution.Day.ToString()
                         + "', 'yyyy/mm/dd hh24:mi:ss'), existence = "
                         + (existence + _quantityAdded)
                         + ", id_distrib_fk = " + ID_distrib
                         + " where id_product = " + ID_product;
            if (MyOracleConnect.Insert(query))
            {
                System.Windows.Forms.MessageBox.Show("Запись о товаре обновлена.");
                MyOracleConnect.Update("commit");
            }
        }

        static public void subtractProduct(int _ID_product, int _existence, int _quantitySubtracted)
        {
            string query = "update store_products set existence = "
                         + (_existence - _quantitySubtracted) 
                         + " where id_product = " + _ID_product;
            if (MyOracleConnect.Insert(query))
            {
                System.Windows.Forms.MessageBox.Show("Запись о товаре обновлена.");
                MyOracleConnect.Update("commit");
            }
        }
        /*
        public static void checkNeedProducts(System.Windows.Forms.DataGridView dgv)
        {
            OracleConnection oc;
            DataSet ds = new DataSet();;
            OracleDataAdapter oda = new OracleDataAdapter("select * from store_products;", oc);
            oda.Fill(ds);
            dgv.DataSource = ds.Tables[0];
        }*/
        
        public static void printDataGridTitle(System.Windows.Forms.DataGridView dgv)
        {
            dgv.Columns[0].HeaderCell.Value = "ИД товара";
            dgv.Columns[1].HeaderCell.Value = "ИД поставщика";
            dgv.Columns[2].HeaderCell.Value = "Название";
            dgv.Columns[3].HeaderCell.Value = "Стоимость (шт.)";
            dgv.Columns[4].HeaderCell.Value = "Дата поставки";
            dgv.Columns[5].HeaderCell.Value = "Стеллаж";
            dgv.Columns[6].HeaderCell.Value = "В наличии (шт.)";
        }

        public static void demoPrintDataGridTitle(System.Windows.Forms.DataGridView dgv)
        {
            dgv.Columns[0].HeaderCell.Value = "ИД товара";
            dgv.Columns[1].HeaderCell.Value = "Название";
        }

        public void printInvoice(int _quantityAdded)
        {                
            int columns = 2; 
            int rows = 7; 
            
            Microsoft.Office.Interop.Word.Application applictaion = new Microsoft.Office.Interop.Word.Application();
            Object missing = Type.Missing; 
            applictaion.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            
            Microsoft.Office.Interop.Word.Document document = applictaion.ActiveDocument;

            Paragraph p = document.Content.Paragraphs.Add(ref missing); 
            p.Range.Text = "Накладная поступления товара:"; 
            p.Range.InsertParagraphAfter(); 
            Table table = document.Tables.Add(p.Range, rows, columns, ref missing, ref missing);            

            table.Borders.Enable = 1; 
            table.Cell(1, 1).Range.Text = "ИД товара"; 
            table.Cell(2, 1).Range.Text = "ИД цеха"; 
            table.Cell(3, 1).Range.Text = "Название"; 
            table.Cell(4, 1).Range.Text = "Стоимость(шт.)"; 
            table.Cell(5, 1).Range.Text = "Дата поставки"; 
            table.Cell(6, 1).Range.Text = "Стеллаж"; 
            table.Cell(7, 1).Range.Text = "Поставлено(шт.)";
            table.Cell(1, 2).Range.Text = ID_product.ToString();
            table.Cell(2, 2).Range.Text = ID_distrib.ToString();
            table.Cell(3, 2).Range.Text = name_pr;
            table.Cell(4, 2).Range.Text = cost.ToString();
            table.Cell(5, 2).Range.Text = dateDistribution.ToString();
            table.Cell(6, 2).Range.Text = storage.ToString();
            table.Cell(7, 2).Range.Text = _quantityAdded.ToString();

            applictaion.Visible = true;
            try 
            { 
                object fileName = "e:\\study\\(1) Учёба\\БарГУ Инж.Фак\\"
                                + "Курсовое проектирование\\CP_2_2\\CP_2_2\\Invoices\\invoice_" 
                                + ID_product.ToString();
                document.SaveAs2(fileName);
            }
            catch (Exception) {}
        }
    }
}