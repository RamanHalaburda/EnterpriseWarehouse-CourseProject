using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;

namespace CP_2_2
{
    public partial class Form1 : Form
    {
        OracleConnection oc;
        DataSet ds;
        OracleDataAdapter oda;
        string oradb = "";
        
        public Form1() { InitializeComponent(); }
// load main form
        private void Form1_Load(object sender, EventArgs e)
        {
            oradb = "Data Source=localhost:1521/XE; User Id=SYSTEM;Password=1"; //; DBA Privilege=default
            oc = new OracleConnection(oradb);
            oc.Open();
            ds = new DataSet();
            comboBox1.SelectedIndex = 0;
            textBox7.Text = System.Convert.ToString(10);
            MyOracleConnect.Update("commit");            
        }
/************************************** tabPage 1 **************************************/
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                ds = new DataSet();
                oda = new OracleDataAdapter("select * from store_products order by id_product", oc);
                oda.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
                Products.printDataGridTitle(dataGridView1);
            }
            if (comboBox1.SelectedIndex == 1)
            {
                ds = new DataSet();
                oda = new OracleDataAdapter("select * from store_clients order by id_client", oc);
                oda.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
                Clients.printDataGridTitle(dataGridView1);                
            }
            if (comboBox1.SelectedIndex == 2)
            {
                ds = new DataSet();
                oda = new OracleDataAdapter("select * from store_distributors order by id_distrib", oc);
                oda.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
                Distributors.printDataGridTitle(dataGridView1);
            }
            if (comboBox1.SelectedIndex == 3)
            {
                ds = new DataSet();
                oda = new OracleDataAdapter("select * from store_bookings order by id_booking", oc);
                oda.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
                Bookings.printDataGridTitle(dataGridView1);
            }
        }
// update table
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                OracleCommandBuilder ocb = new OracleCommandBuilder(oda);
                oda.Update(ds, "table");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }
/************************************** tabPage 2 **************************************/
        private void tabPage2_Enter(object sender, EventArgs e)
        {
            ds = new DataSet();
            oda = new OracleDataAdapter("select * from store_products order by name_pr", oc);
            oda.Fill(ds);
            dataGridView2.DataSource = ds.Tables[0];
            Products.printDataGridTitle(dataGridView2);

            ds = new DataSet();
            oda = new OracleDataAdapter("select * from store_distributors order by title", oc);
            oda.Fill(ds);
            dataGridView3.DataSource = ds.Tables[0];
            Distributors.printDataGridTitle(dataGridView3);
            
            textBox12.Text = DateTime.Today.ToString();
            textBox12.Enabled = false;
            
            textBox9.Enabled = true;                        
            textBox9.Text = "";   
            textBox14.Text = "";            
        }
// fill product
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox8.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString();
            textBox10.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[2].Value.ToString();
            textBox11.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[3].Value.ToString();
            textBox12.Text = DateTime.Today.ToString();
            textBox13.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[5].Value.ToString();
            textBox8.Enabled = false;
            textBox10.Enabled = false;
            textBox11.Enabled = false;
            textBox12.Enabled = false;
            textBox13.Enabled = false;
        }
// fill distributor
        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox9.Text = dataGridView3.Rows[e.RowIndex].Cells[0].Value.ToString();
            textBox9.Enabled = false;
        }
// add in booking
        private void button3_Click(object sender, EventArgs e)
        {

            if (textBox8.Text != "" &&
                textBox9.Text != "" &&
                textBox14.Text != "")
            {
                int quantityAdded = System.Convert.ToInt32(textBox14.Text);
                Products p = new Products(System.Convert.ToInt32(textBox8.Text),
                                          System.Convert.ToInt32(textBox9.Text),
                                          textBox10.Text,
                                          System.Convert.ToInt32(textBox11.Text),
                                          System.Convert.ToDateTime(textBox12.Text),
                                          System.Convert.ToInt32(textBox13.Text),
                                          System.Convert.ToInt32(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[6].Value));
                p.fillProduct(quantityAdded);

                if (checkBox1.Checked == true)
                    p.printInvoice(quantityAdded);
                tabPage2_Enter(this, EventArgs.Empty);
            }
            else
                System.Windows.Forms.MessageBox.Show("Ошибка! Заполнены не все поля.");

            textBox10.Enabled = true;
            textBox11.Enabled = true;
            textBox13.Enabled = true;
            textBox8.Enabled = true;
            textBox8.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox13.Text = "";
            textBox14.Text = "";            
        }
/************************************** tabPage 3 **************************************/
        private void tabPage3_Enter(object sender, EventArgs e)
        {
            ds = new DataSet();
            oda = new OracleDataAdapter("select * from store_products order by name_pr", oc);
            oda.Fill(ds);
            dataGridView4.DataSource = ds.Tables[0];            
            Products.printDataGridTitle(dataGridView4);

            ds = new DataSet();
            oda = new OracleDataAdapter("select id_client, fio from store_clients order by fio", oc);
            oda.Fill(ds);
            dataGridView6.DataSource = ds.Tables[0];
            Clients.demo_printDataGridTitle(dataGridView6);
            
            ds = new DataSet();
            oda = new OracleDataAdapter("select id_booking,id_product_fk,id_client_fk from store_bookings order by id_booking", oc);
            oda.Fill(ds);
            dataGridView5.DataSource = ds.Tables[0];
            Bookings.demo_printDataGridTitle(dataGridView5);
            
            string temp_s = dataGridView5.Rows[dataGridView5.RowCount - 2].Cells[0].Value.ToString();
            int temp_i = System.Convert.ToInt32(temp_s);
            ++temp_i;
            textBox1.Text = System.Convert.ToString(temp_i);
            textBox1.Enabled = false;

            textBox6.Text = DateTime.Today.ToString();
            textBox6.Enabled = false;

            textBox2.Enabled = true;
            textBox3.Enabled = true;
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
        }
// fill client
        private void dataGridView6_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox2.Text = dataGridView6.Rows[e.RowIndex].Cells[0].Value.ToString();
            textBox2.Enabled = false;
        }
// fill product
        private void dataGridView4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox3.Text = dataGridView4.Rows[e.RowIndex].Cells[0].Value.ToString();
            textBox3.Enabled = false;
            textBox4.Text = "";
            textBox5.Text = "";
        }
// add out booking
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox2.Text != "" &&
                textBox3.Text != "" &&
                textBox4.Text != "" &&
                textBox5.Text != "" &&
                textBox6.Text != "")
            {                
                Bookings b = new Bookings(System.Convert.ToInt32(textBox1.Text),
                            System.Convert.ToInt32(textBox2.Text),
                            System.Convert.ToInt32(textBox3.Text),
                            System.Convert.ToInt32(textBox4.Text),
                            System.Convert.ToInt32(textBox5.Text),
                            System.Convert.ToDateTime(textBox6.Text));
                b.insertBooking();
                int _existence = System.Convert.ToInt32(dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[6].Value);
                Products.subtractProduct(System.Convert.ToInt32(textBox3.Text),
                                            _existence,
                                            System.Convert.ToInt32(textBox5.Text));
                if (checkBox2.Checked == true)
                    b.printCheck(dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[2].Value.ToString(),
                                 dataGridView6.Rows[dataGridView6.CurrentCell.RowIndex].Cells[1].Value.ToString());
                tabPage3_Enter(this, EventArgs.Empty);                
            }
            else
                System.Windows.Forms.MessageBox.Show("Ошибка! Заполнены не все поля.");

        }
// calculate sum for out booking and print
        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            if (textBox5.Text != "")
            {
                for (int i = 0; System.Convert.ToInt32(dataGridView4.Rows[i].Cells[0].Value)
                        != System.Convert.ToInt32(textBox3.Text); ++i)
                    textBox4.Text = System.Convert.ToString(System.Convert.ToInt32(textBox5.Text) *
                                    System.Convert.ToInt32(dataGridView4.Rows[i + 1].Cells[3].Value));

                if (System.Convert.ToInt32(textBox5.Text) >
                        System.Convert.ToInt32(dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[6].Value))
                {
                    System.Windows.Forms.MessageBox.Show("Ошибка! Такого количеста товара нет на складе.");
                    textBox4.Text = "";
                    textBox5.Text = "";
                }
            }
        }
// clear two fields
        private void textBox5_Click(object sender, EventArgs e)
        {
            textBox4.Text = "";
            textBox5.Text = "";
        }
/************************************** tabPage 4 **************************************/
        private void tabPage4_Enter(object sender, EventArgs e)
        {
            string temp_s = "";
            int temp_i = 0;

            comboBox1.SelectedItem = "Товары";
            temp_s = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[0].Value.ToString();
            temp_i = System.Convert.ToInt32(temp_s);
            ++temp_i;
            textBox15.Text = System.Convert.ToString(temp_i);
            textBox15.Enabled = false;

            comboBox1.SelectedItem = "Цехи";
            temp_s = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[0].Value.ToString();
            temp_i = System.Convert.ToInt32(temp_s);
            ++temp_i;
            textBox22.Text = System.Convert.ToString(temp_i);
            textBox22.Enabled = false;

            comboBox1.SelectedItem = "Клиенты";
            temp_s = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[0].Value.ToString();
            temp_i = System.Convert.ToInt32(temp_s);
            ++temp_i;
            textBox25.Text = System.Convert.ToString(temp_i);
            textBox25.Enabled = false;

            textBox19.Text = DateTime.Today.ToString();
            textBox19.Enabled = false;

            textBox16.Text = "1";
            textBox20.Text = "0";
            textBox21.Text = "0";
            textBox16.Enabled = false;
            textBox20.Enabled = false;
            textBox21.Enabled = false;
        }
// add product
        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox16.Text != "" && textBox17.Text != "" && textBox18.Text != "")
            {
                Products product = new Products(System.Convert.ToInt32(textBox15.Text),
                                                System.Convert.ToInt32(textBox16.Text),
                                                textBox17.Text,
                                                System.Convert.ToInt32(textBox18.Text),
                                                System.Convert.ToDateTime(textBox19.Text),
                                                System.Convert.ToInt32(textBox20.Text),
                                                System.Convert.ToInt32(textBox21.Text));
                product.addEmptyProduct();
            }
            textBox17.Text = "";
            textBox18.Text = "";
        }
// add distributor
        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox23.Text != "" && textBox24.Text != "")
            {
                Distributors distributor = new Distributors(System.Convert.ToInt32(textBox22.Text),
                                                textBox23.Text,
                                                System.Convert.ToInt32(textBox24.Text));
                distributor.addDistributor();
            }
            int newID = System.Convert.ToInt32(textBox22.Text);
            ++newID;
            textBox22.Text = newID.ToString();
            textBox23.Text = "";
            textBox24.Text = "";
        }        
// add client
        private void button6_Click(object sender, EventArgs e)
        {
            if (textBox26.Text != "" && textBox27.Text != "")
            {
                Clients client = new Clients(System.Convert.ToInt32(textBox25.Text),
                                                 textBox26.Text,
                                                 System.Convert.ToInt32(textBox27.Text));
                client.addClient();
            }
            int newID = System.Convert.ToInt32(textBox25.Text);
            ++newID;
            textBox25.Text = newID.ToString();
            textBox26.Text = "";
            textBox27.Text = "";
        }        
/************************************** tabPage 5 **************************************/
        private void tabPage5_Enter(object sender, EventArgs e)
        {
            ds = new DataSet();
            oda = new OracleDataAdapter("select * from store_products order by id_product", oc);
            oda.Fill(ds);
            dataGridView8.DataSource = ds.Tables[0];
            Products.printDataGridTitle(dataGridView8);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (textBox28.Text != "" && textBox29.Text != "")
            {
                string q = "SELECT * FROM store_products WHERE storage_pr BETWEEN "
                    + System.Convert.ToInt32(textBox28.Text) + " AND "
                    + System.Convert.ToInt32(textBox29.Text) + " order by storage_pr";

                ds = new DataSet();
                oda = new OracleDataAdapter(q, oc);
                oda.Fill(ds);

                dataGridView8.DataSource = ds.Tables[0];
                Products.printDataGridTitle(dataGridView8);
                if (checkBox5.Checked == true)
                {
                    Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook wb = application.Workbooks.Add(System.Reflection.Missing.Value);
                    Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)application.ActiveSheet;
                    string Name = "products(storages " + System.Convert.ToInt32(textBox28.Text) + "-" + System.Convert.ToInt32(textBox29.Text) + ")";
                    ws.Name = Name;                 

                    ws.Cells[1, 1] = "ИД продукта";
                    ws.Cells[1, 2] = "ИД цеха";
                    ws.Cells[1, 3] = "Название";
                    ws.Cells[1, 4] = "Стоимость(шт.)";
                    ws.Cells[1, 5] = "Дата поставки";
                    ws.Cells[1, 6] = "Стеллаж";
                    ws.Cells[1, 7] = "Количество на складе(шт.)";

                    for (int i = 0; i < dataGridView8.Rows.Count; i++)
                        for (int j = 0; j < dataGridView2.ColumnCount; j++)
                            application.Cells[i + 2, j + 1] = dataGridView8.Rows[i].Cells[j].Value;

                    ws.Columns.AutoFit();

                    application.Visible = true;
                    application.UserControl = true;
                    try
                    {
                        string fileName = "e:\\study\\(1) Учёба\\БарГУ Инж.Фак\\"
                                        + "Курсовое проектирование\\CP_2_2\\CP_2_2\\reports_Excel\\report_"
                                        + Name;
                        wb.SaveAs(fileName);
                    }
                    catch (Exception) { }
                }
            }
        }
// output most buying client
        private void button8_Click(object sender, EventArgs e)
        {
            string query_clients = "SELECT DISTINCT * FROM STORE_CLIENTS";
            System.Data.DataTable dt = new System.Data.DataTable();
            
            // print set of id_client in dgv8
            var adapter = new OracleDataAdapter();
            adapter.SelectCommand = new OracleCommand(query_clients, oc);
            adapter.Fill(dt);
            dataGridView8.DataSource = dt;

            // create array of id_client
            int[] arr_id_client = new int[dataGridView8.RowCount];
            string[] arr_fio_client = new string[dataGridView8.RowCount];
            int[] arr_sum = new int[dataGridView8.RowCount];
            for (int i = 0; i < dataGridView8.RowCount - 1; ++i)
            {
                arr_id_client[i] = System.Convert.ToInt32(dataGridView8.Rows[i].Cells[0].Value);
                arr_fio_client[i] = dataGridView8.Rows[i].Cells[1].Value.ToString();
            }

            // query for select sum for any client
            string temp_query = "";
            try
            {
                for (int i = 0; i < arr_id_client.Length - 1; ++i)
                {
                    temp_query = "SELECT SUM(sum_b) from store_bookings WHERE id_client_fk = " + arr_id_client[i];
                    ds = new DataSet();
                    oda = new OracleDataAdapter(temp_query, oc);
                    oda.Fill(ds);
                    dataGridView8.DataSource = ds.Tables[0];
                    arr_sum[i] = System.Convert.ToInt32(dataGridView8.Rows[0].Cells[0].Value);
                }
            }
            catch (Exception) { }

            // print set of id_client in dgv8
            adapter = new OracleDataAdapter();
            dt = new System.Data.DataTable();
            adapter.SelectCommand = new OracleCommand(query_clients, oc);
            adapter.Fill(dt);
            dataGridView8.DataSource = dt;
            for (int i = 0; i < arr_id_client.Length - 1; ++i)
            {
                dataGridView8.Rows[i].Cells[0].Value = System.Convert.ToString(arr_id_client[i]);
                dataGridView8.Rows[i].Cells[1].Value = System.Convert.ToString(arr_fio_client[i]);
                dataGridView8.Rows[i].Cells[2].Value = System.Convert.ToString(arr_sum[i]);
            }
            Clients.demo_printDataGridTitle(dataGridView8);
            dataGridView8.Columns[2].HeaderCell.Value = "Сумма";

            // search most buying client
            int max_sum = arr_sum[0];
            int max_index = 0;
            for (int i = 1; i < arr_id_client.Length - 1; ++i)
                if (arr_sum[i] > max_sum)
                {
                    max_sum = arr_sum[i];
                    max_index = i;
                }

            textBox30.Text = System.Convert.ToString(arr_id_client[max_index]);
            textBox31.Text = System.Convert.ToString(arr_sum[max_index]);
            textBox32.Text = System.Convert.ToString(arr_fio_client[max_index]);

            if (checkBox4.Checked == true)
            {
                Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wb = application.Workbooks.Add(System.Reflection.Missing.Value);
                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)application.ActiveSheet;
                string Name = "most_buying_client";
                ws.Name = Name;

                ws.Cells[1, 1] = "Самый покупающий клиент";
                ws.Cells[2, 1] = "ИД клиента";
                ws.Cells[2, 2] = "ФИО";
                ws.Cells[2, 3] = "Общая сумма покупок";
                ws.Cells[3, 1] = System.Convert.ToString(arr_id_client[max_index]);
                ws.Cells[3, 2] = System.Convert.ToString(arr_fio_client[max_index]);
                ws.Cells[3, 3] = System.Convert.ToString(arr_sum[max_index]);

                ws.Cells[5, 1] = "Общий отчёт по сумме покупок каждого клиента";
                ws.Cells[6, 1] = "ИД клиента";
                ws.Cells[6, 2] = "ФИО";
                ws.Cells[6, 3] = "Общая сумма покупок";

                for (int i = 0; i < dataGridView8.Rows.Count; i++)
                    for (int j = 0; j < dataGridView8.ColumnCount; j++)
                        application.Cells[i + 7, j + 1] = dataGridView8.Rows[i].Cells[j].Value;

                ws.Columns.AutoFit();

                application.Visible = true;
                application.UserControl = true;
                try
                {
                    string fileName = "e:\\study\\(1) Учёба\\БарГУ Инж.Фак\\"
                                    + "Курсовое проектирование\\CP_2_2\\CP_2_2\\reports_Excel\\report_"
                                    + Name;
                    wb.SaveAs(fileName);
                    ws.SaveAs(fileName);
                }
                catch (Exception) { }
            }
        }   
/************************************** tabPage 6 **************************************/
        private void tabPage6_Enter(object sender, EventArgs e)
        {
            if (textBox7.Text != "")
            {
                ds = new DataSet();
                oda = new OracleDataAdapter("select * from store_products where existence < " 
                    + System.Convert.ToInt32(textBox7.Text) + " order by existence asc", oc);
                oda.Fill(ds);
                dataGridView7.DataSource = ds.Tables[0];
            }
            Products.printDataGridTitle(dataGridView7);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox8.Text = dataGridView7.Rows[dataGridView7.CurrentCell.RowIndex].Cells[0].Value.ToString();
            textBox10.Text = dataGridView7.Rows[dataGridView7.CurrentCell.RowIndex].Cells[2].Value.ToString();
            textBox11.Text = dataGridView7.Rows[dataGridView7.CurrentCell.RowIndex].Cells[3].Value.ToString();
            textBox12.Text = DateTime.Today.ToString();
            textBox13.Text = dataGridView7.Rows[dataGridView7.CurrentCell.RowIndex].Cells[5].Value.ToString();
            textBox8.Enabled = false;
            textBox10.Enabled = false;
            textBox11.Enabled = false;
            textBox12.Enabled = false;
            textBox13.Enabled = false;
            tabControl1.SelectedTab = tabPage2;

            if (checkBox6.Checked == true)
            {
                Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wb = application.Workbooks.Add(System.Reflection.Missing.Value);
                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)application.ActiveSheet;
                string Name = "fault_of_products(" + System.Convert.ToInt32(textBox7.Text) + " or less)";
                ws.Name = Name;

                ws.Cells[1, 1] = "ИД продукта";
                ws.Cells[1, 2] = "ИД цеха";
                ws.Cells[1, 3] = "Название";
                ws.Cells[1, 4] = "Стоимость(шт.)";
                ws.Cells[1, 5] = "Дата поставки";
                ws.Cells[1, 6] = "Стеллаж";
                ws.Cells[1, 7] = "Количество(шт.)";

                for (int i = 0; i < dataGridView7.Rows.Count; i++)
                    for (int j = 0; j < dataGridView7.ColumnCount; j++)
                        application.Cells[i + 2, j + 1] = dataGridView7.Rows[i].Cells[j].Value;

                ws.Columns.AutoFit();

                application.Visible = true;
                application.UserControl = true;
                try
                {
                    string fileName = "e:\\study\\(1) Учёба\\БарГУ Инж.Фак\\"
                                    + "Курсовое проектирование\\CP_2_2\\CP_2_2\\reports_Excel\\report_"
                                    + Name;
                    wb.SaveAs(fileName);
                }
                catch (Exception) { }
            }
        }
         
    }
}
