using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace LauncherforTaksiDeliveryAPP
{
    public partial class Form1 : Form
    {
        private SqlConnection Connections = null;
        private object? _data;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Connections = new SqlConnection(ConfigurationManager.ConnectionStrings["TestDB"].ConnectionString);

            Connections.Open();

            if (Connections.State == ConnectionState.Open)
            {
                MessageBox.Show("Все норм");
            }
            DataGridView2Work();

            DataGridView3Work();
        }
        private void DataGridView2Work()//load in DataGrid DataBase
        {
            SqlDataAdapter adapters = new SqlDataAdapter("SELECT * From Students", Connections);

            DataSet DB = new DataSet();

            adapters.Fill(DB);

            dataGridView2.DataSource = DB.Tables[0];
        }
        private void DataGridView3Work()// load in DataGrid DateBase
        {
            SqlDataAdapter adapter = new SqlDataAdapter("SELECT * From Students", Connections);

            DataSet DBS = new DataSet();

            adapter.Fill(DBS);

            dataGridView3.DataSource = DBS.Tables[0];
        }
        private void button1_Click(object sender, EventArgs e)//insert in DataBase new information
        {
            SqlCommand cmd = new SqlCommand("INSERT INTO [Students] (name, Surname, Birthday, mesto_roshdenya, Phone, Email) VALUES (@name, @Surname, @Birthday, @mesto_roshdenya, @Phone, @Email)", Connections);

            DateTime date = DateTime.Parse(textBox3.Text);
            cmd.Parameters.AddWithValue("name", textBox1.Text);
            cmd.Parameters.AddWithValue("Surname", textBox2.Text);
            cmd.Parameters.AddWithValue("Birthday", $"{date.Month}/{date.Day}/{date.Year}");
            cmd.Parameters.AddWithValue("mesto_roshdenya", textBox4.Text);
            cmd.Parameters.AddWithValue("Phone", textBox5.Text);
            cmd.Parameters.AddWithValue("Email", textBox6.Text);


            MessageBox.Show(cmd.ExecuteNonQuery().ToString(), "Все прошло успешно!");
        }

        private void button2_Click(object sender, EventArgs e)//show DateBase
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Students", Connections);
            DataSet data = new DataSet();

            dataAdapter.Fill(data);

            dataGridView1.DataSource = data.Tables[0];
        }

        private void textBox8_TextChanged(object sender, EventArgs e) //filtr words
        {
            ((DataTable)dataGridView2.DataSource).DefaultView.RowFilter = $"name LIKE '%{textBox8.Text}%'";
        }
        private void button3_Click(object sender, EventArgs e) 
        {
            SaveTable(dataGridView3);
        }
        private void SaveTable(DataGridView View)
        {
            //dataGridView3.DataSource = null;
            //_data = dataGridView3.DataSource.ToString();
            //string path = Path.Combine(Environment.CurrentDirectory, "Export_NasilGroup", "data.csv");
            //string rows = dataGridView3.DataSource.Select(c => $"{c.}");

            // Excel.Application ExcelApp = new Excel.Application();
            ///* Excel.Workbook WorkBook =*/ ExcelApp.Workbooks.Add();
            // Excel.Worksheet WorkSheet = (Excel.Worksheet)ExcelApp.ActiveSheet;   
            // for (int i = 0; i <= View.RowCount - 1; i++)
            // {
            //     for (int j = 0; j <= View.ColumnCount - 1; j++)
            //     {
            //         //WorkSheet.Rows[i].Columns[j] = View.Rows[i - 1].Cells[j - 1].Value;
            //         WorkSheet.Cells[i + 1, j + 1] = View[j, i].Value.ToString();
            //     }
            // }
            // //ExcelApp.AlertBeforeOverwriting = false;
            // //WorkBook.SaveAs(path);
            // //ExcelApp.Quit();
            // ExcelApp.Visible = true;
            MessageBox.Show("Sorry, i can't do that now/ Try other items");
        }
    }
}