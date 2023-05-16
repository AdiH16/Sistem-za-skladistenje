using MySql.Data.MySqlClient;
using System;
using System.Windows.Forms;
using OfficeOpenXml;
using System;
using System.IO;
using System.Windows.Forms;

namespace Sistem_za_skladistenje
{
    public partial class Form1 : Form
    {
        myDatabase con = new myDatabase();
        MySqlCommand command;
        MySqlDataAdapter adapter;
        System.Data.DataTable dataTable;
        public Form1()
        {
            InitializeComponent();
            con.Connect();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                comboBox1.Items.Add("Glodala");
                comboBox1.Items.Add("Burgije");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                con.cn.Open();
                string query = "SELECT interna_sifra, sifra_proizvodnje, kolicina FROM skladiste WHERE naziv_alata = '" + textBox1.Text + "'";
                command = new MySqlCommand(query, con.cn);
                string naziv_alata = textBox1.Text;
                string ime_prezime = textBox2.Text;
                string kolicina = textBox3.Text;
                string sifra_proizvodnje = "", interna_sifra = "";
                int max_kolicina = 0;
                DateTime time = DateTime.Now;
                MySqlDataReader sdr = command.ExecuteReader();
                while (sdr.Read())
                {
                    sifra_proizvodnje = sdr["sifra_proizvodnje"].ToString();
                    interna_sifra = sdr["interna_sifra"].ToString();
                    max_kolicina = int.Parse(sdr["kolicina"].ToString());
                }
                sdr.Close();
                if(int.Parse(kolicina) > max_kolicina)
                {
                    comboBox1.SelectedItem = null;
                    textBox1.Text = textBox2.Text = textBox3.Text = "";
                    MessageBox.Show("Nema toliko alata na stanju");
                    con.cn.Close();
                    return;
                }
                query = "INSERT INTO historija (naziv_alata, ime_prezime, kolicina, sifra_proizvodnje, interna_sifra, datum_zaduzenja) VALUES ('" + naziv_alata + "', '" + ime_prezime + "', '" + kolicina + "', '" + sifra_proizvodnje + "', '" + interna_sifra + "', '"+time.ToString("yyyy-MM-dd HH:mm:ss") +"')";
                command = new MySqlCommand(query, con.cn);
                command.ExecuteNonQuery();
                con.cn.Close();
                comboBox1.SelectedItem = null;
                textBox1.Text = textBox2.Text = textBox3.Text = "";
                MessageBox.Show("Iznajmljeno!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                con.cn.Close();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                con.cn.Open();
                string query = "SELECT naziv_alata FROM skladiste WHERE vrsta = '" + comboBox1.SelectedItem + "'";
                command = new MySqlCommand(query, con.cn);
                MySqlDataReader sdr = command.ExecuteReader();
                AutoCompleteStringCollection autoText = new AutoCompleteStringCollection();
                while (sdr.Read())
                {
                    autoText.Add(sdr.GetString(0));
                }
                sdr.Close();
                textBox1.AutoCompleteMode = AutoCompleteMode.Suggest;
                textBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textBox1.AutoCompleteCustomSource = autoText;
                con.cn.Close();
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            finally
            {
                con.cn.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            con.cn.Open();
            string query = "SELECT * FROM historija";
            command = new MySqlCommand(query,con.cn);
            command.ExecuteNonQuery();
            dataTable = new System.Data.DataTable();
            adapter = new MySqlDataAdapter(command);
            adapter.Fill(dataTable);
            dataGridView1.DataSource = dataTable.DefaultView;
            con.cn.Close();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Historija");

                    // Export column headers
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = dataGridView1.Columns[i].HeaderText;
                    }

                    // Export data
                    for (int row = 0; row < dataGridView1.Rows.Count; row++)
                    {
                        for (int col = 0; col < dataGridView1.Columns.Count; col++)
                        {
                            worksheet.Cells[row + 2, col + 1].Value = dataGridView1.Rows[row].Cells[col].Value;
                        }
                    }

                    // Auto-fit columns
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    // Get the Documents folder path
                    string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                    // Generate a unique file name
                    string fileName = "izvjestaj" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

                    // Construct the file path
                    string filePath = Path.Combine(documentsPath, fileName);

                    // Save the Excel file
                    FileInfo file = new FileInfo(filePath);
                    package.SaveAs(file);

                    MessageBox.Show("Izvjestaj je kreiran, mozete ga naci u Documents folderu pod nazivom Izvjestaj ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            con.cn.Open();
            string query = "SELECT * FROM skladiste";
            command = new MySqlCommand(query, con.cn);
            command.ExecuteNonQuery();
            dataTable = new System.Data.DataTable();
            adapter = new MySqlDataAdapter(command);
            adapter.Fill(dataTable);
            dataGridView1.DataSource = dataTable.DefaultView;
            con.cn.Close();
        }
    }
}
        private void button3_Click(object sender, EventArgs e)
        {
              try
              {
                using (ExcelPackage package = new ExcelPackage())
                {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Skladiste");

            // Export column headers
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = dataGridView1.Columns[i].HeaderText;
            }

            // Export data
            for (int row = 0; row < dataGridView1.Rows.Count; row++)
            {
                for (int col = 0; col < dataGridView1.Columns.Count; col++)
                {
                    worksheet.Cells[row + 2, col + 1].Value = dataGridView1.Rows[row].Cells[col].Value;
                }
            }

            // Auto-fit columns
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

            // Get the Documents folder path
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // Generate a unique file name
            string fileName = "izvjestaj 2" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

            // Construct the file path
            string filePath = Path.Combine(documentsPath, fileName);

            // Save the Excel file
            FileInfo file = new FileInfo(filePath);
            package.SaveAs(file);

            MessageBox.Show("Izvjestaj je kreiran, mozete ga naci u Documents folderu pod nazivom Izvjestaj 2");
        }
    }
    catch (Exception ex)
    {
        MessageBox.Show(ex.Message);
    }
}

