using MySql.Data.MySqlClient;
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

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                // Ucitavanje vrsta alata u combobox
                con.cn.Open();

                string query = "SELECT DISTINCT vrsta FROM skladiste";
                MySqlCommand command = new MySqlCommand(query, con.cn);
                MySqlDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    comboBox1.Items.Add(reader.GetString(0));
                    comboBox2.Items.Add(reader.GetString(0));
                }

                reader.Close();

                comboBox3.Items.Add("Svi");
                comboBox3.SelectedItem = "Svi";
                query = "SELECT DISTINCT ime_prezime FROM historija";
                command = new MySqlCommand(query, con.cn);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    comboBox3.Items.Add(reader.GetString(0));
                }
                reader.Close();

                con.cn.Close();
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
                // Uzimanje podataka iz baze o alatu
                string query = "SELECT sifra_kupovine, dostupno FROM skladiste WHERE naziv_alata = '" + textBox1.Text + "' AND vrsta = '" + comboBox1.SelectedItem.ToString() + "'";
                command = new MySqlCommand(query, con.cn);
                // Postavljanje podataka u varijable iz textboxova
                string vrsta = comboBox1.SelectedItem.ToString();
                string naziv_alata = textBox1.Text;
                string ime_prezime = textBox2.Text;
                string kolicina = textBox3.Text;
                string sifra = textBox12.Text;
                string sifra_kupovine = "";
                int max_kolicina = 0;
                DateTime time = DateTime.Now;
                MySqlDataReader sdr = command.ExecuteReader();
                while (sdr.Read())
                {
                    sifra_kupovine = sdr["sifra_kupovine"].ToString();
                    max_kolicina = int.Parse(sdr["dostupno"].ToString());
                }
                sdr.Close();
                if (int.Parse(kolicina) > max_kolicina)
                {
                    comboBox1.SelectedItem = null;
                    textBox1.Text = textBox2.Text = textBox3.Text = "";
                    MessageBox.Show("Nema toliko alata na stanju");
                    con.cn.Close();
                    return;
                }
                // Postavljanje zaduzenja u bazu
                for (int i = 0; i < Int32.Parse(kolicina); i++)
                {
                    query = "INSERT INTO historija (naziv_alata, vrsta, ime_prezime, kolicina, sifra_kupovine, sifra, datum_zaduzenja) VALUES ('" + naziv_alata + "', '" + vrsta + "', '" + ime_prezime + "', '" + 1 + "', '" + sifra_kupovine + "', '" + sifra + "', '" + time.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                    command = new MySqlCommand(query, con.cn);
                    command.ExecuteNonQuery();
                }

                query = "UPDATE skladiste SET dostupno = dostupno - '" + kolicina + "' WHERE naziv_alata = '" + naziv_alata + "'";
                command = new MySqlCommand(query, con.cn);
                command.ExecuteNonQuery();
                con.cn.Close();
                if (!comboBox3.Items.Contains(ime_prezime))
                    comboBox3.Items.Add(ime_prezime);
                // Vracanje sve na pocetno
                comboBox1.SelectedItem = null;
                textBox1.Text = textBox2.Text = textBox3.Text = textBox12.Text = "";
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
                // Ucitavanje naziva alata u textbox
                string query = "SELECT naziv_alata FROM skladiste WHERE vrsta = '" + comboBox1.SelectedItem + "'";
                command = new MySqlCommand(query, con.cn);
                MySqlDataReader sdr = command.ExecuteReader();
                // Inicijaliziranje autocomplete textboxa
                AutoCompleteStringCollection autoText = new AutoCompleteStringCollection();
                while (sdr.Read())
                {
                    // Dodavanje naziva alata u autocomplete
                    autoText.Add(sdr.GetString(0));
                }
                sdr.Close();
                // Postavljanje autocompletea na textbox
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
            // Ucitavanje podataka iz baze o zaduzenjima
            string query;
            if (comboBox3.SelectedItem.ToString() == "Svi")
                query = "SELECT * FROM historija";
            else
                query = "SELECT * FROM historija WHERE ime_prezime = '" + comboBox3.SelectedItem.ToString() + "'";
            command = new MySqlCommand(query, con.cn);
            command.ExecuteNonQuery();
            dataTable = new System.Data.DataTable();
            adapter = new MySqlDataAdapter(command);
            adapter.Fill(dataTable);
            dataGridView1.DataSource = dataTable.DefaultView;
            con.cn.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            con.cn.Open();
            // Ucitavanje podataka iz baze o alatima
            string query = "SELECT * FROM skladiste";
            command = new MySqlCommand(query, con.cn);
            command.ExecuteNonQuery();
            dataTable = new System.Data.DataTable();
            adapter = new MySqlDataAdapter(command);
            adapter.Fill(dataTable);
            dataGridView1.DataSource = dataTable.DefaultView;
            con.cn.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Historija");

                    // Postavljanje zaglavlja tabele
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = dataGridView1.Columns[i].HeaderText;
                    }

                    // Postavljanje podataka u tabele
                    for (int row = 0; row < dataGridView1.Rows.Count; row++)
                    {
                        for (int col = 0; col < dataGridView1.Columns.Count; col++)
                        {
                            var cellValue = dataGridView1.Rows[row].Cells[col].Value;
                            if (cellValue is DateTime)
                            {
                                // if cellValue is a DateTime
                                worksheet.Cells[row + 2, col + 1].Style.Numberformat.Format = "yyyy-mm-dd HH:mm:ss";
                            }
                            worksheet.Cells[row + 2, col + 1].Value = cellValue;
                        }
                    }

                    // Auto-fit kolone
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    // Postavljanje putanje za spremanje
                    string documentsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Arhiva");

                    // Postavljanje imena fajla
                    string fileName = "IzvjestajHistorije.xlsx";

                    // Konstrukcija
                    string filePath = Path.Combine(documentsPath, fileName);

                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath); // Brisanje vec postojeceg fajla
                    }

                    // Spasavanje
                    FileInfo file = new FileInfo(filePath);
                    package.SaveAs(file);

                    MessageBox.Show("Izvjestaj je kreiran, mozete ga naci u Arhiva folderu pod nazivom IzvjestajHistorije.xlsx ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Skladiste");

                    // Postavljanje zaglavlja tabele
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = dataGridView1.Columns[i].HeaderText;
                    }

                    // Postavljanje podataka u tabele
                    for (int row = 0; row < dataGridView1.Rows.Count; row++)
                    {
                        for (int col = 0; col < dataGridView1.Columns.Count; col++)
                        {
                            worksheet.Cells[row + 2, col + 1].Value = dataGridView1.Rows[row].Cells[col].Value;
                        }
                    }

                    // Auto-fit kolone
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    // Postavljanje putanje za spremanje
                    string documentsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Arhiva");

                    // Postavljanje imena fajla
                    string fileName = "IzvjestajSkladiste.xlsx";

                    // Konstrukcija
                    string filePath = Path.Combine(documentsPath, fileName);

                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath); // Brisanje vec postojeceg fajla
                    }

                    // Spasavanje
                    FileInfo file = new FileInfo(filePath);
                    package.SaveAs(file);

                    MessageBox.Show("Izvjestaj je kreiran, mozete ga naci u Arhiva folderu pod nazivom IzvjestajSkladista.xlsx");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                con.cn.Open();
                // Uzimanje podataka iz baze o alatu
                string query = "SELECT kolicina, dostupno FROM skladiste WHERE naziv_alata = '" + textBox1.Text + "' AND vrsta = '" + comboBox1.SelectedItem.ToString() + "'";
                command = new MySqlCommand(query, con.cn);
                MySqlDataReader sdr = command.ExecuteReader();
                int max_kolicina = 0;
                while (sdr.Read())
                {
                    max_kolicina = int.Parse(sdr["kolicina"].ToString()) - int.Parse(sdr["dostupno"].ToString());
                }
                sdr.Close();
                MessageBox.Show(max_kolicina.ToString());
                query = "SELECT sifra_kupovine, kolicina FROM historija WHERE naziv_alata = '" + textBox1.Text + "'  AND vrsta = '" + comboBox1.SelectedItem.ToString() + "'";
                command = new MySqlCommand(query, con.cn);
                // Postavljanje podataka u varijable iz textboxova
                string vrsta = comboBox1.SelectedItem.ToString();
                string naziv_alata = textBox1.Text;
                string ime_prezime = textBox2.Text;
                string kolicina = textBox3.Text;
                string sifra = textBox12.Text;
                string sifra_kupovine = "";

                DateTime time = DateTime.Now;
                sdr = command.ExecuteReader();
                while (sdr.Read())
                {
                    sifra_kupovine = sdr["sifra_kupovine"].ToString();
                }
                sdr.Close();
                if (int.Parse(kolicina) > max_kolicina)
                {
                    MessageBox.Show(max_kolicina.ToString() + " " + kolicina);
                    comboBox1.SelectedItem = null;
                    textBox1.Text = textBox2.Text = textBox3.Text = "";
                    MessageBox.Show("Pokušavate vratiti više nego što je zaduženo");
                    con.cn.Close();
                    return;
                }
                // Postavljanje zaduzenja u bazu
                query = @"UPDATE historija 
                  JOIN (SELECT id FROM historija WHERE naziv_alata = '" + naziv_alata + @"' AND ime_prezime = '" + ime_prezime + @"' AND datum_vracanja IS NULL LIMIT " + kolicina + @") t 
                  ON historija.id = t.id
                  SET datum_vracanja = '" + time.ToString("yyyy-MM-dd HH:mm:ss") + @"';";
                command = new MySqlCommand(query, con.cn);
                command.ExecuteNonQuery();
                query = "UPDATE skladiste SET dostupno = dostupno + '" + kolicina + "' WHERE naziv_alata = '" + naziv_alata + "' AND vrsta = '" + vrsta + "' AND sifra_kupovine = '" + sifra_kupovine + "' AND sifra = '" + sifra + "'";
                command = new MySqlCommand(query, con.cn);
                command.ExecuteNonQuery();
                con.cn.Close();
                // Vracanje sve na pocetno
                comboBox1.SelectedItem = null;
                textBox1.Text = textBox2.Text = textBox3.Text = textBox12.Text = "";
                MessageBox.Show("Razduženo!");
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

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                con.cn.Open();
                // Ucitavanje naziva alata u textbox
                string query = "SELECT naziv_alata FROM skladiste WHERE vrsta = '" + comboBox2.SelectedItem + "'";
                command = new MySqlCommand(query, con.cn);
                MySqlDataReader sdr = command.ExecuteReader();
                // Inicijaliziranje autocomplete textboxa
                AutoCompleteStringCollection autoText = new AutoCompleteStringCollection();
                while (sdr.Read())
                {
                    // Dodavanje naziva alata u autocomplete
                    autoText.Add(sdr.GetString(0));
                }
                sdr.Close();
                // Postavljanje autocompletea na textbox
                textBox4.AutoCompleteMode = AutoCompleteMode.Suggest;
                textBox4.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textBox4.AutoCompleteCustomSource = autoText;
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

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                con.cn.Open();
                // Uzimanje podataka iz baze o alatu
                string query = "SELECT sifra_kupovine, kolicina FROM skladiste WHERE naziv_alata = '" + textBox4.Text + "' AND vrsta = '" + comboBox2.SelectedItem + "'";
                command = new MySqlCommand(query, con.cn);
                // Postavljanje podataka u varijable iz textboxova
                string vrsta = comboBox2.SelectedItem.ToString();
                string naziv_alata = textBox4.Text;
                string kolicina = textBox5.Text;
                string dostupno = textBox6.Text;
                string sifra = textBox14.Text;
                string sifra_kupovine = "";
                int max_kolicina = 0;
                DateTime time = DateTime.Now;
                MySqlDataReader sdr = command.ExecuteReader();
                while (sdr.Read())
                {
                    sifra_kupovine = sdr["sifra_kupovine"].ToString();
                    max_kolicina = int.Parse(sdr["kolicina"].ToString());
                }
                sdr.Close();
                if (int.Parse(kolicina) > max_kolicina)
                {
                    comboBox2.SelectedItem = null;
                    textBox4.Text = textBox5.Text = textBox6.Text = "";
                    MessageBox.Show("Nema toliko alata na stanju");
                    con.cn.Close();
                    return;
                }
                // Postavljanje zaduzenja u bazu
                query = "UPDATE skladiste SET kolicina = kolicina - '" + kolicina + "', dostupno = '" + dostupno + "' WHERE naziv_alata = '" + naziv_alata + "' AND vrsta = '" + vrsta + "' AND sifra_kupovine = '" + sifra_kupovine + "'  AND sifra = '" + sifra + "'";
                command = new MySqlCommand(query, con.cn);
                command.ExecuteNonQuery();
                query = "SELECT kolicina FROM skladiste WHERE naziv_alata = '" + naziv_alata + "' AND vrsta = '" + vrsta + "' AND sifra_kupovine = '" + sifra_kupovine + "'  AND sifra = '" + sifra + "'";
                command = new MySqlCommand(query, con.cn);
                sdr = command.ExecuteReader();

                if (sdr.Read() && int.Parse(sdr["kolicina"].ToString()) == 0)
                {
                    sdr.Close();
                    query = "DELETE FROM skladiste WHERE naziv_alata = '" + naziv_alata + "' AND vrsta = '" + vrsta + "' AND sifra_kupovine = '" + sifra_kupovine + "'  AND sifra = '" + sifra + "'";
                    command = new MySqlCommand(query, con.cn);
                    command.ExecuteNonQuery();
                }
                con.cn.Close();
                // Vracanje sve na pocetno
                comboBox2.SelectedItem = null;
                textBox4.Text = textBox5.Text = textBox6.Text = textBox14.Text = "";
                MessageBox.Show("Izbrisano!");
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

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                con.cn.Open();

                string vrsta = textBox13.Text;
                string naziv_alata = textBox9.Text;
                string kolicina = textBox15.Text;
                string sifra_kupovine = textBox10.Text;
                string cijena = textBox8.Text;
                string mjesto = textBox11.Text;
                string sifra = textBox7.Text;

                if (!comboBox1.Items.Contains(vrsta))
                    comboBox1.Items.Add(vrsta);
                if (!comboBox2.Items.Contains(vrsta))
                    comboBox2.Items.Add(vrsta);

                string checkQuery = "SELECT COUNT(*) FROM skladiste WHERE naziv_alata = '" + naziv_alata + "' AND vrsta = '" + vrsta + "' AND sifra_kupovine = '" + sifra_kupovine + "' AND sifra = '" + sifra + "'";
                MySqlCommand checkCommand = new MySqlCommand(checkQuery, con.cn);
                int count = Convert.ToInt32(checkCommand.ExecuteScalar());

                if (count > 0)
                {
                    // If the tool exists, update the quantity
                    string updateQuery = "UPDATE skladiste SET kolicina = kolicina + '" + kolicina + "', dostupno = dostupno + '" + kolicina + "' WHERE naziv_alata = '" + naziv_alata + "' AND vrsta = '" + vrsta + "' AND sifra_kupovine = '" + sifra_kupovine + "' AND sifra = '" + sifra + "'";
                    MySqlCommand updateCommand = new MySqlCommand(updateQuery, con.cn);
                    updateCommand.ExecuteNonQuery();
                }
                else
                {
                    // If the tool does not exist, insert it
                    string insertQuery = "INSERT INTO skladiste (naziv_alata, vrsta, sifra_kupovine, sifra, cijena, mjesto, kolicina, dostupno) VALUES ('" + naziv_alata + "', '" + vrsta + "', '" + sifra_kupovine + "', '" + sifra + "', '" + cijena + "', '" + mjesto + "', '" + kolicina + "', '" + kolicina + "')";
                    MySqlCommand insertCommand = new MySqlCommand(insertQuery, con.cn);
                    insertCommand.ExecuteNonQuery();
                }

                con.cn.Close();

                comboBox1.SelectedItem = null;
                textBox7.Text = textBox13.Text = textBox9.Text = textBox15.Text = textBox10.Text = textBox8.Text = textBox11.Text = "";

                MessageBox.Show("Dodano!");
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



        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}