using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Globalization;
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
        DateTime pocetak = new DateTime();
        DateTime kraj = new DateTime();
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
                    comboBox7.Items.Add(reader.GetString(0));
                }

                reader.Close();

                query = "SELECT DISTINCT ime FROM radnici";
                command = new MySqlCommand(query, con.cn);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    comboBox5.Items.Add(reader.GetString(0));
                }

                reader.Close();
                DateTime d = new DateTime();
                d = DateTime.Now;

                monthCalendar2.SetDate(d);
                monthCalendar1.SetDate(d);
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
                string query = "SELECT mjesto, dostupno FROM skladiste WHERE naziv_alata = '" + comboBox4.SelectedItem.ToString() + "' AND vrsta = '" + comboBox1.SelectedItem.ToString() + "' AND mjesto = '" + textBox12.Text + "'";
                command = new MySqlCommand(query, con.cn);
                // Postavljanje podataka u varijable iz textboxova
                string vrsta = comboBox1.SelectedItem.ToString();
                string naziv_alata = comboBox4.SelectedItem.ToString();
                string ime_prezime = comboBox5.SelectedItem.ToString();
                string kolicina = textBox3.Text;
                string mjesto = textBox12.Text;
                int max_kolicina = 0;
                DateTime time = DateTime.Now;
                MySqlDataReader sdr = command.ExecuteReader();
                while (sdr.Read())
                {
                    max_kolicina = int.Parse(sdr["dostupno"].ToString());
                }
                sdr.Close();
                if (int.Parse(kolicina) > max_kolicina)
                {
                    comboBox1.SelectedItem = null;
                    comboBox4.SelectedItem = null;
                    comboBox5.SelectedItem = null;
                    textBox3.Text = "";
                    MessageBox.Show("Nema toliko alata na stanju");
                    con.cn.Close();
                    return;
                }
                // Postavljanje zaduzenja u bazu
                for (int i = 0; i < Int32.Parse(kolicina); i++)
                {
                    query = "INSERT INTO historija (naziv_alata, vrsta, ime_prezime, kolicina, mjesto, datum_zaduzenja) VALUES ('" + naziv_alata + "', '" + vrsta + "', '" + ime_prezime + "', '" + 1 + "', '" + mjesto + "', '" + time.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                    command = new MySqlCommand(query, con.cn);
                    command.ExecuteNonQuery();
                }

                query = "UPDATE skladiste SET dostupno = dostupno - '" + kolicina + "' WHERE naziv_alata = '" + naziv_alata + "' AND vrsta = '" + vrsta + "' AND mjesto = '" + mjesto + "'";
                command = new MySqlCommand(query, con.cn);
                command.ExecuteNonQuery();
                con.cn.Close();
                if (!comboBox3.Items.Contains(ime_prezime))
                    comboBox3.Items.Add(ime_prezime);
                if (!comboBox7.Items.Contains(ime_prezime))
                    comboBox7.Items.Add(ime_prezime);
                // Vracanje sve na pocetno
                comboBox4.SelectedItem = null;
                comboBox1.SelectedItem = null;
                comboBox5.SelectedItem = null;
                textBox3.Text = textBox12.Text = "";
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
                comboBox4.Items.Clear();
                // Ucitavanje naziva alata u textbox
                string query = "SELECT naziv_alata FROM skladiste WHERE vrsta = '" + comboBox1.SelectedItem + "'";
                command = new MySqlCommand(query, con.cn);
                MySqlDataReader sdr = command.ExecuteReader();
                // Inicijaliziranje autocomplete textboxa
              
                while (sdr.Read())
                {
                    // Dodavanje naziva alata u autocomplete
                    
                    comboBox4.Items.Add(sdr.GetString(0));
                }
                sdr.Close();
                
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
            {
                if (checkBox19.Checked)
                {
                    query = "SELECT * FROM historija";
                }
                else
                {
                    query = "SELECT * FROM historija WHERE datum_zaduzenja  >= CAST('" + pocetak.ToString("yyyy-MM-dd") + "' AS DATETIME) AND  datum_zaduzenja < CAST('" + kraj.AddDays(1).ToString("yyyy-MM-dd") + "' AS DATETIME); ";
                }
            }
            else
            {
                if (checkBox19.Checked)
                {
                    query = "SELECT * FROM historija WHERE ime_prezime = '" + comboBox3.SelectedItem.ToString() + "'";
                }
                else
                {
                    query = "SELECT * FROM historija WHERE ime_prezime = '" + comboBox3.SelectedItem.ToString() + "' AND datum_zaduzenja  >= CAST('" + pocetak.ToString("yyyy-MM-dd") + "' AS DATETIME) AND  datum_zaduzenja < CAST('" + kraj.AddDays(1).ToString("yyyy-MM-dd") + "' AS DATETIME); ";
                }
            }
                
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

            // Ucitavanje podataka iz baze o alatima, sorted first by 'mjesto' and then by 'vrsta' in alphabetical order
            string query = "SELECT * FROM skladiste ORDER BY mjesto, vrsta";
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
                    double totalSum = 0;
                    // Postavljanje podataka u tabele
                    for (int row = 0; row < dataGridView1.Rows.Count; row++)
                    {
                        for (int col = 0; col < dataGridView1.Columns.Count; col++)
                        {
                            var cellValue = dataGridView1.Rows[row].Cells[col].Value;
                            worksheet.Cells[row + 2, col + 1].Value = cellValue;

                            if (col == 3 && cellValue != null) // Check for null
                            {
                                string dValueString = cellValue.ToString();
                                string fValueString = dataGridView1.Rows[row].Cells[5].Value?.ToString(); // Using ?. for safe navigation

                                if (double.TryParse(dValueString, NumberStyles.Any, CultureInfo.InvariantCulture, out double dValue) &&
                                    double.TryParse(fValueString, NumberStyles.Any, CultureInfo.InvariantCulture, out double fValue))
                                {
                                    totalSum += dValue * fValue;
                                }
                            }
                        }
                    }
                    totalSum = Math.Round(totalSum, 2);

                    // Adding the total sum at the end, below all cells
                    int lastRow = dataGridView1.Rows.Count + 2;
                    worksheet.Cells[lastRow + 1, 1].Value = "Ukupna cijena: ";
                    worksheet.Cells[lastRow + 1, 2].Value = totalSum;
                    worksheet.Cells[lastRow + 1, 1].Style.Font.Bold = true;

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
                string query = "SELECT kolicina, dostupno FROM skladiste WHERE naziv_alata = '" + comboBox4.SelectedItem.ToString() + "' AND vrsta = '" + comboBox1.SelectedItem.ToString() + "' AND mjesto = '" + textBox12.Text + "'";
                command = new MySqlCommand(query, con.cn);
                MySqlDataReader sdr = command.ExecuteReader();
                int max_kolicina = 0;
                while (sdr.Read())
                {
                    max_kolicina = int.Parse(sdr["kolicina"].ToString()) - int.Parse(sdr["dostupno"].ToString());
                }
                sdr.Close();

                query = "SELECT kolicina FROM historija WHERE naziv_alata = '" + comboBox4.SelectedItem.ToString() + "'  AND vrsta = '" + comboBox1.SelectedItem.ToString() + "' AND mjesto = '" + textBox12.Text + "'";
                command = new MySqlCommand(query, con.cn);
                // Postavljanje podataka u varijable iz textboxova
                string vrsta = comboBox1.SelectedItem.ToString();
                string naziv_alata = comboBox4.SelectedItem.ToString();
                string ime_prezime = comboBox5.SelectedItem.ToString();
                string kolicina = textBox3.Text;
                string mjesto = textBox12.Text;
                DateTime time = DateTime.Now;
                
                if (int.Parse(kolicina) > max_kolicina)
                {

                    comboBox1.SelectedItem = null;
                    comboBox4.SelectedItem = null;
                    comboBox5.SelectedItem = null;
                    textBox3.Text = "";
                    MessageBox.Show("Pokušavate vratiti više nego što je zaduženo");
                    con.cn.Close();
                    return;
                }
                // Postavljanje zaduzenja u bazu
                query = @"UPDATE historija 
                  JOIN (SELECT id FROM historija WHERE naziv_alata = '" + naziv_alata + @"' AND mjesto = '" + mjesto + @"' AND ime_prezime = '" + ime_prezime + @"' AND datum_vracanja IS NULL LIMIT " + kolicina + @") t 
                  ON historija.id = t.id
                  SET datum_vracanja = '" + time.ToString("yyyy-MM-dd HH:mm:ss") + @"';";
                command = new MySqlCommand(query, con.cn);
                command.ExecuteNonQuery();
                query = "UPDATE skladiste SET dostupno = dostupno + '" + kolicina + "' WHERE naziv_alata = '" + naziv_alata + "' AND vrsta = '" + vrsta + "' AND mjesto = '" + mjesto + "'";
                command = new MySqlCommand(query, con.cn);
                command.ExecuteNonQuery();
                con.cn.Close();
                // Vracanje sve na pocetno
                comboBox1.SelectedItem = null;
                comboBox4.SelectedItem = null;
                comboBox5.SelectedItem = null;
                textBox3.Text = textBox12.Text = "";
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
                comboBox6.Items.Clear();
                con.cn.Open();
                // Ucitavanje naziva alata u textbox
                string query = "SELECT naziv_alata FROM skladiste WHERE vrsta = '" + comboBox2.SelectedItem + "'";
                command = new MySqlCommand(query, con.cn);
                MySqlDataReader sdr = command.ExecuteReader();
                // Inicijaliziranje autocomplete textboxa

                while (sdr.Read())
                {
                    // Dodavanje naziva alata u autocomplete
                    comboBox6.Items.Add(sdr.GetString(0));
                }
                sdr.Close();

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
                string query = "SELECT kolicina, cijena FROM skladiste WHERE naziv_alata = '" + comboBox6.SelectedItem.ToString() + "' AND vrsta = '" + comboBox2.SelectedItem + "' AND mjesto = '" + textBox14.Text + "'";
                command = new MySqlCommand(query, con.cn);
                // Postavljanje podataka u varijable iz textboxova
                string vrsta = comboBox2.SelectedItem.ToString();
                string naziv_alata = comboBox6.SelectedItem.ToString();
                string radnik = comboBox7.SelectedItem.ToString();
                string kolicina_polomljenog = textBox5.Text;
                string mjesto = textBox14.Text;
                int max_kolicina = 0;
                double cijena = 0;
                DateTime time = DateTime.Now;
                MySqlDataReader sdr = command.ExecuteReader();
                while (sdr.Read())
                {
                    max_kolicina = int.Parse(sdr["kolicina"].ToString());
                    cijena = double.Parse(sdr["cijena"].ToString(), CultureInfo.InvariantCulture);
                }
                sdr.Close();
                if (int.Parse(kolicina_polomljenog) > max_kolicina)
                {
                    comboBox2.SelectedItem = comboBox6.SelectedItem = comboBox7.SelectedItem = null;
                    textBox5.Text = textBox14.Text = "";
                    MessageBox.Show("Nema toliko alata na stanju");
                    con.cn.Close();
                    return;
                }
                query = "SELECT COUNT(*) FROM historija WHERE naziv_alata = '" + comboBox6.SelectedItem.ToString() + "' AND vrsta = '" + comboBox2.SelectedItem + "' AND mjesto = '" + mjesto + "' AND ime_prezime = '" + comboBox7.SelectedItem.ToString() + "' AND datum_vracanja IS NULL";
                command = new MySqlCommand(query, con.cn);
                int kolicina_zaduzenog = Convert.ToInt32(command.ExecuteScalar());
                if (kolicina_zaduzenog > max_kolicina || int.Parse(kolicina_polomljenog) > kolicina_zaduzenog)
                {
                    comboBox2.SelectedItem = comboBox6.SelectedItem = comboBox7.SelectedItem = null;
                    textBox5.Text = textBox14.Text = "";
                    MessageBox.Show("Nema toliko alata na stanju");
                    con.cn.Close();
                    return;
                }
                for (int i = 0; i < Int32.Parse(kolicina_polomljenog); i++)
                {
                    query = "UPDATE historija SET datum_vracanja = '" + time.ToString("yyyy-MM-dd HH:mm:ss") + "'  WHERE naziv_alata = '" + comboBox6.SelectedItem.ToString() + "' AND mjesto = '" + mjesto + "' AND vrsta = '" + comboBox2.SelectedItem + "' AND ime_prezime = '" + comboBox7.SelectedItem.ToString() + "' AND datum_vracanja IS NULL LIMIT 1";
                    command = new MySqlCommand(query, con.cn);
                    command.ExecuteNonQuery();
                }

                // Postavljanje zaduzenja u bazu
                query = "UPDATE skladiste SET kolicina = kolicina - '" + kolicina_polomljenog + "' WHERE naziv_alata = '" + naziv_alata + "' AND vrsta = '" + vrsta + "' AND mjesto = '" + mjesto + "'";
                command = new MySqlCommand(query, con.cn);
                command.ExecuteNonQuery();
                query = "SELECT kolicina FROM skladiste WHERE naziv_alata = '" + naziv_alata + "' AND vrsta = '" + vrsta + "' AND mjesto = '" + mjesto + "'";
                command = new MySqlCommand(query, con.cn);
                sdr = command.ExecuteReader();

                if (sdr.Read() && int.Parse(sdr["kolicina"].ToString()) == 0)
                {
                    sdr.Close();
                    query = "DELETE FROM skladiste WHERE naziv_alata = '" + naziv_alata + "' AND vrsta = '" + vrsta + "' AND mjesto = '" + mjesto + "'";
                    command = new MySqlCommand(query, con.cn);
                    command.ExecuteNonQuery();
                }
                sdr.Close();
                int int_kolicina_p = int.Parse(kolicina_polomljenog);

                query = "INSERT INTO polomljeni (naziv_alata, vrsta, radnik, kolicina, mjesto, datum, cijena) VALUES ('" + naziv_alata + "', '" + vrsta + "', '" + radnik + "', '" + kolicina_polomljenog + "', '" + mjesto + "', '" + time.ToString("yyyy-MM-dd HH:mm:ss") + "', '" + (int_kolicina_p * cijena).ToString(CultureInfo.InvariantCulture) + "')";
                command = new MySqlCommand(query, con.cn);
                command.ExecuteNonQuery();
                con.cn.Close();
                // Vracanje sve na pocetno
                comboBox2.SelectedItem = comboBox6.SelectedItem = comboBox7.SelectedItem = null;
                textBox5.Text = textBox14.Text = "";
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
                string cijena = textBox8.Text;
                string mjesto = textBox11.Text;

                if (!comboBox1.Items.Contains(vrsta))
                    comboBox1.Items.Add(vrsta);
                if (!comboBox2.Items.Contains(vrsta))
                    comboBox2.Items.Add(vrsta);

                string checkQuery = "SELECT COUNT(*) FROM skladiste WHERE naziv_alata = '" + naziv_alata + "' AND vrsta = '" + vrsta + "' AND mjesto = '" + mjesto + "'";
                MySqlCommand checkCommand = new MySqlCommand(checkQuery, con.cn);
                int count = Convert.ToInt32(checkCommand.ExecuteScalar());

                if (count > 0)
                {
                    // If the tool exists, update the quantity
                    string updateQuery = "UPDATE skladiste SET kolicina = kolicina + '" + kolicina + "', dostupno = dostupno + '" + kolicina + "' WHERE naziv_alata = '" + naziv_alata + "' AND vrsta = '" + vrsta + "' AND mjesto = '" + mjesto + "'";
                    MySqlCommand updateCommand = new MySqlCommand(updateQuery, con.cn);
                    updateCommand.ExecuteNonQuery();
                }
                else
                {
                    // If the tool does not exist, insert it
                    string insertQuery = "INSERT INTO skladiste (naziv_alata, vrsta, cijena, mjesto, kolicina, dostupno) VALUES ('" + naziv_alata + "', '" + vrsta + "','" + cijena + "', '" + mjesto + "', '" + kolicina + "', '" + kolicina + "')";
                    MySqlCommand insertCommand = new MySqlCommand(insertQuery, con.cn);
                    insertCommand.ExecuteNonQuery();
                }

                con.cn.Close();

                comboBox1.SelectedItem = null;
                textBox13.Text = textBox9.Text = textBox15.Text = textBox8.Text = textBox11.Text = "";

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

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                con.cn.Open();

                string ime = textBox1.Text;

                if (!comboBox5.Items.Contains(ime))
                    comboBox5.Items.Add(ime);


                string insertQuery = "INSERT INTO radnici (ime) VALUES ('" + ime + "')";
                MySqlCommand insertCommand = new MySqlCommand(insertQuery, con.cn);
                insertCommand.ExecuteNonQuery();

                con.cn.Close();

                textBox1.Text = "";

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

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {
            pocetak = monthCalendar1.SelectionEnd;
        }

        private void monthCalendar2_DateChanged(object sender, DateRangeEventArgs e)
        {
            kraj = monthCalendar2.SelectionEnd;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            con.cn.Open();
            // Ucitavanje podataka iz baze o zaduzenjima
            string query;
            if (comboBox3.SelectedItem.ToString() == "Svi")
            {
                if (checkBox19.Checked)
                {
                    query = "SELECT * FROM polomljeni";
                }
                else
                {
                    query = "SELECT * FROM polomljeni WHERE datum  >= CAST('" + pocetak.ToString("yyyy-MM-dd") + "' AS DATETIME) AND  datum < CAST('" + kraj.AddDays(1).ToString("yyyy-MM-dd") + "' AS DATETIME); ";
                }
            }
            else
            {
                if (checkBox19.Checked)
                {
                    query = "SELECT * FROM polomljeni WHERE radnik = '" + comboBox3.SelectedItem.ToString() + "'";
                }
                else
                {
                    query = "SELECT * FROM polomljeni WHERE radnik = '" + comboBox3.SelectedItem.ToString() + "' AND datum  >= CAST('" + pocetak.ToString("yyyy-MM-dd") + "' AS DATETIME) AND  datum < CAST('" + kraj.AddDays(1).ToString("yyyy-MM-dd") + "' AS DATETIME); ";
                }
            }

            command = new MySqlCommand(query, con.cn);
            command.ExecuteNonQuery();
            dataTable = new System.Data.DataTable();
            adapter = new MySqlDataAdapter(command);
            adapter.Fill(dataTable);
            dataGridView1.DataSource = dataTable.DefaultView;
            con.cn.Close();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Polomljeni Alat");

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
                    string fileName = "IzvjestajPolomljenihAlata.xlsx";

                    // Konstrukcija
                    string filePath = Path.Combine(documentsPath, fileName);

                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath); // Brisanje vec postojeceg fajla
                    }

                    // Spasavanje
                    FileInfo file = new FileInfo(filePath);
                    package.SaveAs(file);

                    MessageBox.Show("Izvjestaj je kreiran, mozete ga naci u Arhiva folderu pod nazivom IzvjestajPolomljenihAlata.xlsx ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}