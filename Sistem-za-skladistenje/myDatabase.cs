using MySql.Data.MySqlClient;

namespace Sistem_za_skladistenje
{
    class myDatabase
    {
        public MySqlConnection cn;
        public void Connect()
        {
            string ip = "192.168.0.100";
            cn = new MySqlConnection("Datasource = " + ip + ";username=Remote;password=admin; database=skladiste;Convert Zero Datetime=True");
        }
    }
}
