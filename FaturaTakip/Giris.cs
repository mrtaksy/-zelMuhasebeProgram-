using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace FaturaTakip
{
    public partial class Giris : Form
    {
        public Giris()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                String sorgu = "Select * from TblKullanicilar where K_Adi='" + textBox1.Text + "' and K_Sifre='" + textBox2.Text + "'";
                OleDbCommand command = new OleDbCommand(sorgu, con);
                con.Open();
                OleDbDataReader reader = command.ExecuteReader();

                if (reader.Read())
                {
                    this.Hide();
                    Form1 f1 = new Form1();
                    f1.Show();
                }
                else
                    MessageBox.Show("Giriş Başarısız Bilgileri Kontrol Ediniz!!");
                con.Close();
            }
            else
                MessageBox.Show("Lütfen Tüm Bilgileri Giriniz!");
          
        }
        OleDbConnection con;
        OleDbDataAdapter da;
        OleDbCommand cmd;
        DataSet ds;
        private void Giris_Load(object sender, EventArgs e)
        {
            con = new OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = db_fatura.accdb");
        }
    }
}
