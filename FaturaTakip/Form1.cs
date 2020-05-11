using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualBasic;
namespace FaturaTakip
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        OleDbConnection con;
        OleDbDataAdapter da;
        OleDbCommand cmd;
        DataSet ds;
        int k = 0;
        int i = 0, j = 1;
        int x = 180;
        public void bankaListele()
        {
            cbBanka.Items.Clear();
            comboBox3.Items.Clear();
            OleDbCommand command = new OleDbCommand("Select * from TblBankalar", con);

            if (con.State == ConnectionState.Closed)
                con.Open();
            OleDbDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                cbBanka.Items.Add(reader.GetString(1));
                comboBox3.Items.Add(reader.GetString(1));
            }
            cbBanka.Items.Add("--Yeni Ekle");
            reader.Close();
            con.Close();
        }
        public void Temizle()
        {
            rbVergiDairesi.Text = "";
            kayitKontrol = 0;
            btnEkle.Text = "Ekle";
            tbMiktar.Text = "";
            tbCekSahip.Text = "";
            tbCekNo.Text = "";
            cbBanka.Text = "";
            dpOdemeTarih.Value = DateTime.Now;
            rbAciklama.Text = "";
            tbVergi.Text = "";
        }
        private void Form1_Load(object sender, EventArgs e)
        {

            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=" + tbVeriTabani.Text + ".accdb");
            bankaListele();
        }

        private void cbBanka_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbBanka.SelectedItem.ToString() == "--Yeni Ekle")
            {
                string bankaGirisi = Interaction.InputBox("Bilgi Girişi", "Banka Adı Giriniz.", "", 0, 0);
                if (bankaGirisi != "")
                {
                    cmd = new OleDbCommand();

                    if (con.State == ConnectionState.Closed)
                        con.Open();
                    cmd.Connection = con;
                    cmd.CommandText = "insert into TblBankalar (Banka_Adı) values ('" + bankaGirisi + "')";
                    cmd.ExecuteNonQuery();
                    con.Close();
                    bankaListele();
                    cbBanka.Text = bankaGirisi;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            /*if (dpOdemeTarih.Text.Split(' ')[3] == "Cumartesi") { 
              dpOdemeTarih.Value=  dpOdemeTarih.Value.AddDays(2);
              MessageBox.Show("Tarih Hafta Sonuna Denk Geldiği İçin Pazartesiye Alınmıştır.\nAçıklamaya Otomatik Olarak Eklenmiştir");
              rbAciklama.Text = rbAciklama.Text + "\nNot: Asıl Tarih :" + dpOdemeTarih.Value.AddDays(-2);
            }
            if (dpOdemeTarih.Text.Split(' ')[3] == "Pazar") { 
                dpOdemeTarih.Value = dpOdemeTarih.Value.AddDays(1);
            MessageBox.Show("Tarih Hafta Sonuna Denk Geldiği İçin Pazartesiye Alınmıştır.\nAçıklamaya Otomatik Olarak Eklenmiştir");
            rbAciklama.Text=rbAciklama.Text+"\nNot: Asıl Tarih :"+dpOdemeTarih.Value.AddDays(-1);
            }*/
            if (tbCekNo.Text != "" && cbBanka.Text != "" && tbCekSahip.Text != "" && tbMiktar.Text != "")
            {
                cmd = new OleDbCommand();
                if (con.State == ConnectionState.Closed)
                    con.Open();
                cmd.Connection = con;
                if (kayitKontrol != 1)
                {
                    cmd.CommandText = "insert into TblFaturalar (CekNo,OdemeTarih,BankaAdı,CekSahibi,Miktar,Aciklama,Durum,Vergi,VergiDairesi) values ('" + tbCekNo.Text + "','" + dpOdemeTarih.Value.ToShortDateString() + "','" + cbBanka.Text + "','" + tbCekSahip.Text + "','" + tbMiktar.Text + "','" + rbAciklama.Text + "','Ödenmemiş','" + tbVergi.Text + "','" + rbVergiDairesi.Text + "')";
                }
                else
                {
                    cmd.CommandText = "update TblFaturalar set OdemeTarih='" + dpOdemeTarih.Value.ToShortDateString() + "',BankaAdı='" + cbBanka.Text + "',CekSahibi='" + tbCekSahip.Text + "',Miktar='" + tbMiktar.Text + "',Aciklama='" + rbAciklama.Text + "',Vergi='" + tbVergi.Text + "',VergiDairesi='" + rbVergiDairesi.Text + "' where CekNo='" + tbCekNo.Text + "'";
                }
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Kayıt Alınmıştır!!");
                Temizle();
            }
            else
            {
                MessageBox.Show("Lütfen alanları doldurunuz.");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Temizle();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            String sorgu = "Select * from TblFaturalar";
            OleDbCommand command = new OleDbCommand(sorgu, con);
            if (con.State == ConnectionState.Closed)
                con.Open();
            OleDbDataReader reader = command.ExecuteReader();
            double toplamCek = 0;
            double odenmisCek = 0;
            double vaktiGelmemisCek = 0;
            double odenmeyenCek = 0;
            while (reader.Read())
            {
                string[] parcalar = reader.GetString(2).ToString().Split('.');
                if (Convert.ToInt16(parcalar[1]) == (cbAy.SelectedIndex + 1))
                {
                    TimeSpan fark = Convert.ToDateTime(reader.GetString(2)) - DateTime.Now;
                    if (fark.TotalDays > 0)
                        vaktiGelmemisCek += Convert.ToDouble(reader.GetString(5));
                    if (reader.GetString(7) == "Ödenmiş")
                        odenmisCek += Convert.ToDouble(reader.GetString(5));
                    else
                        odenmeyenCek += Convert.ToDouble(reader.GetString(5));

                    toplamCek += Convert.ToDouble(reader.GetString(5));
                }
            }
            lblAOdenmis.Text = odenmisCek + " TL";
            lblAOdenmemis.Text = odenmeyenCek + " TL";
            lblAToplam.Text = toplamCek + " TL";
            lblAGelmemiş.Text = vaktiGelmemisCek + " TL";
            reader.Close();
            con.Close();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            String sorgu = "Select * from TblFaturalar";
            OleDbCommand command = new OleDbCommand(sorgu, con);

            if (con.State == ConnectionState.Closed)
                con.Open();
            OleDbDataReader reader = command.ExecuteReader();
            double toplamCek = 0;
            double odenmisCek = 0;
            double vaktiGelmemisCek = 0;
            double odenmeyenCek = 0;
            while (reader.Read())
            {
                string[] parcalar = reader.GetString(2).ToString().Split('.');
                parcalar[2] = parcalar[2].Split(' ')[0];
                if (Convert.ToInt16(parcalar[2]) == Convert.ToInt16(comboBox2.Text))
                {
                    TimeSpan fark = Convert.ToDateTime(reader.GetString(2)) - DateTime.Now;
                    if (fark.TotalDays > 0)
                        vaktiGelmemisCek += Convert.ToDouble(reader.GetString(5));
                    if (reader.GetString(7) == "Ödenmiş")
                        odenmisCek += Convert.ToDouble(reader.GetString(5));
                    else
                        odenmeyenCek += Convert.ToDouble(reader.GetString(5));

                    toplamCek += Convert.ToDouble(reader.GetString(5));
                }
            }
            lblYOdenmis.Text = odenmisCek + " TL";
            lblYOdenmemis.Text = odenmeyenCek + " TL";
            lblYToplam.Text = toplamCek + " TL";
            lblYGelmemis.Text = vaktiGelmemisCek + " TL";
            reader.Close();
            con.Close();
        }
        List<KeyValuePair<string, Double>> bankalar = new List<KeyValuePair<string, Double>>();
        private void button3_Click(object sender, EventArgs e)
        {
            bankalar.Clear();
            listBox1.Items.Clear();
            Double toplam = 0;
            if (radioButton1.Checked)
            {
                if (textBox1.Text != "")
                {
                    da = new OleDbDataAdapter("Select *from TblFaturalar where CekNo='" + textBox1.Text + "'", con);
                    ds = new DataSet();

                    if (con.State == ConnectionState.Closed)
                        con.Open();
                    da.Fill(ds, "TblFaturalar");
                    dataGridView1.DataSource = ds.Tables["TblFaturalar"];
                    dataGridView1.Columns[0].Visible = false; //KOLON GİZLEME
                    con.Close();

                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {
                        int bankaKontrol = 0;
                        for (int k = 0; k < bankalar.Count; k++)
                        {
                            if (bankalar[k].Key == dataGridView1.Rows[i].Cells["BankaAdı"].Value.ToString())
                            {
                                bankalar[k] = new KeyValuePair<string, double>(dataGridView1.Rows[i].Cells["BankaAdı"].Value.ToString(), bankalar[k].Value + Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value));
                                bankaKontrol = 1;
                                break;
                            }
                        }
                        if (bankaKontrol == 0)
                        {
                            bankalar.Add(new KeyValuePair<string, double>(dataGridView1.Rows[i].Cells["BankaAdı"].Value.ToString(), Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value)));

                        }
                        toplam += Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value);
                    }
                }
                else
                    MessageBox.Show("Lütfen Çek No Giriniz");
            }
            else if (radioButton2.Checked)
            {
                DateTime pazar, cumartesi;
                string query;
                if (dateTimePicker1.Value.ToString("dddd") == "Pazartesi")
                {
                    pazar = dateTimePicker1.Value.AddDays(-1);
                    cumartesi = dateTimePicker1.Value.AddDays(-2);
                    query = "Select *from TblFaturalar where (OdemeTarih='" + dateTimePicker1.Value.ToShortDateString() + "' or OdemeTarih='" + pazar.ToShortDateString() + "' or OdemeTarih='" + cumartesi.ToShortDateString() + "') ";
                }
                else
                {
                    query = "Select *from TblFaturalar where OdemeTarih='" + dateTimePicker1.Value.ToShortDateString() + "' ";

                }
                if (comboBox3.SelectedIndex != -1)
                    query += "and BankaAdı='" + comboBox3.SelectedItem.ToString() + "'";
                da = new OleDbDataAdapter(query, con);
                ds = new DataSet();

                if (con.State == ConnectionState.Closed)
                    con.Open();
                da.Fill(ds, "TblFaturalar");
                dataGridView1.DataSource = ds.Tables["TblFaturalar"];
                con.Close();
                dataGridView1.Columns[0].Visible = false; //KOLON GİZLEME

                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    DateTime d = Convert.ToDateTime(dataGridView1.Rows[i].Cells[2].Value);
                    if (d.ToString("dddd") == "Cumartesi")
                        d = d.AddDays(2);
                    else if (d.ToString("dddd") == "Pazar")
                        d = d.AddDays(1);
                    dataGridView1.Rows[i].Cells[2].Value = d.ToShortDateString();
                    if (dataGridView1.Rows[i].Cells[7].Value.ToString() == "Ödenmemiş")
                    {
                        toplam += Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value);
                        int bankaKontrol = 0;
                        for (int k = 0; k < bankalar.Count; k++)
                        {
                            if (bankalar[k].Key == dataGridView1.Rows[i].Cells["BankaAdı"].Value.ToString())
                            {
                                bankalar[k] = new KeyValuePair<string, double>(dataGridView1.Rows[i].Cells["BankaAdı"].Value.ToString(), bankalar[k].Value + Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value));
                                bankaKontrol = 1;
                                break;
                            }
                        }
                        if (bankaKontrol == 0)
                        {
                            bankalar.Add(new KeyValuePair<string, double>(dataGridView1.Rows[i].Cells["BankaAdı"].Value.ToString(), Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value)));

                        }
                    }


                }
                DataGridViewColumn column = dataGridView1.Columns[6];
                column.Width = 265;
            }
            else
            {
                MessageBox.Show("Lütfen Seçeneklerden Birini Seçiniz!");
            }
            lblSorgulaToplam.Text = toplam.ToString() + " TL";

            for (int i = 0; i < bankalar.Count; i++)
            {
                listBox1.Items.Add(bankalar[i].Key + " : " + bankalar[i].Value + " TL");
            }


        }
        void faturaListele()
        {
            da = new OleDbDataAdapter("Select *from TblFaturalar order by CDate(OdemeTarih) asc", con);
            ds = new DataSet();

            if (con.State == ConnectionState.Closed)
                con.Open();
            da.Fill(ds, "TblFaturalar");
            dataGridView2.DataSource = ds.Tables["TblFaturalar"];
            con.Close();
            dataGridView2.Columns[0].Visible = false; //KOLON GİZLEME
            List<int> gizlenecekler = new List<int>();
            Double toplamOdenmis = 0;
            Double toplamOdenmemis = 0;
            DataGridViewColumn column = dataGridView2.Columns[6];
            column.Width = 185;
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {

                TimeSpan fark = Convert.ToDateTime(dataGridView2.Rows[i].Cells[2].Value.ToString()) - DateTime.Now;
                if (fark.TotalDays < 0 && dataGridView2.Rows[i].Cells[7].Value.ToString() != "Ödenmiş")
                    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.Red;

                string[] parcalar = dataGridView2.Rows[i].Cells[2].Value.ToString().Split('.');
                if (Convert.ToInt16(parcalar[1]) != (comboBox1.SelectedIndex + 1))
                {
                    gizlenecekler.Add(i);
                }
                else
                {
                    if (dataGridView2.Rows[i].Cells[7].Value.ToString() != "Ödenmiş")
                        toplamOdenmemis += Convert.ToDouble(dataGridView2.Rows[i].Cells[5].Value.ToString());
                    else
                        toplamOdenmis += Convert.ToDouble(dataGridView2.Rows[i].Cells[5].Value.ToString());
                }


            }
            int sayac = 0;
            for (int i = 0; i < gizlenecekler.Count; i++)
            {
                dataGridView2.Rows.RemoveAt(gizlenecekler[i] - sayac);
                sayac++;
            }
            lblToplamOdenmemis.Text = toplamOdenmemis + " TL";
            lblToplamOdenmis.Text = toplamOdenmis + " TL";
        }
        private void button4_Click(object sender, EventArgs e)
        {
            faturaListele();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            k = 0;
            i = 0; j = 1;
            x = 180;
            this.printDocument1.DefaultPageSettings.Landscape = true;
            printPreviewDialog1.Document = printDocument1;
            printPreviewDialog1.ShowDialog();

        }

        StringFormat ortahizala = new StringFormat(); //ortadan hizalama

        StringFormat sol = new StringFormat(); //soldan hizalama
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

            try
            {

                Pen myPen = new Pen(Color.Black);// Kalem Oluşturma

                System.Drawing.Font baslik = new System.Drawing.Font("Arial", 9, FontStyle.Bold);
                System.Drawing.Font altbaslik = new System.Drawing.Font("Arial", 8, FontStyle.Regular);
                System.Drawing.Font dipnot = new System.Drawing.Font("Arial", 7, FontStyle.Regular);

                ortahizala.Alignment = StringAlignment.Center;//hizalama
                sol.Alignment = StringAlignment.Near;//soldan hizalama


                e.Graphics.DrawString("Rapor", baslik, Brushes.Black, 585, 40, ortahizala);// yazı başlığı

                e.Graphics.DrawString("Düzenlenme Tarihi: " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString(), altbaslik, Brushes.Black, 850, 770);
                e.Graphics.DrawLine(myPen, 20, 760, 1160, 760); // Çizgi çizdik...




                System.Drawing.Printing.PageSettings p = printDocument1.DefaultPageSettings;



                e.Graphics.DrawLine(new Pen(Color.Black, 1), 20, 170, 1160, 170);// üst çizgi

                e.Graphics.DrawString("Sıra No", baslik, Brushes.Black, 20, 180, sol);

                e.Graphics.DrawString("Çek No", baslik, Brushes.Black, 80, 180, sol);

                e.Graphics.DrawString("Ödeme Tarih", baslik, Brushes.Black, 200, 180, sol);

                e.Graphics.DrawString("Banka Adı", baslik, Brushes.Black, 350, 180, sol);

                e.Graphics.DrawString("Çek Sahibi", baslik, Brushes.Black, 500, 180, sol);

                e.Graphics.DrawString("Miktar", baslik, Brushes.Black, 650, 180, sol);

                e.Graphics.DrawString("Açıklama", baslik, Brushes.Black, 720, 180, sol);

                e.Graphics.DrawString("Vergi No", baslik, Brushes.Black, 990, 180, sol);

                e.Graphics.DrawString("Durum", baslik, Brushes.Black, 1090, 180, sol);
                e.Graphics.DrawLine(new Pen(Color.Black, 1), 20, 200, 1160, 200);//alt çizgi



                while (i < dataGridView2.Rows.Count - 1)//iki adet döngü oluşturduk. i döngüsü ile tadagrivleri yazdırıyoruz,, j döngüsü ile sıra numarası verdiriyoruz...
                {

                    x += 25;
                    e.Graphics.DrawString((i + 1).ToString(), altbaslik, Brushes.Black, 20, x, sol);// sıra
                    e.Graphics.DrawString(dataGridView2.Rows[i].Cells[1].Value.ToString(), altbaslik, Brushes.Black, 80, x, sol);// CEK
                    e.Graphics.DrawString(String.Format("{0:D}", Convert.ToDateTime(dataGridView2.Rows[i].Cells[2].Value.ToString())), altbaslik, Brushes.Black, 200, x);//ODEME
                    e.Graphics.DrawString(dataGridView2.Rows[i].Cells[3].Value.ToString(), altbaslik, Brushes.Black, 350, x, sol);//BANKA
                    e.Graphics.DrawString(TruncateLongString(dataGridView2.Rows[i].Cells[4].Value.ToString(), 15), altbaslik, Brushes.Black, 500, x, sol);//ÇEK SAHIBI
                    e.Graphics.DrawString(dataGridView2.Rows[i].Cells[5].Value.ToString(), altbaslik, Brushes.Black, 650, x, sol);//MİKTAR
                    e.Graphics.DrawString(TruncateLongString(dataGridView2.Rows[i].Cells[6].Value.ToString(), 40), altbaslik, Brushes.Black, 720, x, sol);//AÇIKLAMA
                    e.Graphics.DrawString(dataGridView2.Rows[i].Cells[8].Value.ToString(), altbaslik, Brushes.Black, 990, x, sol);//VERGI NO
                    e.Graphics.DrawString(dataGridView2.Rows[i].Cells[7].Value.ToString(), altbaslik, Brushes.Black, 1090, x, sol);//DURUM
                    e.Graphics.DrawLine(new Pen(Color.Black, 1), 20, x + 20, 1160, x + 20);


                    i++;
                    j++;



                    if (k > 17)
                    {
                        x = 180;
                        k = 0;
                        e.HasMorePages = true;
                        return;//It will call PrintPage event again

                    }
                    else // if the number of item(per page) is more than 20 then add one page
                    {
                        k++;
                        e.HasMorePages = false; //e.HasMorePages raised the PrintPage event once per page .
                    }

                }


                
                k = 0;
                i = 0; j = 1;
                x = 180;
            }
     


            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());

            }
        }

        private string TruncateLongString(string str, int maxLength)
        {
            if (string.IsNullOrEmpty(str))
                return str;
            return str.Substring(0, Math.Min(str.Length, maxLength));
        }

        private void button6_Click(object sender, EventArgs e)
        {
            int count = 0;
            foreach (DataGridViewRow drv in dataGridView1.SelectedRows)
            {
                if (drv.Cells[1].Value != null)
                {
                    string sorgu = "Update TblFaturalar Set Durum='Ödenmiş' Where CekNo='" + drv.Cells[1].Value.ToString() + "'";
                    OleDbCommand komut = new OleDbCommand(sorgu, con);

                    if (con.State == ConnectionState.Closed)
                        con.Open();
                    komut.ExecuteNonQuery();
                    con.Close();
                    count++;
                }
            }
            if (count > 0)
                MessageBox.Show("Toplam " + count + " Düzenleme Yapılmıştır Lütfen Tekrar Listeleme Yapınız!");
            else
                MessageBox.Show("Seçili kayıt yoktur.");
        }
        int kayitKontrol = 0;
        private void button7_Click(object sender, EventArgs e)
        {
            String sorgu = "Select *from TblFaturalar where CekNo='" + tbCekNo.Text + "'";
            OleDbCommand command = new OleDbCommand(sorgu, con);

            if (con.State == ConnectionState.Closed)
                con.Open();
            OleDbDataReader reader = command.ExecuteReader();
            if (reader.Read())
            {
                btnEkle.Text = "Düzenle";
                tbCekSahip.Text = reader.GetString(4);
                dpOdemeTarih.Text = reader.GetString(2);
                cbBanka.Text = reader.GetString(3);
                tbMiktar.Text = reader.GetString(5);
                rbAciklama.Text = reader.GetString(6);
                if (reader["Vergi"] != null)
                    tbVergi.Text = reader["Vergi"].ToString();
                if (reader["VergiDairesi"] != null)
                    rbVergiDairesi.Text = reader["VergiDairesi"].ToString();
                kayitKontrol = 1;
            }
            else
            {
                MessageBox.Show("Kayıt Bulunamamıştır!");
                btnEkle.Text = "Ekle";
            }
            con.Close();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (kayitKontrol == 1)
            {
                string sorgu = "Delete from TblFaturalar where CekNo='" + tbCekNo.Text + "'";
                OleDbCommand komut = new OleDbCommand(sorgu, con);

                if (con.State == ConnectionState.Closed)
                    con.Open();
                komut.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Kayıt Silinmiştir");
                Temizle();

            }
            else
                MessageBox.Show("Lütfen Önce Kaydı Bulunuz!.");
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            Double toplam = 0;
            da = new OleDbDataAdapter("Select *from TblFaturalar where BankaAdı='" + comboBox3.SelectedItem.ToString() + "'", con);
            ds = new DataSet();

            if (con.State == ConnectionState.Closed)
                con.Open();
            da.Fill(ds, "TblFaturalar");
            dataGridView1.DataSource = ds.Tables["TblFaturalar"];
            con.Close();
            dataGridView1.Columns[0].Visible = false; //KOLON GİZLEME

            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                DateTime d = Convert.ToDateTime(dataGridView1.Rows[i].Cells[2].Value);
                if (d.ToString("dddd") == "Cumartesi")
                    d = d.AddDays(2);
                else if (d.ToString("dddd") == "Pazar")
                    d = d.AddDays(1);

                dataGridView1.Rows[i].Cells[2].Value = d.ToShortDateString();
                if (dataGridView1.Rows[i].Cells[7].Value.ToString() == "Ödenmemiş")
                    toplam += Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value);
            }
            lblSorgulaToplam.Text = toplam + " TL";
        }

        private void button9_Click(object sender, EventArgs e)
        {
            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=" + tbVeriTabani.Text + ".accdb");
            bankaListele();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            int count = 0;
            foreach (DataGridViewRow drv in dataGridView2.SelectedRows)
            {
                if (drv.Cells[1].Value != null)
                {
                    string sorgu = "Delete from TblFaturalar where CekNo='" + drv.Cells[1].Value.ToString() + "'";
                    OleDbCommand komut = new OleDbCommand(sorgu, con);

                    if (con.State == ConnectionState.Closed)
                        con.Open();
                    komut.ExecuteNonQuery();
                    con.Close();
                    count++;

                }

            }
            if (count > 0)
            {
                MessageBox.Show("Toplam " + count + " adet kayıt silinmiştir.");
            }
            else
            {
                MessageBox.Show("Lütfen silinecek kayıtları seçiniz");

            }
            faturaListele();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void tAKVİMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Takvim t = new Takvim(tbVeriTabani.Text);
            t.Show();
        }
    }
}
