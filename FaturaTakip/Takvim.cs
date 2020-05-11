using Calendar.NET;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FaturaTakip
{
    public partial class Takvim : Form
    {
        OleDbConnection con;
        OleDbDataAdapter da;
        OleDbCommand cmd;
        DataSet ds;
        String veriTabani;
        public class DayEventCount
        {
            public CustomEvent cevent;
            public DateTime date;
            public int count = 0;
        }

        [CustomRecurringFunction("RehabDates", "Calculates which days I should be getting Rehab")]
        private bool RehabDays(IEvent evnt, DateTime day)
        {
            if (day.DayOfWeek == DayOfWeek.Monday || day.DayOfWeek == DayOfWeek.Friday)
            {
                if (day.Ticks >= (new DateTime(2012, 7, 1)).Ticks)
                    return false;
                return true;
            }

            return false;
        }

        public Takvim(String veriTabani)
        {
            InitializeComponent();
            this.veriTabani = veriTabani;
            calendar1.CalendarDate = DateTime.Now;
            calendar1.CalendarView = CalendarViews.Month;
            calendar1.AllowEditingEvents = true;


        }
        [CustomRecurringFunction("Get Monday and Wednesday", "Selects the Monday and Wednesday of each month")]
        public bool GetMondayAndWednesday(IEvent evnt, DateTime dt)
        {
            if (dt.DayOfWeek == DayOfWeek.Monday || dt.DayOfWeek == DayOfWeek.Wednesday)
                return true;
            return false;
        }

        private void Takvim_Load(object sender, EventArgs e)
        {
            List<DayEventCount> dayEventCounts = new List<DayEventCount>();
            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=" + veriTabani + ".accdb");
            OleDbCommand command = new OleDbCommand("Select OdemeTarih,SUM(Miktar) from TblFaturalar GROUP BY OdemeTarih ", con);
            con.Open();
            OleDbDataReader reader = command.ExecuteReader();
            Dictionary<DateTime, Double> pztsi = new Dictionary<DateTime, Double>();
            while (reader.Read())
            {
                if (Convert.ToDateTime(reader.GetString(0)).DayOfWeek == DayOfWeek.Sunday || Convert.ToDateTime(reader.GetString(0)).DayOfWeek == DayOfWeek.Saturday || Convert.ToDateTime(reader.GetString(0)).DayOfWeek == DayOfWeek.Monday)
                {

                    DateTime time;
                    if (Convert.ToDateTime(reader.GetString(0)).DayOfWeek == DayOfWeek.Sunday)
                        time = Convert.ToDateTime(reader.GetString(0)).AddDays(1);
                    else if (Convert.ToDateTime(reader.GetString(0)).DayOfWeek == DayOfWeek.Saturday)
                        time = Convert.ToDateTime(reader.GetString(0)).AddDays(2);
                    else
                        time = Convert.ToDateTime(reader.GetString(0));

                    if (pztsi.ContainsKey(time))
                    {
                        pztsi[time] += reader.GetDouble(1);
                    }
                    else
                    {
                        pztsi.Add(time, reader.GetDouble(1));
                    }
                }
                else
                {

                    var ce = new CustomEvent();
                    ce.IgnoreTimeComponent = false;
                    ce.EventText = "Toplam :" + reader.GetDouble(1) + " TL";
                    ce.Date = Convert.ToDateTime(reader.GetString(0));
                    ce.EventLengthInHours = 2f;
                    ce.RecurringFrequency = RecurringFrequencies.None;
                    ce.EventFont = new Font("Arial", 10, FontStyle.Regular);
                    ce.Enabled = true;
                    ce.EventColor = Color.DarkBlue;
                    ce.Rank = 0;

                    calendar1.AddEvent(ce);

                }
            }
            reader.Close();
            foreach (var item in pztsi)
            {
                var ce = new CustomEvent();
                ce.IgnoreTimeComponent = false;
                ce.EventText = "Toplam :" + item.Value + " TL";
                ce.Date = item.Key;
                ce.EventLengthInHours = 2f;
                ce.RecurringFrequency = RecurringFrequencies.None;
                ce.EventFont = new Font("Arial", 10, FontStyle.Regular);
                ce.Enabled = true;
                ce.EventColor = Color.DarkBlue;
                ce.Rank = 0;
                calendar1.AddEvent(ce);

            }
            command = new OleDbCommand("Select OdemeTarih,BankaAdı,SUM(Miktar) from TblFaturalar GROUP BY OdemeTarih,BankaAdı ", con);

            reader = command.ExecuteReader();
            Dictionary<String, Double> bankalar = new Dictionary<String, Double>();
            List<DateTime> tarihler = new List<DateTime>();
            while (reader.Read())
            {
                if (Convert.ToDateTime(reader.GetString(0)).DayOfWeek == DayOfWeek.Sunday || Convert.ToDateTime(reader.GetString(0)).DayOfWeek == DayOfWeek.Saturday || Convert.ToDateTime(reader.GetString(0)).DayOfWeek == DayOfWeek.Monday)
                {
                    DateTime time;
                    if (Convert.ToDateTime(reader.GetString(0)).DayOfWeek == DayOfWeek.Sunday)
                        time = Convert.ToDateTime(reader.GetString(0)).AddDays(1);
                    else if (Convert.ToDateTime(reader.GetString(0)).DayOfWeek == DayOfWeek.Saturday)
                        time = Convert.ToDateTime(reader.GetString(0)).AddDays(2);
                    else
                        time = Convert.ToDateTime(reader.GetString(0));
                    if (bankalar.ContainsKey(time + "&" + reader.GetString(1)) && tarihler.Contains(time))
                    {
                        bankalar[time + "&" + reader.GetString(1)] += reader.GetDouble(2);
                    }
                    else
                    {
                        bankalar.Add(time + "&" + reader.GetString(1), reader.GetDouble(2));
                        tarihler.Add(time);
                    }
                }
                else
                {
                    int dayEventRow = -1;
                    int row = 0;
                    foreach (var item in dayEventCounts)
                    {
                        if (Convert.ToDateTime(reader.GetString(0)) == item.date)
                        {
                            dayEventRow = row;
                        }
                        row++;
                    }
                    if (dayEventRow != -1)
                    {
                        dayEventCounts[dayEventRow].count++;
                    }
                    else
                    {
                        DayEventCount dayEventCount = new DayEventCount();
                        dayEventCount.date = Convert.ToDateTime(reader.GetString(0));
                        dayEventCount.count = 1;
                        dayEventCounts.Add(dayEventCount);
                    }

                    if (dayEventRow == -1 || dayEventCounts[dayEventRow].count <= 5)
                    {
                        var ce = new CustomEvent();
                        ce.IgnoreTimeComponent = false;
                        if (reader.GetString(1).Length < 17)
                            ce.EventText = reader.GetString(1) + ":" + reader.GetDouble(2) + " TL";
                        else
                            ce.EventText = reader.GetString(1).Substring(0, 15) + "..:" + reader.GetDouble(2) + " TL";
                        ce.Date = Convert.ToDateTime(reader.GetString(0));
                        ce.EventLengthInHours = 2f;
                        ce.RecurringFrequency = RecurringFrequencies.None;
                        ce.EventFont = new Font("Arial", 10, FontStyle.Regular);
                        ce.Enabled = true;
                        calendar1.AddEvent(ce);
                        if (dayEventRow != -1&&dayEventCounts[dayEventRow].count == 5)
                        {
                            dayEventCounts[dayEventRow].cevent = ce;
                        }
                    }
                }
            }
            for (int i = 0; i < bankalar.Count; i++)
            {
                int dayEventRow = -1;
                int row = 0;
                foreach (var item in dayEventCounts)
                {
                    if (tarihler[i] == item.date)
                    {
                        dayEventRow = row;
                    }
                    row++;
                }
                if (dayEventRow != -1)
                {
                    dayEventCounts[dayEventRow].count++;
                }
                else
                {
                    DayEventCount dayEventCount = new DayEventCount();
                    dayEventCount.date = tarihler[i];
                    dayEventCount.count = 1;
                    dayEventCounts.Add(dayEventCount);

                }

                if (dayEventRow == -1 || dayEventCounts[dayEventRow].count <= 5)
                {
                    var ce = new CustomEvent();
                    ce.IgnoreTimeComponent = false;
                    if (bankalar.ElementAt(i).Key.Split('&')[1].Length < 17)
                        ce.EventText = bankalar.ElementAt(i).Key.Split('&')[1] + ":" + bankalar.ElementAt(i).Value + " TL";
                    else
                        ce.EventText = bankalar.ElementAt(i).Key.Split('&')[1].Substring(0, 15) + "..:" + bankalar.ElementAt(i).Value + " TL";
                    ce.Date = tarihler[i];
                    ce.EventLengthInHours = 2f;
                    ce.EventLengthInHours = 2f;
                    ce.RecurringFrequency = RecurringFrequencies.None;
                    ce.EventFont = new Font("Arial", 10, FontStyle.Regular);
                    ce.Enabled = true;
                    calendar1.AddEvent(ce);
                    if (dayEventRow != -1&&dayEventCounts[dayEventRow].count == 5)
                    {
                        dayEventCounts[dayEventRow].cevent = ce;
                    }
                }
            }
            foreach (var item in dayEventCounts)
            {
                if (item.count > 5)
                {
                    calendar1.RemoveEvent(item.cevent);
                    var ce = new CustomEvent();
                    ce.IgnoreTimeComponent = false;
                    ce.EventText = "Devamı var...";
                    item.date = item.date.AddHours(23);
                    ce.Date = item.date;
                    ce.EventLengthInHours = 2f;
                    ce.EventLengthInHours = 2f;
                    ce.RecurringFrequency = RecurringFrequencies.None;
                    ce.EventFont = new Font("Arial", 10, FontStyle.Regular);
                    ce.Enabled = true;
                    calendar1.AddEvent(ce);

                }
            }

            con.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PrintDocument doc = new PrintDocument();
            doc.PrintPage += this.Doc_PrintPage;
            PrintDialog dlgSettings = new PrintDialog();
            dlgSettings.Document = doc;
            if (dlgSettings.ShowDialog() == DialogResult.OK)
            {
                doc.Print();
            }
        }
        private void Doc_PrintPage(object sender, PrintPageEventArgs e)
        {
            Size s = calendar1.Size;
            Calendar.NET.Calendar current = calendar1;
            double rate = current.Width / 600;
            current.Width = 600;
            current.Height = Convert.ToInt32(current.Height / rate);
            float x = e.MarginBounds.Left;
            float y = e.MarginBounds.Top;
            Bitmap bmp = new Bitmap(current.Width, current.Height);
            current.DrawToBitmap(bmp, new Rectangle(0, 0, current.Width, current.Height));
            e.Graphics.DrawImage((Image)bmp, x, y);
            calendar1.Size = s;

        }

    }
}
