using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using OpenQA.Selenium.Chrome;
using System.Text.RegularExpressions;
using RestSharp;
using System.Data.OleDb;
using System.Threading;
using System.Net;

namespace ServiceNotifEreporting
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("Service is started at " + DateTime.Now);
            SendTelegram("-778112650", "Service Notif E-Reporting start TimeStamp : " + DateTime.Now.ToString("hh:mm:ss"));
            string file = AppDomain.CurrentDomain.BaseDirectory + "\\numberregistered.txt";
            string[] text = File.ReadAllLines(file);
            List<string> number = new List<string>();
            foreach (string item in text)
            {
                string[] numberdata = item.Split(' ');
                number.Add(numberdata[0]);
            }
            foreach (var item in number)
            {
                SendMessage(item, "Service Notif E-Reporting start TimeStamp : " + DateTime.Now.ToString("hh:mm:ss"));
            }
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 60000;  //1 menit
            timer.Elapsed += new ElapsedEventHandler(this.OnTimer);
            timer.Start();
        }
        public void OnTimer(object sender, ElapsedEventArgs args)
        {
            try
            {
                DayOfWeek day = DateTime.Now.DayOfWeek;
                if ((day != DayOfWeek.Saturday) || (day != DayOfWeek.Sunday))
                {
                    string file = AppDomain.CurrentDomain.BaseDirectory + "\\numberregistered.txt";
                    string[] text = File.ReadAllLines(file);
                    List<string> number = new List<string>();
                    foreach (string item in text)
                    {
                        string[] numberdata = item.Split(' ');
                        number.Add(numberdata[0]);
                    }
                    var jam = DateTime.Now.Hour;
                    var menit = DateTime.Now.Minute;
                    if (jam > 7 && jam < 12 && (menit == 15 || menit == 30 || menit == 45 || menit == 00))
                    {
                        var total = cek_web();
                        if (total == 0)
                        {
                            SendTelegram("-778112650", "Notification E-Reporting\n\nThere is no data that needs to be validated\nTimeStamp : " + DateTime.Now.ToString("hh:mm:ss"));
                            foreach (var item in number)
                            {
                                SendMessage(item, "Notification E-Reporting\n\nThere is no data that needs to be validated\nTimeStamp : " + DateTime.Now.ToString("HH:mm"));
                            }
                        }
                    }
                    if (jam == 11 && menit == 00 || jam == 13 && menit == 00 || jam == 12 && menit == 00 || jam == 13 && menit == 59)
                    {
                        var total = cek_web();
                        if (total == 0)
                        {
                            SendTelegram("-778112650", "Notification E-Reporting\n\nThere is no data that needs to be validated\nTimeStamp : " + DateTime.Now.ToString("hh:mm:ss"));
                            foreach (var item in number)
                            {
                                SendMessage(item, "Notification E-Reporting\n\nThere is no data that needs to be validated\nTimeStamp : " + DateTime.Now.ToString("HH:mm"));
                            }
                        }
                    }
                    if (jam > 12 && jam < 14 && (menit == 5 || menit == 10 || menit == 15 || menit == 20 || menit == 25 || menit == 30 || menit == 35 || menit == 40 || menit == 45 || menit == 50 || menit == 55 || menit == 00))
                    {
                        var total = cek_web();
                        if (total == 0)
                        {
                            SendTelegram("-778112650", "Notification E-Reporting\n\nThere is no data that needs to be validated\nTimeStamp : " + DateTime.Now.ToString("hh:mm:ss"));
                            foreach (var item in number)
                            {
                                SendMessage(item, "Notification E-Reporting\n\nThere is no data that needs to be validated\nTimeStamp : " + DateTime.Now.ToString("HH:mm"));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SendTelegram("-778112650", "Process notification e reporting failed : " + ex.Message + "/n" + DateTime.Now.ToString("hh:mm:ss"));
            }
            //cekpenempatanmargin();
        }
        public static int cek_web()
        {
            List<separator> finaldata = new List<separator>();
            IWebDriver ChromeDriver = new ChromeDriver();
            try
            {
                monitoringServices("SPINotifereporting", "Service untuk notifikasi e reporting bappebti", "10.10.10.99", "Live");

                ////Thread.Sleep(15000);
                ChromeDriver.Navigate().GoToUrl("http://lapkeu.bappebti.go.id/login/login.php");

                IWebElement username = ChromeDriver.FindElement(By.Id("id_sc_field_login"));
                IWebElement password = ChromeDriver.FindElement(By.Id("id_sc_field_password"));
                IWebElement submit_login = ChromeDriver.FindElement(By.Id("sub_form_b"));

                username.SendKeys("sarahptkbi.com");
                password.SendKeys("12345");
                submit_login.Click();
                IWebElement verify = ChromeDriver.FindElement(By.ClassName("scGridFieldOddLink"));
                verify.Click();

                ChromeDriver.Navigate().GoToUrl("http://lapkeu.bappebti.go.id/validasi_margin_list/validasi_margin_list.php?sc_item_menu=item_317&amp");

                IWebElement table2 = ChromeDriver.FindElement(By.Id("apl_validasi_margin_list#?#1"));

                var html_text = table2.GetAttribute("innerHTML");

                string allText = html_text;//File.ReadAllText(pathHTML);
                var htmlremove = Regex.Replace(allText, "<.*?>", String.Empty);
                var data = htmlremove.Split(new[] { "\r\n\r\n    \r\n     \r\n\r\n     " }, StringSplitOptions.None);
                var year = DateTime.Now.Year;
                for (int i = 1; i < data.Count(); i++)
                {
                    var text = data[i].Split(new[] { "\r\n     " }, StringSplitOptions.None);
                    var split_date = text[3].Split('/');
                    if (text[7] == "Belum Divalidasi" && split_date[2] == year.ToString())
                    {
                        finaldata.Add(new separator
                        {
                            nama = text[0],
                            tgllaporan = text[1],
                            matauang = text[2],
                            tgl_tf = text[3],
                            waktu = text[4],
                            jumlah = text[5],
                            bank = text[6],
                            status = text[7]
                        });
                    }
                }

                WriteToFile(finaldata.Count.ToString());
                if (finaldata.Count == 0)
                {
                    WriteToFile("tidak ada data !");
                }
                else
                {
                    List<string> data_send = new List<string>();
                    int no = 1;
                    foreach (var item in finaldata)
                    {
                        data_send.Add(no + ". Nama : " + item.nama + "\n Tanggal Laporan: " + item.tgllaporan + "\n Mata Uang: " + item.matauang + "\n Tanggal Transfer: " + item.tgl_tf + "\n Waktu Transfer: " + item.waktu + "\n Jumlah: " + item.jumlah + "\n Bank: " + item.bank + "\n Status: " + item.status);
                        no++;
                    }

                    SendTelegram("-778112650", "Notification E-Reporting\n\n" + string.Join("\n\n", data_send) + "\nFrom Whatsapp API PT KBI Testing\n\nTimeStamp : " + DateTime.Now.ToString("hh:mm:ss"));
                    
                    string file = AppDomain.CurrentDomain.BaseDirectory + "\\numberregistered.txt";
                    string[] text = File.ReadAllLines(file);
                    List<string> number = new List<string>();
                    foreach (string item in text)
                    {
                        string[] numberdata = item.Split(' ');
                        number.Add(numberdata[0]);
                    }
                    foreach (var item in number)
                    {
                        SendMessage(item, "Notification E-Reporting\n\n" + string.Join("\n\n", data_send) + "\nFrom Whatsapp API PT KBI Testing\n\nTimeStamp : " + DateTime.Now.ToString("hh:mm:ss"));
                    }
                }
                ChromeDriver.Quit();

            }
            catch (Exception ex)
            {
                ChromeDriver.Quit();
                WriteToFile("Get data fail : " + ex.Message);
            }
            return (finaldata.Count);
        }
        public static void cekpenempatanmargin()
        {
            var chat_id = "6281380048440-1614311302@g.us";
            try
            {
                //IWebDriver ChromeDriver = new ChromeDriver();
                //ChromeDriver.Navigate().GoToUrl("http://lapkeu.bappebti.go.id/login/login.php");

                //IWebElement username = ChromeDriver.FindElement(By.Id("id_sc_field_login"));
                //IWebElement password = ChromeDriver.FindElement(By.Id("id_sc_field_password"));
                //IWebElement submit_login = ChromeDriver.FindElement(By.Id("sub_form_b"));

                //username.SendKeys("sarahptkbi.com");
                //password.SendKeys("12345");
                //submit_login.Click();

                //IWebElement verify = ChromeDriver.FindElement(By.ClassName("scGridFieldOddLink"));
                //verify.Click();

                //ChromeDriver.Navigate().GoToUrl("http://lapkeu.bappebti.go.id/persentase_margin_list/persentase_margin_list.php");

                //IWebElement export = ChromeDriver.FindElement(By.XPath("//img[@title='Ekspor data ke format excel']"));
                //export.Click();
                //Thread.Sleep(30000);
                IWebDriver ChromeDriver = new ChromeDriver();
                ChromeDriver.Navigate().GoToUrl("http://lapkeu.bappebti.go.id/login/login.php");

                IWebElement username = ChromeDriver.FindElement(By.Id("id_sc_field_login"));
                IWebElement password = ChromeDriver.FindElement(By.Id("id_sc_field_password"));
                IWebElement submit_login = ChromeDriver.FindElement(By.Id("sub_form_b"));

                username.SendKeys("sarahptkbi.com");
                password.SendKeys("12345");
                submit_login.Click();
                IWebElement verify1 = ChromeDriver.FindElement(By.ClassName("scGridFieldOddLink"));
                verify1.Click();
                ChromeDriver.Navigate().GoToUrl("http://lapkeu.bappebti.go.id/persentase_margin_list/persentase_margin_list.php");
                ChromeDriver.Manage().Window.Size = new System.Drawing.Size(1552, 840);
                //vars["WindowHandles"] = ChromeDriver.WindowHandles;
                ChromeDriver.FindElement(By.CssSelector("tbody:nth-child(1) img")).Click();

                List<parameter> data_mentah = new List<parameter>();
                string datelong = DateTime.Now.AddDays(-1).ToString("dd MMMM yyyy", new System.Globalization.CultureInfo("id-ID"));
                string[] parse = datelong.Split(' ');
                int tgl = (Convert.ToInt32(parse[0]));
                string path = @"C:\Users\hasto\Downloads\PERSENTASE PENEMPATAN MARGIN Periode " + tgl.ToString() + " " + parse[1] + " " + parse[2] + ".xlsx";
                //File.Copy(path, @"E:\Service_e_reporting\PERSENTASE PENEMPATAN MARGIN Periode " + tgl.ToString() + " " + parse[1] + " " + parse[2] + ".xlsx");
                string con = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=Yes;';";
                using (OleDbConnection connection = new OleDbConnection(con))
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand("Select * From [Worksheet$]", connection);
                    using (OleDbDataReader dr = command.ExecuteReader())
                    {

                        while (dr.Read())
                        {
                            if (dr[14].ToString() == "dibawah ketentuan" && dr[20].ToString() != "")
                            {
                                var nama = dr[1].ToString();
                                var data = new parameter
                                {
                                    nama = dr[1].ToString(),
                                    periodeLaporan = dr[2].ToString(),
                                    totalSebelum = dr[12].ToString() + " %",
                                    totalSesudah = dr[13].ToString() + " %",
                                    statusEntri = dr[18].ToString(),
                                    statusValidasi = dr[20].ToString(),
                                    waktuValidasi = dr[19].ToString()
                                };
                                data_mentah.Add(data);
                            }
                        }
                    }
                    connection.Close();
                }
                //ChromeDriver.Quit();
                File.Delete(@"E:\Service_e_reporting\PERSENTASE PENEMPATAN MARGIN Periode " + tgl.ToString() + " " + parse[1] + " " + parse[2] + ".xlsx");
                List<string> data_send = new List<string>();
                int no = 1;
                foreach (var item in data_mentah)
                {
                    data_send.Add(no + ". Nama : " + item.nama + "\n Periode Laporan: " + item.periodeLaporan + "\n Total Sebelum (%): " + item.totalSebelum + "\n Total Sesudah (%): " + item.totalSesudah + "\n Status Validasi: " + item.statusValidasi + "\n Status Entri: " + item.statusEntri + "\n Waktu Validasi: " + item.waktuValidasi);
                    no++;
                }
                SendMessage(chat_id, "#Notification E-Reporting\n\nPERSENTASE PENEMPATAN MARGIN PADA LEMBAGA KLIRING" + string.Join("\n\n", data_send) + "\n\nTimeStamp : " + DateTime.Now.AddMinutes(4).ToString("hh:mm:ss"));
            }
            catch (Exception x)
            {
                WriteToFile(x.Message);
            }
        }
        private static void SendMessage(string chatId, string body)
        {
            var client = new RestClient("https://api.1msg.io/127354/sendMessage?token=jkdjtwjkwq2gfkac");
            client.Timeout = -1;
            var requestWa = new RestRequest(Method.POST);
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("phone", chatId);
            requestWa.AddParameter("body", body);
            IRestResponse responseWa = client.Execute(requestWa);
            WriteToFile(responseWa.Content);
        }
        public class separator
        {
            public string nama { get; set; }
            public string tgllaporan { get; set; }
            public string matauang { get; set; }
            public string tgl_tf { get; set; }
            public string waktu { get; set; }
            public string jumlah { get; set; }
            public string bank { get; set; }
            public string status { get; set; }
        }
        public class parameter
        {
            public string nama { get; set; }
            public string periodeLaporan { get; set; }
            public string totalSebelum { get; set; }
            public string totalSesudah { get; set; }
            public string statusEntri { get; set; }
            public string waktuValidasi { get; set; }
            public string statusValidasi { get; set; }

        }
        public static void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }
        protected override void OnStop()
        {
            SendTelegram("-778112650", "Service Notif E-Reporting stopped TimeStamp : " + DateTime.Now.ToString("hh:mm:ss"));
        }
        private static void SendTelegram(string chatId, string body)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.DefaultConnectionLimit = 9999;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls12;

            var client = new RestClient("https://api.telegram.org/bot5478187618:AAENfPcaia3OMwc3alj57qil0uN7JrPFPP4/sendMessage?chat_id=" + chatId + "&text=" + body);
            RestRequest requestWa = new RestRequest("https://api.telegram.org/bot5478187618:AAENfPcaia3OMwc3alj57qil0uN7JrPFPP4/sendMessage?chat_id=" + chatId + "&text=" + body, Method.GET);
            requestWa.Timeout = -1;
            var responseWa = client.ExecutePostAsync(requestWa);
            WriteToFile(responseWa.Result.Content);
        }
        private static string monitoringServices(string servicename, string servicedescription, string servicelocation, string appstatus)
        {
            string jsonString = "{" +
                                "\"service_name\" : \"" + servicename + "\"," +
                                "\"service_description\": \"" + servicedescription + "\"," +
                                "\"service_location\":\"" + servicelocation + "\"," +
                                "\"app_status\":\"" + appstatus + "\"," +
                                "}";
            var client = new RestClient("http://10.10.10.99:84/api/ServiceStatus");

            RestRequest requestWa = new RestRequest("http://10.10.10.99:84/api/ServiceStatus", Method.POST);
            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("data", jsonString);
            var responseWa = client.ExecutePostAsync(requestWa);
            return (responseWa.Result.Content);
        }
    }
}
