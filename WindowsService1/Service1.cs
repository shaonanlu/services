using System;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Timers;

namespace WindowsService1
{
    public partial class Service1 : ServiceBase
    {
        FileInfo fi;
        StringBuilder sb;
        DirectoryInfo dirInfo;
        FileSystemWatcher _watch = new FileSystemWatcher();
        SqlConnection conn = new SqlConnection("Data Source=.\\sqlexpress;Initial Catalog=txtdatabase;User Id=admin;Password=admin;");
        SqlCommand cmd = new SqlCommand();
        public Service1()
        {
            InitializeComponent();
            this.AutoLog = false;
            if (!System.Diagnostics.EventLog.SourceExists("MySource"))
            {
                System.Diagnostics.EventLog.CreateEventSource("MySource", "MyLog");
            }
            eventLog1.Source = "MySource";
        }

        protected override void OnStart(string[] args)
        {
            eventLog1.WriteEntry("Start.");
            //設定所要監控的資料夾
            _watch.Path = @"C:\";
            //設定所要監控的變更類型
            _watch.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;
            //設定所要監控的檔案
            _watch.Filter = "*.TXT";
            //設定是否監控子資料夾
            _watch.IncludeSubdirectories = true;
            //設定是否啟動元件，此部分必須要設定為 true，不然事件是不會被觸發的
            _watch.EnableRaisingEvents = true;
            //設定觸發事件
            _watch.Created += new FileSystemEventHandler(_watch_Created);
            _watch.Changed += new FileSystemEventHandler(_watch_Changed);
            _watch.Renamed += new RenamedEventHandler(_watch_Renamed);
            _watch.Deleted += new FileSystemEventHandler(_watch_Deleted);

            try
            {
                conn.Open();
                cmd.Connection = conn;
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry(ex.Message);
            }
        }

        private void _watch_Deleted(object sender, FileSystemEventArgs e)
        {
            sb = new StringBuilder();
            sb.AppendLine("被刪除的檔名為：" + e.Name);
            sb.AppendLine("檔案所在位址為：" + e.FullPath.Replace(e.Name, ""));
            sb.AppendLine("刪除時間：" + DateTime.Now.ToString());
            eventLog1.WriteEntry(sb.ToString());
        }

        private void _watch_Renamed(object sender, RenamedEventArgs e)
        {
            sb = new StringBuilder();
            fi = new FileInfo(e.FullPath.ToString());
            sb.AppendLine("檔名更新前：" + e.OldName.ToString());
            sb.AppendLine("檔名更新後：" + e.Name.ToString());
            sb.AppendLine("檔名更新前路徑：" + e.OldFullPath.ToString());
            sb.AppendLine("檔名更新後路徑：" + e.FullPath.ToString());
            sb.AppendLine("建立時間：" + fi.LastAccessTime.ToString());
            eventLog1.WriteEntry(sb.ToString());
        }

        private void _watch_Changed(object sender, FileSystemEventArgs e)
        {
            sb = new StringBuilder();
            dirInfo = new DirectoryInfo(e.FullPath.ToString());
            sb.AppendLine("被異動的檔名為：" + e.Name);
            sb.AppendLine("檔案所在位址為：" + e.FullPath.Replace(e.Name, ""));
            sb.AppendLine("異動內容時間為：" + dirInfo.LastWriteTime.ToString());
            eventLog1.WriteEntry(sb.ToString());
        }

        private void _watch_Created(object sender, FileSystemEventArgs e)
        {
            string a, b;
            sb = new StringBuilder();
            dirInfo = new DirectoryInfo(e.FullPath.ToString());
            //sb.AppendLine("新建檔案於：" + dirInfo.FullName.Replace(dirInfo.Name, ""));
            a = dirInfo.FullName.Replace(dirInfo.Name, "");
            //sb.AppendLine("新建檔案名稱：" + dirInfo.Name);
            b = dirInfo.Name;
            sb.AppendLine("建立時間：" + dirInfo.CreationTime.ToString());
            sb.AppendLine("目錄下共有：" + dirInfo.Parent.GetFiles().Count() + " 檔案");
            sb.AppendLine("目錄下共有：" + dirInfo.Parent.GetDirectories().Count() + " 資料夾(連線有成功)");
            eventLog1.WriteEntry(sb.ToString());

            int i;
            try
            {
                cmd.CommandText = @"INSERT INTO txtFileCreated (path,filename) VALUES ('" + a + "','" + b + "')";
                i = cmd.ExecuteNonQuery();                
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry(ex.Message);
            }
            
        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("Stop.");
            conn.Close();
        }
    }
}
