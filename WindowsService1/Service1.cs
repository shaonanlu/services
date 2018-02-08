using System.ServiceProcess;
using System.Timers;

namespace WindowsService1
{
    public partial class Service1 : ServiceBase
    {
        private Timer MyTimer;
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
            eventLog1.WriteEntry("Start Timer.");
            MyTimer = new Timer();
            MyTimer.Elapsed += new ElapsedEventHandler(MyTimer_Elapsed);
            MyTimer.Interval = 10 * 1000;
            MyTimer.Start();
        }

        private void MyTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            eventLog1.WriteEntry("Timer Ticked.");
        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("Stop Timer.");
            MyTimer.Stop();
            MyTimer = null;
        }
    }
}
