using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Xml.Linq;

namespace JISR_SERVICE_V1
{
    public partial class AttendanceService : ServiceBase
    {
        #region Private data members
        private bool isLoggedIn;
        private bool isFirstCalled;
        private DateTime? lastSavedLogDate;

        private Timer timer;
        private int eventId = 0;

        private SqlConnection connection = null;
        private SqlDataAdapter adapter = null;
        private DataSet attendanceDataset = null;
        private const string tableName = "attendance";
        #endregion

        #region Constructor, OnStart and OnStop function members
        public AttendanceService()
        {
            InitializeComponent();

            // Define custom event log
            if (!System.Diagnostics.EventLog.SourceExists("elAttendanceSource"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "elAttendanceSource", "elAttendanceLog");
            }

            elAttendance.Source = "elAttendanceSource";
            elAttendance.Log = "elAttendanceLog";

            // Initialize Timer
            timer = new Timer();
            timer.Interval = 60000;
            timer.Elapsed += new ElapsedEventHandler(this.OnTimerElapsed);

        }

        protected override void OnStart(string[] args)
        {
            elAttendance.WriteEntry("JISR Service for Attendance STARTED.", EventLogEntryType.Information, eventId++);

            // Update the service state to Start Pending.
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            // Initialize members
            this.isLoggedIn = false;
            this.isFirstCalled = true;
            this.lastSavedLogDate = null;

            // Load Configurations from xml file
            Configurations.Load(elAttendance);

            // Set Timer Interval
            timer.Interval = Configurations.TimerInterval; 

            // Ping API and login accordingly
            PingAPI().ContinueWith(pnigResponse => {

                EnableTimer();
                if (!this.isLoggedIn)
                {
                    Login().ContinueWith(loginResponse => { EnableTimer(); });
                }
            });

            // Update the service state to Running
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
            
        }

        private void EnableTimer()
        {
            // if logged in successfully then do the following: 
            if (this.isLoggedIn)
            {
                // 1. Initialize Connection and DataSet
                if (connection == null) connection = GetSqlConnection();
                if (attendanceDataset == null) attendanceDataset = new DataSet();

                // 2. Start timer to send data to api
                timer.Enabled = true;
            }
        }

        protected override void OnStop()
        {
            elAttendance.WriteEntry("JISR Service for Attendance STOPPED.", EventLogEntryType.Information, eventId++);
            timer.Enabled = false;
            connection = null;
            attendanceDataset = null;
            adapter = null;
            elAttendance.Clear();
        }
        #endregion

        #region SQL Connection and Query
        private SqlConnection GetSqlConnection()
        {
            if (connection == null)
            {
                return new SqlConnection(Configurations.ConnectionString);
            }
            else
            {
                return connection;
            }
        }

        private string GetSQL()
        {
            string tbl = GetTableName(DateTime.Now, "DeviceLogs");
            string sql = String.Format("select DeviceLogId,EmployeeCode,LogDate,Direction from {0},Employees where {0}.UserId=Employees.EmployeeId and {1} order by DeviceLogId", tbl, GetWhere(DateTime.Now));
            return sql;
        }

        private string GetTableName(DateTime date, string tbl)
        {
            int month = date.Month;
            int year = date.Year;
            return String.Format("{0}_{1}_{2}", tbl, month, year);
        }

        private string GetWhere(DateTime date)
        {
            string onlyIfFirstTime = "";

            if (this.isFirstCalled)
            {
                this.isFirstCalled = false;
            }
            else
            {
                //onlyIfFirstTime = String.Format("and  logDate between '{0}' and '{1}'", today.AddMinutes(-1), today);
                if (this.lastSavedLogDate != null)
                {
                    onlyIfFirstTime = String.Format("and  logDate > '{0}'", this.lastSavedLogDate.Value.ToString("MM/dd/yyyy HH:mm:ss"));
                }
            }

            return String.Format("cast(logDate as date) = '{0}' {1}", date.ToString("MM/dd/yyyy"), onlyIfFirstTime);
        }
        #endregion

        #region Authentication
        private async Task PingAPI()
        {
            try
            {
                // ping api
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri(Configurations.BaseAddress);

                    using (var pingResponse = await client.GetAsync(String.Format("ping?access_token={0}", Configurations.AccessToken)))
                    {
                        var pingJsonResult = await pingResponse.Content.ReadAsStringAsync();
                        Dictionary<string, string> pingDic = JsonConvert.DeserializeObject<Dictionary<string, string>>(pingJsonResult);

                        if (pingDic["success"] == "true")
                        {
                            elAttendance.WriteEntry("Status: You are logged in.", EventLogEntryType.Information, eventId++);
                            this.isLoggedIn = true;
                        }
                        else
                        {
                            elAttendance.WriteEntry("Status: You are logged out.", EventLogEntryType.Information, eventId++);
                            this.isLoggedIn = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                elAttendance.WriteEntry("Error occurs in Ping API: \n" + ex.Message, EventLogEntryType.Error, eventId++);
            }
        }

        private async Task Login()
        {

            if (!this.isLoggedIn)
            {
                // login
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri(Configurations.BaseAddress);

                    var serializedLogin = JsonConvert.SerializeObject(
                        new Login
                        {
                            login = Configurations.Login,
                            password = Configurations.Password
                        }
                    );

                    var content = new StringContent(serializedLogin, Encoding.UTF8, "application/json");
                    var result = await client.PostAsync("sessions", content);
                    var loginJsonResult = await result.Content.ReadAsStringAsync();
                    var loginDic = JsonConvert.DeserializeObject<Dictionary<string, string>>(loginJsonResult);

                    if (loginDic["success"] == "true")
                    {
                        Configurations.AccessToken = loginDic["access_token"];
                        Configurations.Save("AccessToken", loginDic["access_token"]);

                        elAttendance.WriteEntry(loginDic["message"], EventLogEntryType.SuccessAudit, eventId++);
                        this.isLoggedIn = true;
                    }
                    else
                    {
                        elAttendance.WriteEntry(loginDic["error"], EventLogEntryType.Error, eventId++);
                    }
                }
            }
            else
            {
                elAttendance.WriteEntry("You already logged in.", EventLogEntryType.Information, eventId++);
            }

        }

        private async void Logout()
        {
            if (this.isLoggedIn)
            {
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri(Configurations.BaseAddress);

                    var result = await client.DeleteAsync(String.Format("sessions?access_token={0}", Configurations.AccessToken));
                    var logoutJsonResult = await result.Content.ReadAsStringAsync();
                    Dictionary<string, string> logoutDic = JsonConvert.DeserializeObject<Dictionary<string, string>>(logoutJsonResult);

                    if (logoutDic["success"] == "true")
                    {
                        elAttendance.WriteEntry(logoutDic["message"], EventLogEntryType.SuccessAudit, eventId++);
                        this.isLoggedIn = false;
                        timer.Enabled = false;
                    }

                }
            }
            else
            {
                elAttendance.WriteEntry("You already logged out.", EventLogEntryType.Information, eventId++);
            }
        }
        #endregion

        #region Timer OnElapsed Event Handler

        public void OnTimerElapsed(object sender, System.Timers.ElapsedEventArgs args)
        {
            try
            {
                // setup sql_data_adapter and get data
                string sqlStatement = GetSQL();
                
                if (adapter == null)
                {
                    adapter = new SqlDataAdapter(sqlStatement, GetSqlConnection());
                }
                else
                {
                    adapter.SelectCommand.CommandText = sqlStatement;
                }

                // fill data in dataset
                attendanceDataset.Reset();
                adapter.Fill(attendanceDataset, tableName);

                // check if there are data need to be sent to api
                int rowsCount = attendanceDataset.Tables[tableName].Rows.Count;
                if (rowsCount > 0)
                {
                    // add notification to lbxNotifications
                    elAttendance.WriteEntry("attendance logs to be sent to api: " + attendanceDataset.Tables[tableName].Rows.Count, EventLogEntryType.Information, eventId++);

                    // send data to api
                    SendAttendanceLogsToAPI(attendanceDataset.Tables[tableName]);
                }
                else
                {
                    elAttendance.WriteEntry("no new attendance logs. no data sent to api on " + DateTime.Now, EventLogEntryType.Information, eventId++);
                }
            }
            catch (Exception ex)
            {
                elAttendance.WriteEntry("OnTimerElapsed: " + ex.Message, EventLogEntryType.Error, eventId++);
            }
        }

        private async void SendAttendanceLogsToAPI(DataTable AttendanceLogs)
        {
            // map attendance logs as needed in api params
            List<Record> logsList = MapLogsToApiParams(AttendanceLogs);
            dynamic logsListWarper = new { record = logsList };

            // serialize logs to json
            var logsSerialized = await JsonConvert.SerializeObjectAsync(logsListWarper);

            // send logs to api
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(Configurations.BaseAddress);
                var content = new StringContent(logsSerialized, Encoding.UTF8, "application/json");
                var result = await client.PostAsync("device_attendances?access_token=" + Configurations.AccessToken, content);

                var attendanceJsonResult = await result.Content.ReadAsStringAsync();
                var attendanceDic = JsonConvert.DeserializeObject<Dictionary<string, string>>(attendanceJsonResult);

                if (attendanceDic["success"] == "true")
                {
                    elAttendance.WriteEntry(String.Format("Data sent successfully on {0}. records_updated: {1}", DateTime.Now, attendanceDic["records_updated"]), EventLogEntryType.Information, eventId++);
                    // save lastSavedLogDate
                    this.lastSavedLogDate = Convert.ToDateTime(AttendanceLogs.Rows[AttendanceLogs.Rows.Count - 1]["LogDate"]);
                }
                else
                {
                    elAttendance.WriteEntry(attendanceDic["error"], EventLogEntryType.Error, eventId++);
                }
            }
        }

        private List<Record> MapLogsToApiParams(DataTable AttendanceLogs)
        {
            List<Record> list = new List<Record>();

            foreach (DataRow row in AttendanceLogs.Rows)
            {
                list.Add(
                    new Record
                    {
                        id = Convert.ToString(row["EmployeeCode"]),
                        day = Convert.ToDateTime(row["LogDate"]).ToString("dd/MM/yyyy"),
                        time = Convert.ToDateTime(row["LogDate"]).ToString("HH:mm"),
                        direction = row["Direction"].ToString()
                    }
                );
            }

            return list;
        }
        #endregion

        #region Set service status
        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(IntPtr handle, ref ServiceStatus serviceStatus);
        #endregion
    }

    #region model classes
    public class Login
    {
        public string login { get; set; }
        public string password { get; set; }
    }

    public class Record
    {
        public string id { get; set; }
        public string day { get; set; }
        public string time { get; set; }
        public string direction { get; set; }
    }

    public static class Configurations
    {
        public static string AccessToken { get; set; }
        public static string ConnectionString { get; set; }
        public static string BaseAddress { get; set; }
        public static string Login { get; set; }
        public static string Password { get; set; }
        public static double TimerInterval { get; set;  }
        private static XElement configurations { get; set; }

        private static EventLog el = null;
        public static void Load(EventLog eventLog = null)
        {
            try
            {
                if (eventLog != null) el = eventLog;
                
                configurations = XElement.Load("c:\\configurations.xml");
                XElement attendance = configurations.Elements().First();

                AccessToken = attendance.Element("AccessToken").Value;
                BaseAddress = attendance.Element("BaseAddress").Value;
                ConnectionString = attendance.Element("ConnectionString").Value;
                Login = attendance.Element("Login").Value;
                Password = attendance.Element("Password").Value;
                TimerInterval = Convert.ToDouble(attendance.Element("TimerInterval").Value);
            }
            catch (Exception ex)
            {
                if (el != null)
                {
                    el.WriteEntry("Error occurs in Configurations.Load() method: " + ex.Message);
                }
            }
        }
        public static void Save(string Node, string Value)
        {
            try
            {
                XElement attendance = configurations.Elements().First();
                attendance.SetElementValue(Node, Value);
                configurations.Save("c:\\configurations.xml");
            }
            catch (Exception ex)
            {
                if (el != null)
                {
                    el.WriteEntry("Error occurs in Configurations.Save() method: " + ex.Message);
                }
            }
        }
    }
    #endregion

    #region Ienum ans struct for implementing service pending status
    public enum ServiceState
    {
        SERVICE_STOPPED = 0x00000001,
        SERVICE_START_PENDING = 0x00000002,
        SERVICE_STOP_PENDING = 0x00000003,
        SERVICE_RUNNING = 0x00000004,
        SERVICE_CONTINUE_PENDING = 0x00000005,
        SERVICE_PAUSE_PENDING = 0x00000006,
        SERVICE_PAUSED = 0x00000007,
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct ServiceStatus
    {
        public long dwServiceType;
        public ServiceState dwCurrentState;
        public long dwControlsAccepted;
        public long dwWin32ExitCode;
        public long dwServiceSpecificExitCode;
        public long dwCheckPoint;
        public long dwWaitHint;
    };
    #endregion
}
