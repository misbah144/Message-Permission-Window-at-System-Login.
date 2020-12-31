using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.IO;

    namespace CodeofConduct
    {

        public partial class Form2 : Form
        {
            string user = Environment.UserName.ToString();
            string displayName = "";
            string department = "";

            [StructLayout(LayoutKind.Sequential)]
            private struct KBDLLHOOKSTRUCT
            {
                public Keys key;
                public int scanCode;
                public int flags;
                public int time;
                public IntPtr extra;
            }

            private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr SetWindowsHookEx(int id, LowLevelKeyboardProc callback, IntPtr hMod, uint dwThreadId);
            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern bool UnhookWindowsHookEx(IntPtr hook);
            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr CallNextHookEx(IntPtr hook, int nCode, IntPtr wp, IntPtr lp);
            [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr GetModuleHandle(string name);
            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            private static extern short GetAsyncKeyState(Keys key);

            private const int SW_HIDE = 0;
            private const int SW_SHOW = 1;

            [DllImport("user32.dll")]
            private static extern int FindWindow(string className, string windowText);

            [DllImport("user32.dll")]
            private static extern int ShowWindow(int hwnd, int command);


            private IntPtr ptrHook;
            private LowLevelKeyboardProc objKeyboardProcess;

           

            public Form2()
            {
                try
                {

                    DirectorySearcher DrSearcher = new System.DirectoryServices.DirectorySearcher("(samaccountname=" + user + ")");

                    //  DirectorySearcher DrSearcher = new System.DirectoryServices.DirectorySearcher("(samaccountname=test.user1)");

                    SearchResult SrchRes = DrSearcher.FindOne();
                    DirectoryEntry DrEntry = SrchRes.GetDirectoryEntry();

                //    DirectoryEntry DrEntry = new DirectoryEntry(@"LDAP://CN=Users,DC=almeezangroup");
                    
                    // string FirstName = DrEntry.Properties["givenName"][0].ToString();
                    // string LastName = DrEntry.Properties["sn"][0].ToString();
                    // string UserEmail = DrEntry.Properties["mail"][0].ToString();

                    //displayName = DrEntry.Properties["displayName"][0].ToString();
                    //  department = DrEntry.Properties["department"][0].ToString();



                    if (DrEntry.Properties.Contains("displayName"))
                    {
                        displayName = DrEntry.Properties["displayName"][0].ToString();
                    }
                    else
                    {
                        displayName = "";
                    }


                    if (DrEntry.Properties.Contains("department"))
                    {
                        department = DrEntry.Properties["department"][0].ToString();
                    }
                    else
                    {
                        department = "";
                    }





                    ProcessModule objCurrentModule = Process.GetCurrentProcess().MainModule;
                    objKeyboardProcess = new LowLevelKeyboardProc(captureKey);
                    ptrHook = SetWindowsHookEx(13, objKeyboardProcess, GetModuleHandle(objCurrentModule.ModuleName), 0);


                    InitializeComponent();

                    //Process p = Process.GetCurrentProcess();
                    //int process_id = p.Id;
                    //ShowWindow(process_id, SW_HIDE);

                    DataTable dt = GetTable("Select * from Code_Conduct where username ='" + user + "'");
                    if (dt.Rows.Count > 0)
                    {
                        DataTable dt_table = GetTable("Select * from vw_hr_codeofconduct where username ='" + user + "'");

                        if (dt_table.Rows.Count > 0)
                        {

                        }
                        else
                        {
                            Environment.Exit(0);
                        }
                    }


                }

                catch (Exception ex)
                {
                    string LOG_FILE_NAME = "log_upload_bloomberg_" + System.DateTime.Now.Date.ToString().Substring(0, 10) + ".txt";
                    LOG_FILE_NAME = LOG_FILE_NAME.Replace("/", string.Empty);
                    StreamWriter log;

                    //Server.MapPath("~/MisReports/Investment/investment_bloomberg/MORNING.xlsx");

                    if (!File.Exists(@"C:\"+ LOG_FILE_NAME))
                    // if (!File.Exists(@"D:\investment_bloomberg\Logs\"+LOG_FILE_NAME))
                    {

                        //  log = new StreamWriter(Server.MapPath("~/MisReports/Investment/investment_bloomberg/Logs/") + LOG_FILE_NAME);
                        // log = new StreamWriter(@"D:\investment_bloomberg\Logs\"+LOG_FILE_NAME);

                        log = new StreamWriter(@"C:\"  + LOG_FILE_NAME);
                        // log = new StreamWriter(@"D:\investment_bloomberg\Logs\"+LOG_FILE_NAME);
                        log.Close();
                        log = File.AppendText(@"C:\" + LOG_FILE_NAME);

                    }
                    else
                    {
                        log = File.AppendText(@"C:\" + LOG_FILE_NAME);
                        // log = File.AppendText(@"D:\investment_bloomberg\Logs\" + LOG_FILE_NAME);
                    }

                    string cust_string =   System.DateTime.Now.Date.ToString() + "||" + ex.ToString() + "|" ;
                    log.WriteLine(cust_string);
                    // Close the stream:
                    log.Close();

                }
               
            }

            

            protected override CreateParams CreateParams
            {
                get
                {
                    var cp = base.CreateParams;
                    cp.ExStyle |= 0x80;  // Turn on WS_EX_TOOLWINDOW
                    return cp;
                }
            }

            bool HasAltModifier(int flags)
            {
                return (flags & 0x20) == 0x20;
            }
            private IntPtr captureKey(int nCode, IntPtr wp, IntPtr lp)
            {
                if (nCode >= 0)
                {
                    KBDLLHOOKSTRUCT objKeyInfo = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lp, typeof(KBDLLHOOKSTRUCT));

                    if (objKeyInfo.key == Keys.RWin || objKeyInfo.key == Keys.LWin || objKeyInfo.key == Keys.Tab && HasAltModifier(objKeyInfo.flags) || objKeyInfo.key == Keys.Escape && (ModifierKeys & Keys.Control) == Keys.Control) 
                    {
                        return (IntPtr)1;
                    }
                    }
                return CallNextHookEx(ptrHook, nCode, wp, lp);
            }

            private void Form2_FormClosing(object sender, FormClosingEventArgs e)
            {
                e.Cancel = true;
               

            }

          

            private void Form2_KeyPress(object sender, KeyPressEventArgs e)
            {
                //textBox1.Text += e.KeyChar.ToString();
            }

           

            private void pictureBox1_Click(object sender, EventArgs e)
            {

            }


            private void Form2_Load(object sender, EventArgs e)
            {
                // BlockInput(true);
                this.Location = new Point(0, 0);
                this.Size = Screen.PrimaryScreen.WorkingArea.Size;
                this.WindowState = FormWindowState.Maximized;
           

                button1.Location = new Point(0, 0);
                pictureBox1.Location = new Point(0, 0);
                panel1.Location = new Point(0, 0);
              //  hScrollBar1.Location = new Point(0, 360);
                this.ControlBox = false;
                this.ShowInTaskbar = false;

            }


            public void datainserting(string user, int status, string FullName, string depart)
            {
                // string connectionString = ConfigurationManager.ConnectionStrings["Test"].ConnectionString;

                String query = "sp_Code_Conduct";

                SqlCommand command = new SqlCommand(query, GetConnection());
                command.CommandType = CommandType.StoredProcedure;
                command.Parameters.AddWithValue("@username", user);
                command.Parameters.AddWithValue("@status", status);
                command.Parameters.AddWithValue("@displayName", FullName);
                command.Parameters.AddWithValue("@department", depart);
              
                command.ExecuteNonQuery();

            }

            public static DataTable GetTable(string query)
            {
                //string dbcon = ConfigurationManager.ConnectionStrings["Sebis"].ConnectionString;
                //SqlConnection con = new SqlConnection(dbcon);
                try
                {
                    DataTable dt = new DataTable();

                    //con.Open();
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, GetConnection());
                    adapter.Fill(dt);
                    // con.Close();
                    return dt;
                }
                catch (Exception)
                {
                    Environment.Exit(0);
                    throw;
                }
            }


            public static SqlConnection GetConnection()
            {

                string dbcon = "Data Source=1.1.1.1;Initial Catalog=abc;;User Id=user;Password=pword";
                try
                {

                    SqlConnection con = new SqlConnection(dbcon);
                    con.Close();
                    con.Open();
                    return con;

                }
                catch (Exception)
                {
                    Environment.Exit(0);
                    throw;
                }
            }

            private void button1_Click(object sender, EventArgs e)
            {
                try
               
                {
                    datainserting(user, 1, displayName, department);
                  
                    this.Dispose();
                    this.Close();
                    System.Windows.Forms.Application.Exit();
                    //InputBlocker.Block(10000);
                }

                catch (System.Exception excep)
                {

                    MessageBox.Show(excep.Message);

                }
            }

            private void pictureBox1_Click_1(object sender, EventArgs e)
            {

            }

         

           

         

        }
   
   
}
