using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net; //GM V1.00 July 8 2016
using System.Windows.Forms; //GM V1.00 July 8 2016
using System.Net.NetworkInformation; //GM V1.00 July 8 2016
using System.IO; //GM V1.10 July 11 2016
using System.Diagnostics; //GM V1.10 July 11 2016
using System.Threading; //GM V1.10 July 11 2016
using System.Management; //GM V1.10 July 11 2016
using System.DirectoryServices.AccountManagement; //GM V1.30 July 13 2016
using System.DirectoryServices; //GM V1.30 July 13 2016

namespace Wifi_Checker
{
    public class WifiApplicationContext : ApplicationContext
    {
        static NotifyIcon notifyIcon = new NotifyIcon();
        static string CompanyIP = "False"; 
        static string WifiStatus = "Not Found";
        static string EthernetStatus = "Not Found"; //GM V1.10 July 11 2016
        static string NoCompanymessage = "No Company network was detected." + "\n" +"Please verify your network connections." +"\n\n"+ "Click 'Yes' to restart the program." +"\n\n" + "Click 'No' to continue." + "\n\n" + "Click 'Cancel' to hide this message."; //GM V1.20 July 12 2016
        static string NoWirelessmessage = "No Company Wireless network connection detected." + "\n" + "Do you want to start Diagnostics?" + "\n\n" + "Click 'Cancel' to hide this message." ; //GM V1.20 July 12 2016
        static string Desktop = "True"; //GM V1.10 July 11 2016
        static string BatteryStatus = "Not Found"; //GM V1.10 July 11 2016
        static string CancelStatus = "False"; //GM V1.20 July 12 2016
        static string StartLoginScript = "False"; //GM V1.30 July 13 2016
        

        static MenuItem EthernetMenuItem = new MenuItem("Check Ethernet", new EventHandler(ClickEthernet)); //GM V1.10 July 11 2016
        static MenuItem WiFiMenuItem = new MenuItem("Check WiFi", new EventHandler(ClickWiFi)); //GM V1.10 July 11 2016
        static MenuItem DiagMenuItem = new MenuItem("Start Diagnostics", new EventHandler(ClickDiag)); //GM V1.10 July 11 2016
        static MenuItem HelpItem = new MenuItem("Help", new EventHandler(ViewPDF)); //GM V1.30 July 13 2016

        public WifiApplicationContext()
        {
            DesktopCheck();
            notifyIcon.Icon = Wifi_Checker.Properties.Resources.icon_bad;
            notifyIcon.Click += new EventHandler(KeyPressClick);
            notifyIcon.Visible = true;

            //GM V1.10 Stop if Desktop
            //MessageBox.Show(Desktop + "--" + BatteryStatus);

            if (Desktop.Equals("False"))
            {
                Startup();
                if (WifiStatus.Equals("True"))
                {
                    notifyIcon.ContextMenu = new ContextMenu(new MenuItem[] { EthernetMenuItem, WiFiMenuItem, HelpItem }); //GM V1.10 July 11 2016
                    ShowMessages();
                }
                else
                {
                    notifyIcon.ContextMenu = new ContextMenu(new MenuItem[] { DiagMenuItem, EthernetMenuItem, WiFiMenuItem, HelpItem }); //GM V1.10 July 11 2016
                    ShowMessages();
                }
            }
            if (Desktop.Equals("True") && BatteryStatus.Equals("Not Found"))
              {
                  notifyIcon.Visible = false;
                  //MessageBox.Show("This does not appear to be a mobile device.", "Company Wireless Checker", MessageBoxButtons.OK, MessageBoxIcon.Error);
                  Environment.Exit(0);
              }
            
        }


        void KeyPressClick(object sender, EventArgs e)
        {
            //GM V1.00 July 5 2016
            if (Control.ModifierKeys == Keys.Control)
            {
                //MessageBox.Show("Exiting Program");
                notifyIcon.Visible = false;
                Application.Exit();
            }

            if (Control.ModifierKeys == Keys.LShiftKey || Control.ModifierKeys == Keys.RShiftKey)
            {
                MessageBox.Show("Shift Key");
            }

            if (Control.ModifierKeys == Keys.Alt)
            {
                MessageBox.Show("Manually starting diagnostics.", "Company Wireless Checker", MessageBoxButtons.OK, MessageBoxIcon.Information);  //GM V1.30 July 13 2016
                StartDiagnostics();
                if (StartLoginScript.Equals("True"))
                {
                    LoginScript();
                }

            }
        }

        static void Startup()
        {
            //Changed from Application to Startup Function GM V1.10 July 2016
            //Check network before starting //GM V1.00 July 8 2016
            CompanyIPCheck(); //V1.20 July 12 2016

            if (CompanyIP.Equals("False"))
            {
                MessageBox.Show("We did not detect a Company network connection.", "Company Wireless Checker", MessageBoxButtons.OK, MessageBoxIcon.Error);
                StartDiagnostics();
            }
            else
            {
                VerifyWifi();
            }

            notifyIcon.BalloonTipTitle = "Company Wifi Checker";
            Thread.Sleep(3000); //GM V1.20 July 12 2016
            ShowMessages();

            if (StartLoginScript.Equals("True"))
            {
                LoginScript();
            }

        }

        static void DesktopCheck()
        {
            //GM V1.10 July 11 2016
            //http://www.gkspk.com/view/programming/working-with-windows-services-using-csharp-and-wmi/
            //Check Chassis
            try
            {
                ManagementClass systemEnclosures = new ManagementClass("Win32_SystemEnclosure");
                foreach (ManagementObject obj in systemEnclosures.GetInstances())
                {
                    foreach (int i in (UInt16[])(obj["ChassisTypes"]))
                    {
                        if (i == 1 || i == 2 || i == 8 || i == 9 || i == 10 || i == 11 || i == 12 || i == 14 || i == 18 || i == 17 || i == 21)
                        {
                            Desktop = "False";
                        }
                    }
                }
                //MessageBox.Show(Desktop);
                if (Desktop.Equals("False"))
                {
                    System.Management.ObjectQuery query = new ObjectQuery("Select * FROM Win32_Battery");
                    ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);

                    ManagementObjectCollection collection = searcher.Get();

                    foreach (ManagementObject mo in collection)
                    {
                        foreach (PropertyData property in mo.Properties)
                        {
                            string Name = property.Name;
                            if (Name.Equals("Status"))
                            {
                                BatteryStatus = "Found";
                            }
                            Console.WriteLine("Property {0}: Value is {1}", property.Name, property.Value);
                        }
                    }
                }
            }
            catch
            {
                //Continue
            }
        }

        static void CompanyIPCheck() //GM V1.20 July 12 2016
        {
            IPHostEntry IPHostEntry = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName());
            foreach (System.Net.IPAddress IPAddress in IPHostEntry.AddressList)
            {
                // InterNetwork indicates that an IP version 4 address is expected 
                // when a Socket connects to an endpoint
                if (IPAddress.AddressFamily.ToString() == "InterNetwork")
                {
                    string IPText = IPAddress.ToString();
                    if (IPText.StartsWith("192.168."))
                    {
                        CompanyIP = "True";
                    }
                }
            }
        }

        static void ShowMessages()  //GM V1.20 July 12 2016
        {
            if (WifiStatus.Equals("Not Found"))
            {
                notifyIcon.BalloonTipText = "No WiFi adapter found.";
                notifyIcon.Text = "Company Wifi Checker - No WiFi adapter found.";
            }

            if (WifiStatus.Equals("True"))
            {
                notifyIcon.BalloonTipText = "Enterprise WiFi Connected.";
                notifyIcon.Text = "Company Wifi Checker - Enterprise WiFi Connected.";
            }

            if (WifiStatus.Equals("Down"))
            {
                notifyIcon.BalloonTipText = "Not connected to a wireless network.";
                notifyIcon.Text = "Company Wifi Checker - Right Click for more options.";
                QuestionWifi();
            }

            if (WifiStatus.Equals("False"))
            {
                notifyIcon.BalloonTipText = "Enterprise WiFi NOT Connected.";
                notifyIcon.Text = "Company Wifi Checker - Right Click for more options.";
                QuestionWifi();
            }

            notifyIcon.ShowBalloonTip(1);
        }

        static void ClickEthernet(object sender, EventArgs e)
        {
            CancelStatus = "False"; //GM V1.20 July 12 2016
            VerifyEthernet();
            MessageEthernet();
        }

        static void ClickWiFi(object sender, EventArgs e)
        {
            CancelStatus = "False"; //GM V1.20 July 12 2016
            VerifyWifi();
            Thread.Sleep(3000); //GM V1.20 July 12 2016
            
            if (WifiStatus.Equals("True"))
            {
                notifyIcon.ContextMenu = new ContextMenu(new MenuItem[] { EthernetMenuItem, WiFiMenuItem }); //GM V1.10 July 11 2016
                notifyIcon.Icon = Wifi_Checker.Properties.Resources.icon_good;
                MessageWifi();
            }
            else
            {
                notifyIcon.ContextMenu = new ContextMenu(new MenuItem[] { DiagMenuItem, EthernetMenuItem, WiFiMenuItem }); //GM V1.10 July 11 2016
                notifyIcon.Icon = Wifi_Checker.Properties.Resources.icon_bad;
                StartDiagnostics();
            }
        }
        
        static void ClickDiag(object sender, EventArgs e)
        {
            CancelStatus = "False"; //GM V1.20 July 12 2016
            StartDiagnostics();
            MessageWifi();
        }

        static void QuestionWifi()
        {
            if (CancelStatus.Equals("True")) //GM V1.20 July 12 2016
            {
                return;
            }
            else
            {
                //GM V1.10 July 11 2016
                var DiagAnswer = MessageBox.Show(NoWirelessmessage, "Company Wireless Checker - Warning", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question); //GM V1.20 July 12 2016
                if (DiagAnswer == DialogResult.Yes)
                {
                    StartDiagnostics();
                }
                if (DiagAnswer == DialogResult.No)
                {
                    ConnectCompany();
                    Thread.Sleep(3000); //GM V1.20 July 12 2016
                    VerifyWifi();
                }
                if (DiagAnswer == DialogResult.Cancel) //GM V1.20 July 12 2016
                {
                    CancelStatus = "True";
                }
            }
            
        }

        static void MessageEthernet()
        {
            //GM V1.10 July 11 2016
            if (EthernetStatus.Equals("True"))
            {
                MessageBox.Show("Valid ethernet connection found.", "Company Wireless Checker", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (EthernetStatus.Equals("False"))
            {
                MessageBox.Show("A valid Company IP was not detected.", "Company Wireless Checker", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (EthernetStatus.Equals("Not Found"))
            {
                MessageBox.Show("No ethernet adapter was found.", "Company Wireless Checker", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        static void MessageWifi()
        {
            //GM V1.10 July 11 2016
            if (WifiStatus.Equals("True"))
            {
                MessageBox.Show("Valid WiFi connection found.", "Company Wireless Checker", MessageBoxButtons.OK, MessageBoxIcon.Information);
                notifyIcon.ContextMenu = new ContextMenu(new MenuItem[] { EthernetMenuItem, WiFiMenuItem }); //GM V1.10 July 11 2016
            }
            if (WifiStatus.Equals("False"))
            {
                MessageBox.Show("A valid Company wireless IP was not detected.", "Company Wireless Checker", MessageBoxButtons.OK, MessageBoxIcon.Error);
                notifyIcon.ContextMenu = new ContextMenu(new MenuItem[] { DiagMenuItem, EthernetMenuItem, WiFiMenuItem }); //GM V1.10 July 11 2016
            }
            if (WifiStatus.Equals("Down"))
            {
                MessageBox.Show("The wireless adapter is not connected to a wireless network.", "Company Wireless Checker", MessageBoxButtons.OK, MessageBoxIcon.Error);
                notifyIcon.ContextMenu = new ContextMenu(new MenuItem[] { DiagMenuItem, EthernetMenuItem, WiFiMenuItem }); //GM V1.10 July 11 2016
            }
            if (WifiStatus.Equals("Not Found"))
            {
                MessageBox.Show("No wireless adapter was found.", "Company Wireless Checker", MessageBoxButtons.OK, MessageBoxIcon.Error);
                notifyIcon.ContextMenu = new ContextMenu(new MenuItem[] { DiagMenuItem, EthernetMenuItem, WiFiMenuItem }); //GM V1.10 July 11 2016
            }

            else
            {
                
            }
        }
        
        static void StartDiagnostics()
        {
            //Something;
            //GM V1.10 July 11 2016
            StartLoginScript = "False"; //GM V1.30 July 13 2016
            VerifyEthernet();
            if(EthernetStatus.Equals("True"))
            {
                UpdateGroupPolicy();
                Thread.Sleep(30);
                ConnectCompany();
                Thread.Sleep(3000); //GM V1.20 July 12 2016
                VerifyWifi();
            }
            else
            {
                var answer = MessageBox.Show(NoCompanymessage, "Company Wireless Checker - Warning", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question); //GM V1.20 July 12 2016
                if (answer == DialogResult.Yes)
                {
                    Application.Restart();
                }
                if (answer == DialogResult.No)
                {
                    ConnectCompany();
                    Thread.Sleep(3000); //GM V1.20 July 12 2016
                    VerifyWifi();
                }
                if (answer == DialogResult.Cancel)
                {
                    CancelStatus = "True";
                }
                
            }

            //Update Taskbar Icon Display
            notifyIcon.BalloonTipTitle = "Company Wifi Checker";

            if (WifiStatus.Equals("Not Found"))
            {
                notifyIcon.BalloonTipText = "No WiFi adapter found.";
                notifyIcon.Text = "Company Wifi Checker - No WiFi adapter found.";
                notifyIcon.Icon = Wifi_Checker.Properties.Resources.icon_bad;
            }

            if (WifiStatus.Equals("True"))
            {
                notifyIcon.BalloonTipText = "Enterprise WiFi Connected.";
                notifyIcon.Text = "Company Wifi Checker - Enterprise WiFi Connected.";
                notifyIcon.Icon = Wifi_Checker.Properties.Resources.icon_good;
                StartLoginScript = "True"; //GM V1.30 July 13 2016
            }

            if (WifiStatus.Equals("False") || WifiStatus.Equals("Down")) //GM V1.10 July 11 2016
            {
                notifyIcon.BalloonTipText = "Enterprise WiFi NOT Connected.";
                notifyIcon.Text = "Company Wifi Checker - Right Click for more options.";
                notifyIcon.Icon = Wifi_Checker.Properties.Resources.icon_bad;
            }

            notifyIcon.ShowBalloonTip(1);
        }

        static void UpdateGroupPolicy()
        {
            //GM V1.10 July 11 2016
            //http://stackoverflow.com/questions/18195203/how-to-programatically-update-group-policy-with-c-i-e-gpupdate-force
            MessageBox.Show("Updating Computer Policy please wait...", "Company Wireless Checker", MessageBoxButtons.OK, MessageBoxIcon.Information);
            FileInfo execFile = new FileInfo("gpupdate.exe");
            Process proc = new Process();
            proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            proc.StartInfo.FileName = execFile.Name;
            proc.StartInfo.Arguments = "/force /target:Computer /wait:60";
            proc.Start();
            //Wait for GPUpdate to finish
            proc.WaitForExit();
            MessageBox.Show("Update procedure has finished.", "Company Wireless Checker", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        static void VerifyEthernet()
        {
            //GM V1.10 July 11 2016
            //Check for network before updating policy
            CompanyIPCheck(); //GM V1.20 July 2016

            NetworkInterface[] adapters = NetworkInterface.GetAllNetworkInterfaces();
            foreach (NetworkInterface n in adapters)
            {
                //Console.WriteLine("   {0} is {1}", n.Name, n.OperationalStatus);
                //Console.WriteLine("   {0} is {1} -- {2} -- {3}", n.Name, n.OperationalStatus, n.Id, n.NetworkInterfaceType);
                //MessageBox.Show(Convert.ToString(n.Name) + "--" + Convert.ToString(n.OperationalStatus) + "--" + Convert.ToString(n.Id) + "--" + Convert.ToString(n.NetworkInterfaceType));
                string strInterfaceType = Convert.ToString(n.NetworkInterfaceType);
                string strInterfaceName = Convert.ToString(n.Name); //GM V1.10 July 11 2016
                string strOpeartionalSatus = Convert.ToString(n.OperationalStatus); //GM V1.10 July 11 2016
                
                if (strInterfaceName.StartsWith("Local") && strInterfaceType.StartsWith("Ethernet") && strOpeartionalSatus.Equals("Up") && CompanyIP.Equals("True"))
                {
                    EthernetStatus = "True";
                    break;
                }
                if (strInterfaceName.StartsWith("Local") && strInterfaceType.StartsWith("Ethernet") && strOpeartionalSatus.Equals("Down") && CompanyIP.Equals("False"))
                {
                    EthernetStatus = "False";
                    break;
                }
                if (strInterfaceName.StartsWith("Local") && strInterfaceType.StartsWith("Ethernet") && strOpeartionalSatus.Equals("Down"))
                {
                    EthernetStatus = "False";
                    break;
                }
            }
        }

        static void VerifyWifi()
        {
            CompanyIPCheck(); //GM V1.20 July 12 2016
            NetworkInterface[] adapters = NetworkInterface.GetAllNetworkInterfaces();
            foreach (NetworkInterface n in adapters)
            {
                //Console.WriteLine("   {0} is {1}", n.Name, n.OperationalStatus);
                //Console.WriteLine("   {0} is {1} -- {2} -- {3}", n.Name, n.OperationalStatus, n.Id, n.NetworkInterfaceType);
                //MessageBox.Show(Convert.ToString(n.Name) + "--" + Convert.ToString(n.OperationalStatus) + "--" + Convert.ToString(n.Id) + "--" + Convert.ToString(n.NetworkInterfaceType));
                string strInterfaceType = Convert.ToString(n.NetworkInterfaceType);
                string strOpeartionalSatus = Convert.ToString(n.OperationalStatus); //GM V1.10 July 11 2016

                if(strInterfaceType.StartsWith("Wireless") && CompanyIP.Equals("True") && strOpeartionalSatus.Equals("Up"))
                {
                    WifiStatus = "True";
                    notifyIcon.Icon = Wifi_Checker.Properties.Resources.icon_good;
                    break;
                }
                if (strInterfaceType.StartsWith("Wireless") && CompanyIP.Equals("False") && strOpeartionalSatus.Equals("Down"))
                {
                    WifiStatus = "False";
                    break;
                }
                if (strInterfaceType.StartsWith("Wireless") && strOpeartionalSatus.Equals("Down"))
                {
                    WifiStatus = "Down";
                    break;
                }

            }
        }

        static void ConnectCompany()
        {
            //GM V1.10 July 11 2016
            FileInfo execFile = new FileInfo("netsh.exe");
            Process proc = new Process();
            proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            proc.StartInfo.FileName = execFile.Name;
            proc.StartInfo.Arguments = "wlan connect name=\"Company-enterprise\"";
            proc.Start();
            //Wait for GPUpdate to finish
            proc.WaitForExit();
            MessageBox.Show("Attempting to connect to Company Wireless...", "Company Wireless Checker", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        static void LoginScript()
        {
            //GM V1.30 July 13 2016
            string scriptfailure = "False";
            string homefailure = "False";
            string errormessage = "Pass";
            string strhomedrive = "Not Found";
            string strhomedirectory = "Not Found";
                    
            //Connect home folder - works but causes issues with subsequent logins.
            //try
            //{
                //UserPrincipal User = UserPrincipal.Current;
                //DirectoryEntry ldapConnection = new DirectoryEntry("company.ca");
                //ldapConnection.Path = "LDAP://OU=Company Users,DC=company,DC=ca";
                //ldapConnection.AuthenticationType = AuthenticationTypes.Secure;
                
                //DirectorySearcher search = new DirectorySearcher(ldapConnection);
                //search.Filter = "(cn=" + User + ")";

                //string[] requiredProperties = new string[] { "homeDirectory", "homeDrive" };
                //foreach (String property in requiredProperties) 
                    //search.PropertiesToLoad.Add(property);

                //SearchResult result = search.FindOne();
            
                //if (result != null)
                //{
                    //foreach (String property in requiredProperties)
                        //foreach (Object myCollection in result.Properties[property]) 
                        //{
                            //MessageBox.Show(String.Format("{0,-20} : {1}", property, myCollection.ToString()));
                            //if (property.Contains("homeDirectory"))
                            //{
                            //    strhomedirectory = Convert.ToString(myCollection.ToString());
                            //}
                            //if (property.Contains("homeDrive"))
                            //{
                            //    strhomedrive = Convert.ToString(myCollection.ToString());
                            //}
                        //}
                //}
            //}
            //catch
            //{
                //homefailure = "True";
            //}

            if(strhomedirectory.ToString().Equals("Not Found") && strhomedrive.ToString().Equals("Not Found"))
            {
                homefailure = "True";
            }
            //else
            //{
                //try
                //{
                    //FileInfo execFile = new FileInfo("net.exe");
                    //Process proc = new Process();
                    //proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    //proc.StartInfo.FileName = execFile.Name;
                    //proc.StartInfo.Arguments = "use " + strhomedrive + " " + strhomedirectory;
                    //proc.Start();
                    //Wait for Home folder to finish
                    //proc.WaitForExit();
                //}
                //catch
                //{
                    //homefailure = "True";
                //}
            //}

            MessageBox.Show("Attempting to connect network drives and printers.", "Company Wireless Checker", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //Run login script
            try
            {

                FileInfo execFile = new FileInfo("wscript.exe");
                Process proc = new Process();
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                proc.StartInfo.FileName = execFile.Name;
                proc.StartInfo.Arguments = "\\\\company.ca\\NETLOGON\\CompanyLogin.vbs";
                proc.Start();
                //Wait for Script to finish
                proc.WaitForExit();
            }
            catch
            {
                scriptfailure = "True";
            }                

            if (scriptfailure.Equals("True"))
            {
                errormessage = "Login script was not found." + "\n";
            }

            if (homefailure.Equals("True"))
            {
                errormessage = "Your 'J' drive may not be connected." + "\n";
            }

            //MessageBox.Show(scriptfailure + "\n" + homefailure + "\n" + errormessage + "\n" + strhomedrive + "\n" + strhomedirectory);
            
            if (errormessage.Equals("Pass"))
                MessageBox.Show("If there are any issues, you may have to log off and log on again.", "Company Wireless Checker", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show(errormessage + "\n You may have to log off and log on again.", "Company Wireless Checker", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        static void ViewPDF(object sender, EventArgs e)
        {
            //GM V1.30 July 11 2016
            //MessageBox.Show("Test Two", "Company Wireless Checker", MessageBoxButtons.OK, MessageBoxIcon.Information);
            String openPDFFile = @"help.pdf";
            System.IO.File.WriteAllBytes(openPDFFile, Wifi_Checker.Properties.Resources.Help);
            System.Diagnostics.Process.Start(openPDFFile);      
        }
    }
}
