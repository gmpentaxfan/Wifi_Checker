using System;
using System.Collections.Generic;
//GM V1.00 July 8 2016 using System.Linq;
//GM V1.00 July 8 2016 using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics; //GM V1.20 July 12 2016


namespace Wifi_Checker
{
    
    static class Program

    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
         {
            //Allows only one instance to run
            //http://stackoverflow.com/questions/6486195/ensuring-only-one-application-instance
            //GM V1.00 July 8 2016
            bool result;
            var mutex = new System.Threading.Mutex(true, "Wifi Checker", out result);
            //MessageBox.Show("Test Kill Process"); //GM V1.20 July 12 2016

            if (!result)
            {
                MessageBox.Show("Restarting Program.", "Company WiFi Checker", MessageBoxButtons.OK, MessageBoxIcon.Information);//GM V1.20 July 12 2016
                //Kill other process
                int procID = Process.GetCurrentProcess().Id; //GM V1.20 July 12 2016
                try
                {
                    foreach (Process name in Process.GetProcessesByName("WiFi Checker"))
                    {
                        string ProcessName = Convert.ToString(name.ProcessName);
                        int otherprocID = name.Id;
                        if (otherprocID != procID)
                        {
                            name.Kill();
                        }

                    }
                    //MessageBox.Show(Convert.ToString(Process.GetProcessesByName("WiFi Checker")) + "\n" + "Try");
                }
                catch
                {
                    //MessageBox.Show(Convert.ToString(Process.GetProcessesByName("WiFi Checker")) + "\n" + "Catch");
                }
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());
            Application.Run(new WifiApplicationContext());
            
        }
    }
}
