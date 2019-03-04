using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcessMonitor
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                bool flg = false;
                Process[] processList = Process.GetProcesses();
                foreach (Process p in processList)
                {
                    if (p.ProcessName == "EXCEL")
                    {
                        flg = true;
                        if (p.Responding)
                        {
                            
                            Console.WriteLine("OK");
                            break;
                        }
                        else
                        {
                            p.CloseMainWindow();
                            //RunApp();
                        }
                    }
                }
                if (!flg)
                {
                    //RunApp();
                }
                System.Threading.Thread.Sleep(5000);
            }
        }
    }
}
