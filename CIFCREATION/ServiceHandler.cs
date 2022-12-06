using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using excel = Microsoft.Office.Interop.Excel;
using System.ServiceProcess;
using System.Threading;
using System.Configuration;
using Oracle.ManagedDataAccess.Client;

namespace CIFCREATION
{
    class ServiceHandler
    {
        public static void StartService(string serviceName, int timeoutMilliseconds)
        {
            ServiceController myService = new ServiceController();
            myService.ServiceName = serviceName;
            try
            {
                TimeSpan timeout = TimeSpan.FromMilliseconds(timeoutMilliseconds);

                myService.Start();
                myService.WaitForStatus(ServiceControllerStatus.Running, timeout);
            }
            catch (Exception e)
            {
                Console.WriteLine("To run service please run this program in administrative mode " + e);
            }
        }

        public static void StopService(string serviceName, int timeoutMilliseconds)
        {
            ServiceController myService = new ServiceController();
            myService.ServiceName = serviceName;
            try
            {
                TimeSpan timeout = TimeSpan.FromMilliseconds(timeoutMilliseconds);

                myService.Stop();
                myService.WaitForStatus(ServiceControllerStatus.Stopped, timeout);
            }
            catch (Exception e)
            {
                
                Console.WriteLine("To run service please run this program in administrative mode " + e.Message);
            }

        }
    }
}
