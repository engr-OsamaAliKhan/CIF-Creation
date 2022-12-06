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
    class FileMaker
    {
        public static void WriteFile(string Filename, string input)
        {
            FileStream fs = new FileStream(Filename, FileMode.Append, FileAccess.Write);
            if (fs.CanWrite)
            {
                byte[] buffer = Encoding.ASCII.GetBytes(input);
                fs.Write(buffer, 0, buffer.Length);
            }
            fs.Flush();
            fs.Close();

        }
    }
}
