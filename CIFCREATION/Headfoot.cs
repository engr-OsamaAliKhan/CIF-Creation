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
    class Headfoot
    {
        public static void MakeHeader(string fpath, XmlNode header)
        {

            // making Header in Cif File 
            XmlNodeList hnodes = header.ChildNodes;
            string rn = string.Empty;
            string hvalue = string.Empty;
            string dt = string.Empty;
            string filename = string.Empty;
            string padname = string.Empty;
            string vrvalue = string.Empty;
            string headr = string.Empty;

            foreach (XmlNode node in hnodes)
            {

                if (node.Name == "RECORDNUMBER")
                {
                    int rnlen = int.Parse(node.Attributes[1].Value);
                    Console.WriteLine("Length of Record Number " + rnlen);
                    rn = "000001";

                }
                if (node.Name == "RECORDCATEGORY")
                {
                    hvalue = node.Attributes[3].Value;
                    Console.WriteLine("Value of Record Type " + hvalue);
                }
                if (node.Name == "RECORDDATE")
                {
                    int rdlen = int.Parse(node.Attributes[1].Value);
                    Console.WriteLine("Length of RecordDate " + rdlen);
                    dt = DateTime.Now.ToString("yyyyMMdd");
                    Console.WriteLine("Date is " + dt);

                }
                if (node.Name == "FILENAME")
                {
                    int fnlen = int.Parse(node.Attributes[1].Value);
                    Console.WriteLine("Length of Filename " + fnlen);
                    filename = "FILENAME";
                    padname = filename.PadRight(fnlen);
                }
                if (node.Name == "VERSION")
                {
                    vrvalue = node.Attributes[3].Value;
                    Console.WriteLine("Value of Version " + vrvalue);
                }

                headr = $"{rn}{hvalue}{dt}{padname}{vrvalue}";
            }
            Console.WriteLine(headr);
            FileMaker.WriteFile(fpath, headr);

        }

        public static void MakeFooter(string fpath, XmlNode footer, int rn)
        {
            XmlNodeList fnodes = footer.ChildNodes;
            string dlvalue = string.Empty;
            string ftr = string.Empty;
            string hvalue = string.Empty;
            string dt = string.Empty;
            string ftrn = rn.ToString();


            string filename = string.Empty;
            string padname = string.Empty;
            foreach (XmlNode node in fnodes)
            {

                if (node.Name == "RECORDNUMBER")
                {
                    int rnlen = int.Parse(node.Attributes[1].Value);
                    Console.WriteLine("Length of Record Number " + rnlen);
                    ftrn = ftrn.PadLeft(rnlen, '0');
                }
                if (node.Name == "RECORDCATEGORY")
                {
                    hvalue = node.Attributes[3].Value;
                    Console.WriteLine("Value of Record Type " + hvalue);
                }
                if (node.Name == "RECORDDATE")
                {
                    int rdlen = int.Parse(node.Attributes[1].Value);
                    Console.WriteLine("Length of RecordDate " + rdlen);
                    dt = DateTime.Now.ToString("yyyyMMdd");
                    Console.WriteLine("Date is " + dt);

                }
                if (node.Name == "FILENAME")
                {
                    int fnlen = int.Parse(node.Attributes[1].Value);
                    Console.WriteLine("Length of Filename " + fnlen);
                    filename = "FILENAME";
                    padname = filename.PadRight(fnlen);
                }
                if (node.Name == "DELIMITER")
                {
                    dlvalue = node.Attributes[3].Value;
                    Console.WriteLine("Value of Version " + dlvalue);
                }
                ftr = $"\n{ftrn}{hvalue}{dt}{padname}{dlvalue}";
            }

            Console.WriteLine(ftr);
            FileMaker.WriteFile(fpath, ftr);

        }
    }
}
