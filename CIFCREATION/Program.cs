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
using System.Reflection;

namespace CIFCREATION
{
    class Program
    {
         public static OracleConnection connection;


        public int ExcelRows(excel.Range rng3)
        {
            int lastUsedRow = rng3.Cells.Find("*", System.Reflection.Missing.Value,
                                           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                           Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            return lastUsedRow;
        }

        public void OpenConnection()
        {
            try
            {
                connection = new OracleConnection(ConfigurationManager.ConnectionStrings["con"].ConnectionString);
                Console.WriteLine(connection);
                connection.Open();
            }
            catch (Exception e)
            {
                Console.WriteLine("Issue while Connecting Database : " + e);
            }

        }

        public void CloseConnection()
        {
            connection.Close();
        }
        static void Main(string[] args)
        {
            
            Program pr = new Program();
            pr.OpenConnection();
            string hp = "Select PARSING_METAXML_FILEPATH from TBLIMPFORMAT where FORMATNAME like '%Customer Import'";
            
            OracleCommand cmd1 = new OracleCommand(hp, connection);
            OracleDataReader rd1 = cmd1.ExecuteReader();
            rd1.Read();
            string path = rd1.GetString(0);
            Console.WriteLine(path);
            Thread.Sleep(2000);
            string hp1 = "select SRC_FILE_LOC from TBLIMPFILELOC where OPERATION_TYPE like '%Import%'";
            OracleCommand cmd2 = new OracleCommand(hp1, connection);
            OracleDataReader rd2 = cmd2.ExecuteReader();
            rd2.Read();
            string path1 = rd2.GetString(0);
            Console.WriteLine(path1);
            Thread.Sleep(2000);

            pr.CloseConnection();
            
            

            String datetime = DateTime.Now.ToString("yyyyMMddHHmmss");
            string fpath = $"{path1}CIF-FILE{datetime}.txt";
            Console.WriteLine("File Path: "+fpath);
            //Load Xml
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(path);
            XmlNode customerimport = xDoc.SelectSingleNode("CUSTOMERIMPORT");
            XmlNode header = customerimport.SelectSingleNode("HE");
            XmlNode card = customerimport.SelectSingleNode("CARD");
            XmlNode customer = customerimport.SelectSingleNode("CUSTOMER");
            XmlNode account = customerimport.SelectSingleNode("ACCOUNT");
            XmlNode footer = customerimport.SelectSingleNode("FO");
            //Load Excel
            excel.Application x1 = new excel.Application();
            string dirpath= Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
            string excelFileName = ConfigurationManager.AppSettings["ExcelFileName"];
            string excelpath = dirpath + "\\" + excelFileName;
            Console.WriteLine("Excel Path: "+excelpath);
            excel.Workbook wb = x1.Workbooks.Open(excelpath);
            excel._Worksheet sheet1 = wb.Sheets[1];
            excel.Range rng1 = sheet1.UsedRange;

            // making Header in CIFFILE
            Body bd = new Body();
            Headfoot.MakeHeader(fpath,header);
            //start working on create customer
            

            int rows =pr.ExcelRows(rng1);
            Console.WriteLine(rows);
            int i;
            for (i=2;i<=rows;i++) {

                bd.MakeCard(fpath,card,rng1,i,connection);
                bd.MakeCustomer(fpath, customer, rng1, i);
                bd.MakeAccount(fpath, account, rng1, i,connection);
            }

            // making footer of CIF file
            Headfoot.MakeFooter(fpath,footer,rows+1);
            Console.WriteLine("Running Auto Import Service 1");
            ServiceHandler.StartService("Auto Import Service 1", 60000);
            Thread.Sleep(1800000);
            Console.WriteLine("Stopping Auto Import Service 1");
            ServiceHandler.StopService("Auto Import Service 1", 60000);
            wb.Close();
            
            Console.Read();
            
        }
    }
}
