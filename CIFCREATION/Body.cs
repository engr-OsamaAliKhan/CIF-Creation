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
    class Body
    {
        Program pro = new Program();
        public void MakeCard(string fpath, XmlNode card, excel.Range rng1, int rownum, OracleConnection conn)
        {
            int len;
            int xslen;
            string rn = string.Empty;
            string af = string.Empty;
            string cnic = string.Empty;
            string cardname = string.Empty;
            string customertype = string.Empty;
            string pcnic = string.Empty;
            string pid = string.Empty;
            string regentype = string.Empty;
            string dlbr = string.Empty;
            string dpid = string.Empty;
            string oldpan = string.Empty;
            string rc = string.Empty;
            string carddata = string.Empty;
            XmlNodeList cardnodes = card.ChildNodes;

            //to make dynamic data cell[column][row]

            foreach (XmlNode node in cardnodes)
            {

                if (node.Name == "RECORDNUMBER")
                {
                    rn = rng1.Cells[1][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    Console.WriteLine("Length of Record Number " + len);
                    xslen = rn.Length;
                    if (len != xslen)
                    {
                        Console.WriteLine("length not matched in recordnumber column in excel do you want to skip this data Press y to exit");
                        string y = Console.ReadLine();
                        if (y.Contains("y")) { break; }
                    }

                }
                if (node.Name == "RECORDCATEGORY")
                {
                    rc = node.Attributes[3].Value;
                    Console.WriteLine("Value of Record Type " + rc);
                }
                if (node.Name == "ACTIONFLAG")
                {
                    af = rng1.Cells[2][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    xslen = af.Length;
                    if (len != xslen)
                    {
                        Console.WriteLine("length not matched in recordnumber column in excel do you want to skip this data Press y to exit");
                        string y = Console.ReadLine();
                        if (y.Contains("y")) { break; }
                    }
                }
                if (node.Name == "CNIC")
                {
                    cnic = rng1.Cells[5][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    cnic = cnic.PadRight(len);
                }
                if (node.Name == "CARDNAME")
                {
                    cardname = rng1.Cells[6][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    cardname = cardname.PadRight(len);
                }
                if (node.Name == "CUSTOMERTYPE")
                {
                    customertype = rng1.Cells[7][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    xslen = customertype.Length;
                    if (len != xslen)
                    {
                        Console.WriteLine("length not matched in recordnumber column in excel do you want to skip this data Press y to exit");
                        string y = Console.ReadLine();
                        if (y.Contains("y")) { break; }
                    }
                }
                if (node.Name == "PRIMARY_CNIC")
                {
                    pcnic = rng1.Cells[8][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    pcnic = pcnic.PadRight(len);
                }

                if (node.Name == "PRODUCTID")
                {
                    pid = rng1.Cells[10][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);

                }
                if (node.Name == "REGENERATION_TYPE")
                {
                    regentype = rng1.Cells[11][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);

                }
                if (node.Name == "OLDPAN")
                {
                    regentype = rng1.Cells[11][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    if (regentype.Equals("00"))
                    {
                        oldpan = (oldpan ?? "").PadRight(len);
                    }
                    else
                    {
                        pro.OpenConnection();
                        cnic = rng1.Cells[5][rownum].Value;
                        len = int.Parse(node.Attributes[1].Value);
                        string hp = "Select MAX(CARDNUMBER) as OLDCARDNUMBER from TBLDEBITCARD where CUSTOMERID=(select CUSTOMERID from TBLCUSTOMER where CNIC='" + cnic + "')";
                        OracleCommand cmd1 = new OracleCommand(hp, conn);
                        OracleDataReader rd1 = cmd1.ExecuteReader();
                        rd1.Read();
                        oldpan = rd1.GetString(0);
                        oldpan = oldpan.PadRight(len);

                        pro.CloseConnection();

                    }

                }
                if (node.Name == "DELIVERY_BRANCH")
                {
                    dlbr = rng1.Cells[12][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);

                }


            }
            carddata = $"\n{rn}{rc}{af}{cnic}{cardname}{customertype}{pcnic}{oldpan}{pid}{regentype}{dlbr}";
            Console.WriteLine(carddata);
            FileMaker.WriteFile(fpath, carddata);



        }

        public void MakeCustomer(string fpath, XmlNode customer, excel.Range rng1, int rownum)
        {
            int len;
            int xslen;
            string rn = string.Empty;
            string rc = string.Empty;
            string af = string.Empty;
            string cnic = string.Empty;
            string title = string.Empty;
            string fullname = string.Empty;
            string dob = string.Empty;
            string mname = string.Empty;
            string paf = string.Empty;
            string haddress1 = string.Empty;
            string haddress2 = string.Empty;
            string haddress3 = string.Empty;
            string haddress4 = string.Empty;
            string hpcode = string.Empty;
            string hphone = string.Empty;
            string email = string.Empty;
            string reserve = string.Empty;
            string company = string.Empty;
            string officeaddress1 = string.Empty;
            string officeaddress2 = string.Empty;
            string officeaddress3 = string.Empty;
            string officeaddress4 = string.Empty;
            string officeaddress5 = string.Empty;
            string opc = string.Empty;
            string officephone = string.Empty;
            string mnumber = string.Empty;
            string bf = string.Empty;
            string adt = string.Empty;
            string nationality = string.Empty;
            string fname = string.Empty;
            string dlm = string.Empty;
            string reserve3 = string.Empty;
            string reserve4 = string.Empty;
            string consumer = string.Empty;
            string pno = string.Empty;
            string cust = string.Empty;


            XmlNodeList customernodes = customer.ChildNodes;

            foreach (XmlNode node in customernodes)
            {

                if (node.Name == "RECORDNUMBER")
                {
                    rn = rng1.Cells[1][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    Console.WriteLine("Length of Record Number " + len);
                    xslen = rn.Length;
                    if (len != xslen)
                    {
                        Console.WriteLine("length not matched in recordnumber column in excel do you want to skip this data Press y to exit");
                        string y = Console.ReadLine();
                        if (y.Contains("y")) { break; }
                    }

                }
                if (node.Name == "RECORDCATEGORY")
                {
                    rc = node.Attributes[3].Value;
                    Console.WriteLine("Value of Record Type " + rc);
                }
                if (node.Name == "ACTIONFLAG")
                {
                    af = rng1.Cells[3][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    xslen = af.Length;
                    if (len != xslen)
                    {
                        Console.WriteLine("length not matched in recordnumber column in excel do you want to skip this data Press y to exit");
                        string y = Console.ReadLine();
                        if (y.Contains("y")) { break; }
                    }
                }
                if (node.Name == "CNIC")
                {
                    cnic = rng1.Cells[5][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    cnic = cnic.PadRight(len);
                }
                if (node.Name == "TITLE")
                {
                    title = rng1.Cells[14][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    title = title.PadRight(len);
                }

                if (node.Name == "FULLNAME")
                {
                    fullname = rng1.Cells[15][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    fullname = fullname.PadRight(len);
                }
                if (node.Name == "DATEOFBIRTH")
                {
                    dob = rng1.Cells[16][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    dob = dob.PadRight(len);
                }
                if (node.Name == "MOTHERSNAME")
                {
                    mname = rng1.Cells[17][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    mname = mname.PadRight(len);
                }
                if (node.Name == "PREFERRED_ADDRESS_FLAG")
                {
                    paf = rng1.Cells[18][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    paf = paf.PadRight(len);
                }
                if (node.Name == "HOMEADDRESS1")
                {
                    haddress1 = rng1.Cells[19][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    haddress1 = haddress1.PadRight(len);
                }
                if (node.Name == "HOMEADDRESS2")
                {
                    haddress2 = rng1.Cells[20][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    haddress2 = haddress2.PadRight(len);
                }
                if (node.Name == "HOMEADDRESS3")
                {
                    haddress3 = rng1.Cells[21][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    haddress3 = haddress3.PadRight(len);
                }
                if (node.Name == "HOMEADDRESS4")
                {
                    haddress4 = rng1.Cells[22][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    haddress4 = haddress4.PadRight(len);
                }
                if (node.Name == "HOMEPOSTALCODE")
                {
                    hpcode = rng1.Cells[23][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    hpcode = hpcode.PadRight(len);
                }
                if (node.Name == "HOMEPHONE")
                {
                    hphone = rng1.Cells[24][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    hphone = hphone.PadRight(len);
                }
                if (node.Name == "EMAIL")
                {
                    email = rng1.Cells[25][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    email = email.PadRight(len);
                }
                if (node.Name == "RESERVED")
                {

                    len = int.Parse(node.Attributes[1].Value);
                    reserve = (reserve ?? "").PadRight(len);
                }
                if (node.Name == "COMPANY")
                {
                    company = rng1.Cells[26][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    company = company.PadRight(len);
                }
                if (node.Name == "OFFICEADDRESS1")
                {
                    officeaddress1 = rng1.Cells[27][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    officeaddress1 = officeaddress1.PadRight(len);
                }
                if (node.Name == "OFFICEADDRESS2")
                {
                    officeaddress2 = rng1.Cells[28][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    officeaddress2 = officeaddress2.PadRight(len);
                }
                if (node.Name == "OFFICEADDRESS3")
                {
                    officeaddress3 = rng1.Cells[29][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    officeaddress3 = officeaddress3.PadRight(len);
                }
                if (node.Name == "OFFICEADDRESS4")
                {
                    officeaddress4 = rng1.Cells[30][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    officeaddress4 = officeaddress4.PadRight(len);
                }
                if (node.Name == "OFFICEADDRESS5")
                {
                    officeaddress5 = rng1.Cells[31][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    officeaddress5 = officeaddress5.PadRight(len);
                }
                if (node.Name == "OFFICEPOSTALCODE")
                {
                    opc = rng1.Cells[32][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    opc = opc.PadRight(len);
                }
                if (node.Name == "OFFICEPHONE")
                {
                    officephone = rng1.Cells[33][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    officephone = officephone.PadRight(len);
                }
                if (node.Name == "MOBILENUMBER")
                {
                    mnumber = rng1.Cells[34][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    mnumber = mnumber.PadRight(len);
                }

                if (node.Name == "CONSUMERRNUMBER")
                {
                    len = int.Parse(node.Attributes[1].Value);
                    consumer = (consumer ?? "").PadRight(len);
                }
                if (node.Name == "BILLINGFLAG")
                {
                    bf = rng1.Cells[35][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    bf = bf.PadRight(len);
                }
                if (node.Name == "ANNIVERSARYDATE")
                {
                    adt = rng1.Cells[36][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);

                }
                if (node.Name == "PASSPORT_NO")
                {
                    len = int.Parse(node.Attributes[1].Value);
                    pno = (pno ?? "").PadRight(len);
                }
                if (node.Name == "NATIONALITY")
                {
                    nationality = rng1.Cells[37][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    nationality = nationality.PadRight(len);
                }
                if (node.Name == "RESERVED3")
                {
                    len = int.Parse(node.Attributes[1].Value);
                    reserve3 = (reserve3 ?? "").PadRight(len);
                }
                if (node.Name == "RESERVED4")
                {
                    len = int.Parse(node.Attributes[1].Value);
                    reserve4 = (reserve4 ?? "").PadRight(len);
                }
                if (node.Name == "FATHERSNAME")
                {
                    fname = rng1.Cells[38][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    fname = fname.PadRight(len);
                }
                if (node.Name == "DELIMITER")
                {
                    dlm = node.Attributes[3].Value;
                }



                cust = $"\n{rn}{rc}{af}{cnic}{title}{fullname}{dob}{mname}{paf}{haddress1}{haddress2}{haddress3}{haddress4}{hpcode}{hphone}{email}{reserve}{company}{officeaddress1}{officeaddress2}{officeaddress3}{officeaddress4}{officeaddress5}{opc}{officephone}{mnumber}{consumer}{bf}{adt}{pno}{nationality}{reserve3}{reserve4}{fname}{dlm}";



            }

            Console.WriteLine(cust);
            FileMaker.WriteFile(fpath, cust);

        }

        public void MakeAccount(string fpath, XmlNode account, excel.Range rng1, int rownum, OracleConnection conn)
        {

            int len;
            int xslen;
            string rn = string.Empty;
            string rc = string.Empty;
            string af = string.Empty;
            string cnic = string.Empty;
            string atitle = string.Empty;
            string dlbr = string.Empty;
            string acctid = string.Empty;
            string accttype = string.Empty;
            string acctcurrency = string.Empty;
            string status = string.Empty;
            string bimd = string.Empty;
            string deft = string.Empty;
            string cnum = string.Empty;
            string acct = string.Empty;
            string regentype = string.Empty;

            XmlNodeList anodes = account.ChildNodes;
            foreach (XmlNode node in anodes)
            {
                if (node.Name == "RECORDNUMBER")
                {
                    rn = rng1.Cells[1][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    Console.WriteLine("Length of Record Number " + len);
                    xslen = rn.Length;
                    if (len != xslen)
                    {
                        Console.WriteLine("length not matched in recordnumber column in excel do you want to skip this data Press y to exit");
                        string y = Console.ReadLine();
                        if (y.Contains("y")) { break; }
                    }

                }
                if (node.Name == "RECORDCATEGORY")
                {
                    rc = node.Attributes[3].Value;
                    Console.WriteLine("Value of Record Type " + rc);
                }
                if (node.Name == "ACTIONFLAG")
                {
                    af = rng1.Cells[4][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    xslen = af.Length;
                    if (len != xslen)
                    {
                        Console.WriteLine("length not matched in recordnumber column in excel do you want to skip this data Press y to exit");
                        string y = Console.ReadLine();
                        if (y.Contains("y")) { break; }
                    }
                }
                if (node.Name == "CNIC")
                {
                    cnic = rng1.Cells[5][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    cnic = cnic.PadRight(len);
                }
                if (node.Name == "TITLE")
                {
                    atitle = rng1.Cells[45][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    atitle = atitle.PadRight(len);
                }
                if (node.Name == "BRANCHID")
                {
                    dlbr = rng1.Cells[12][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);

                }
                if (node.Name == "ACCOUNT_ID")
                {
                    acctid = rng1.Cells[39][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    acctid = acctid.PadRight(len);
                }
                if (node.Name == "ACCOUNT_TYPE")
                {
                    accttype = rng1.Cells[40][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);

                }
                if (node.Name == "ACCOUNT_CURRENCY")
                {
                    acctcurrency = rng1.Cells[41][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);

                }
                if (node.Name == "STATUS")
                {
                    status = rng1.Cells[42][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);

                }
                if (node.Name == "BANKIMD")
                {
                    bimd = rng1.Cells[43][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);

                }
                if (node.Name == "IS_DEFAULT")
                {
                    deft = rng1.Cells[44][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);

                }
                if (node.Name == "CARDNUMBER")
                {
                    regentype = rng1.Cells[11][rownum].Value;
                    len = int.Parse(node.Attributes[1].Value);
                    if (regentype.Equals("00") || regentype.Equals("01"))
                    {
                        cnum = (cnum ?? "").PadRight(len);
                    }
                    else
                    {
                        pro.OpenConnection();
                        cnic = rng1.Cells[5][rownum].Value;
                        len = int.Parse(node.Attributes[1].Value);
                        string hp = "Select MAX(CARDNUMBER) as OLDCARDNUMBER from TBLDEBITCARD where CUSTOMERID=(select CUSTOMERID from TBLCUSTOMER where CNIC='" + cnic + "')";
                        OracleCommand cmd1 = new OracleCommand(hp, conn);
                        OracleDataReader rd1 = cmd1.ExecuteReader();
                        rd1.Read();
                        cnum = rd1.GetString(0);
                        cnum = cnum.PadRight(len);

                        pro.CloseConnection();
                    }
                }
                
            }
            acct = $"\n{rn}{rc}{af}{cnic}{cnum}{acctid}{accttype}{acctcurrency}{status}{atitle}{bimd}{dlbr}{deft}";
            Console.WriteLine(acct);
            FileMaker.WriteFile(fpath, acct);
        }
    }
}
