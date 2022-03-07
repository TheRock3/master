using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using Bytescout.Spreadsheet;
using System.Data.SqlClient;
using System.IO;
using System.Data;


namespace HelloWorld
{
    class Program
    {
        static void Main(string[] args)
        {
            //string filePath = @"C:\Users\rakeshyad\Desktop\Apps\Payroll\Documentation\Book4.xlsx";
            //string filePath = @"C:\Users\rakeshyad\Desktop\Apps\Excise\PRD_FINAL\";
            string filePath = @"C:\Users\rakeshyad\Desktop\Apps\Sales\table_mig\";
            //string filePath = @"C: \Users\rakeshyad\Desktop\Apps\MotorcycleRefundBooking\table_mig\";
            //Excel.Application xlApp = new Excel.Application(); 
            // Create new Spreadsheet

            // Write message
            Console.Write("Press any key to continue...");
            DirSearch_ex3(filePath, "");
            // Wait user input
            Console.ReadKey();

        }

        static void ReadWrite(string name,string filePath)
        {
            try
            {
                Spreadsheet document = new Spreadsheet();
                document.LoadFromFile(filePath);

                // Get worksheet by name
                Worksheet worksheet = document.Workbook.Worksheets.ByName("Sheet 1");

                // Check dates
                for (int i = 1; i < worksheet.NotEmptyRowMax; i++)
                {
                    // Set current cell
                    Cell tableName = worksheet.Cell(i, 0);
                    //string Sql = "select count(*) from ";
                    string Sql = "select EMPL_CODE,CO from H_LEAVE order by empl_code ";
                    string conn = "";
                    using (SqlConnection con = new SqlConnection(conn))
                    {
                        try
                        {
                            SqlDataAdapter da = new SqlDataAdapter();
                            DataSet ds = new DataSet();
                            SqlCommand cmd = new SqlCommand(Sql, con);
                            con.Open();
                            da.Fill(ds);
                            int rowsCount = Convert.ToInt32(cmd.ExecuteScalar());
                            worksheet.Cell(i, 2).Value = rowsCount.ToString();
                        }
                        catch (Exception ex)
                        {

                        }
                        
                    }
                }
                if(filePath.Contains("04.csv"))
                {
                    filePath = filePath.Replace("04.csv", "Apr.csv");
                }
                else if (filePath.Contains("05.csv"))
                {
                    filePath = filePath.Replace("05.csv", "May.csv");
                }
                else if (filePath.Contains("06.csv"))
                {
                    filePath = filePath.Replace("06.csv", "Jun.csv");
                }
                filePath = filePath.Replace(".csv", "_count.csv");
                document.SaveAs(filePath);
                // Close document
                document.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }


        static void DirSearch_ex3(string sDir, string sb1)
        {



            try
            {


               

                string[] files = Directory.GetFiles(sDir);
                foreach (string f in files)
                {
                    if (f.EndsWith("table_count.csv"))
                    {
                        int lineNo = 1;
                        using (StreamReader file = new StreamReader(f))
                        {
                            string ln;
                            StringBuilder sbuild = new StringBuilder();
                            //string name = f.Split(new string[] { "\\" }, StringSplitOptions.None)[10].Split('.')[0];
                            ReadWrite("", f);
                            //string[] appName = { "AVS", "KISOK", "HERO_WEB_BRA", "Payroll" };
                            /*using (StreamWriter w = File.AppendText(@"C:\Users\rakeshyad\Desktop\LookUp\LookUpFinal.txt"))
                            {

                                w.WriteLine();
                                // w.WriteLine("---------------------DirName-------------");
                                w.WriteLine();
                                w.WriteLine("---------------------fileName-------------");
                                w.WriteLine(f);
                                w.WriteLine();
                                w.WriteLine();
                                w.WriteLine();
                                Boolean isTrue = false;

                                while ((ln = file.ReadLine()) != null)
                                {
                                    //if (ln.ToUpper().Contains("PUBLIC") && (ln.ToUpper().Contains("DATATABLE") || ln.ToUpper().Contains("INT")))
                                    //{

                                    //    w.WriteLine();
                                    //    w.WriteLine();
                                    //    w.WriteLine("---------------------MethodName-------------");
                                    //    w.WriteLine();
                                    //    w.WriteLine();
                                    //    w.WriteLine(ln.ToString());
                                    //}
                                    bool found = false;
                                    foreach (string sub in subjects)
                                    {
                                        if (string.IsNullOrEmpty(ln.ToString()))
                                        {
                                            break;
                                        }

                                        if (ln.ToUpper().Contains(sub.ToUpper()) &&
                                            !ln.ToUpper().Contains("#REGION"))
                                        //    (keys[keys.Length - 1].Length==0||keys[keys.Length - 1][0]!='_'||
                                        //Char.IsLetter(keys[keys.Length - 1][0]) == false))
                                        //{
                                        //    foreach (char ch in sub.ToUpper().Trim())
                                        //    {
                                        //        int calLength = 0;
                                        //        string line = ln.ToUpper().Trim().ToString();
                                        //        foreach(char chr in line)
                                        //        {
                                        //            if (chr.Equals(ch))
                                        //            {
                                        //                calLength += 1;
                                        //            }
                                        //            else
                                        //            {
                                        //                calLength = 0;
                                        //            }
                                        //        }
                                        //        found = calLength == sub.Length ? true : false;
                                        //    }
                                        //}
                                        //ln.IndexOf(sub)>0&&
                                        //ln.IndexOf(sub)<ln.Length&&
                                        //ln[ln.IndexOf(sub) - 1].ToString().Equals(" ")&&
                                        //ln[ln.IndexOf(sub) - 1].ToString().Equals(" "))
                                        {

                                            // Console.WriteLine(ln);

                                            //if (ln.ToUpper().Contains("select REPT_DESC from M_TOOL_REPORT_MSTR".ToUpper())||ln.ToUpper().Contains("Select_Report()"))
                                            //{
                                            //    Console.WriteLine(ln);
                                            //}
                                            //w.Write(lineNo);
                                            if (ln.Contains(" entities."))
                                            {
                                                string[] entString = ln.Split(new string[] { " entities." }, StringSplitOptions.None);
                                                int entKey = 1;
                                                string entValue;
                                                while (entKey < entString.Length)
                                                {
                                                    entValue = entString[entKey++].Split()[0];
                                                    w.WriteLine("[Table]");
                                                    if (!entValue.Contains('('))
                                                    {
                                                        w.WriteLine(entValue);
                                                    }
                                                }
                                                w.WriteLine();
                                                w.WriteLine();

                                                //w.WriteLine(ln.Split(new string[] { " entities." }, StringSplitOptions.None));
                                            }
                                            w.WriteLine(sub);
                                            w.WriteLine();
                                            w.WriteLine();
                                            isTrue = true;
                                            //break;
                                        }
                                    }
                                    lineNo += 1;
                                }
                                //if (isTrue)
                                //{
                                //    w.WriteLine("---------------------fileName-------------");
                                //    w.WriteLine(f);
                                //    w.WriteLine();
                                //    w.WriteLine();
                                //    w.WriteLine();
                                //    w.WriteLine();

                                //}

                            }
                            file.Close();*/
                        }
                    }

                }

                foreach (string d in Directory.GetDirectories(sDir))
                {
                    DirSearch_ex3(d,"");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.ReadLine();
            }
            //File.WriteAllText(@"C:\users\rakeshyad\desktop\" + "Table3" + '-' + DateTime.Now.ToString("dd-MMM-yyyy") + ".xls", sb.ToString());
            //Console.WriteLine();
        }
    }
}
