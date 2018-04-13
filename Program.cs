using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

namespace ExcelManager
{
    public class Read_From_Excel
    {
      
        static void Main(string[] args)
        
            //public static void getExcelFile()
        {
            string name = "SIP/Signtech-SIPTrunk-COMM/07797737223,300,Tt";
            int index = name.IndexOf("-");
            if (index > 0)
                name = name.Substring(0, index);
            Console.WriteLine(name);
            Console.WriteLine(name[0]);

            if (name[0].ToString() == "S")
            {
                name = name.Substring(name.IndexOf("/") + 1);
            }

            Console.WriteLine(name);








            
           // string file = System.Reflection.Assembly.GetExecutingAssembly().Location;
           // string file2 = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
           // file2 = file2.Replace("ExcelManager.exe","Invoice");
           // file2 = file2.Replace("file:///","");

           // Console.WriteLine(file2);

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\brendan\Desktop\AlexTest\Copy of excel2");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            Excel._Worksheet xlWorksheet2 = xlWorkbook.Sheets[2];
            Excel.Range xlRange2 = xlWorksheet2.UsedRange;
            Excel._Worksheet xlWorksheet3 = xlWorkbook.Sheets[3];
            Excel.Range xlRange3 = xlWorksheet3.UsedRange;
            Excel._Worksheet xlWorksheet4 = xlWorkbook.Sheets[4];
            Excel.Range xlRange4 = xlWorksheet4.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            int rowCount2 = xlRange2.Rows.Count;
            int colCount2 = xlRange2.Columns.Count;
            int rowCount3 = xlRange3.Rows.Count;
            int colCount3 = xlRange3.Columns.Count;
            int rowCount4 = xlRange4.Rows.Count;
            int colCount4 = xlRange4.Columns.Count;

            Console.WriteLine(rowCount);
            Console.WriteLine(colCount);
            Console.WriteLine(rowCount2);
            Console.WriteLine(colCount2);
            Console.WriteLine(rowCount3);
            Console.WriteLine(colCount3);
            Console.WriteLine(rowCount4);
            Console.WriteLine(colCount4);




            double totalAdd = 0;
            double totalTax = 0;
            double grandTotal = 0;
            bool last = false;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!


            //for (int i = 1; i <= rowCount; i++)
           


            for (int i = 1; i <= rowCount2; i++)

            {
                int j = 14;
                //for (int j = 14; j <= 14; j++)
                //for (int j = 1; j <= 1; j++)
                {
                    //new line
                    //if (j == 1)

                    //Console.Write("\r\n");

                    //write the value to the console
                    
                   
                    
                    if (xlRange2.Cells[i, j] != null && xlRange2.Cells[i, j].Value2 == null)
                    {
                        name = xlRange2.Cells[i, 9].Value2;
                        if (name[0].ToString() == "S")
                        {

                            index = name.IndexOf("-");
                            if (index > 0)
                                name = name.Substring(0, index);

                            name = name.Substring(name.IndexOf("/") + 1);

                            if (name != "Nitel")
                            {
                                for (int x = 1; x <= 49; x++)
                                {
                                    int y = 8;
                                    Console.WriteLine(name);
                                    string write = xlRange3.Cells[x, y].Value2;
                                    Console.WriteLine(write);
                                    if (xlRange3.Cells[x, y].Value2 == name)
                                    {
                                        xlRange2.Cells[i, 14].Value2 = xlRange3.Cells[x, 9].Value2;
                                    }
                                }
                            }
                        }

                                   
                    }
                   
                    //Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");

                }
            }
        /*
            for (int i = 2; i <= rowCount; i++)

            //for (int i = 1; i <= 1; i++)

            {
                                  
                        if (xlRange.Cells[i, 17] != null && xlRange.Cells[i, 17].Value2 != null)
                        {
          
                        double quantity = xlRange.Cells[i, 17].Value2;
                        double amount = xlRange.Cells[i, 18].Value2;
                        double total = quantity * amount;
                        xlRange.Cells[i, 19].Value2 = total;
                }
                 
                      
            }
            for (int i = 2; i <= rowCount; i++)

            //for (int i = 1; i <= 1; i++)

            {
                if (last == false)
                {
                    totalAdd = totalAdd + xlRange.Cells[i, 19].Value2;
                    totalTax = totalTax + xlRange.Cells[i, 22].Value2;
                }
                if (xlRange.Cells[i, 10].Value2 != xlRange.Cells[i + 1, 10].Value2)
                {
                    last = true;
                }
                if (xlRange.Cells[i, 10] != null && xlRange.Cells[i, 10].Value2 != null)
                {
                    invoiceNum = xlRange.Cells[i, 10].Value2;
                }
                //if(invoiceNum == invoiceNum2 || invoiceNum2 == "start")
            
                
                if (xlRange.Cells[i, 10] != null && xlRange.Cells[i, 10].Value2 != null)
                {
                    invoiceNum2 = xlRange.Cells[i, 10].Value2;
                }
                if (last == true)
                {
                    xlRange.Cells[i, 20].Value2 = totalAdd;
                    xlRange.Cells[i, 23].Value2 = totalTax;
                    grandTotal = totalAdd + totalTax;
                    xlRange.Cells[i, 24].Value2 = grandTotal;

                    totalTax = 0;
                    totalAdd = 0;
                    grandTotal = 0;
                    last = false;
                }
                    
                    
            }
            */
       
        


            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            Console.WriteLine("Done");
            Console.ReadLine();
            Console.ReadKey();
            
        }
    }
}
