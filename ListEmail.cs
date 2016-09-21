
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;




namespace Newsletter
{
    public class ListEmail
    {
        public void ListingEmail()
        {
            //call a path 
            string path = "C:\\Users\\Desktop\\Desktop\\Pasta1.xlsx";
            
            //Create a application excel
            Excel.Application excel = new Excel.Application();
            //Open the Workbooks
            Excel.Workbook wb = excel.Workbooks.Open(path);
            //Open the Sheet
            Excel.Worksheet excelSheet = wb.ActiveSheet;

            
            //Create String test, int linha and SendEmail send
            string test = " ";
            int linha = 1;
            SendEmail send = new SendEmail();


            while (test != null)
            {
                               
                try
                {
                    //Read the first cell
                    //[linha,coluna]
                    test = excelSheet.Cells[linha, 3].Value.ToString();
                   // Call SendEmail.SendingEmail()
                    send.SendingEmail(test);
                   
                }
                catch (Exception)
                {

                   break;
                }
                //Show Sended Email                
                Console.WriteLine(test);
                linha++;    
            }
            Console.ReadKey();
            wb.Close();

        }
    }
}
