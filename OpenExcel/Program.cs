using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;


namespace OpenExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"C:\Users\Y118428\Desktop\ShiftAnalysisRawData.xlsx";
            //string path = @"C:\Users\Y118428\Desktop\JulyLookback.xlsx";
            //string path = @"C:\Users\Y118428\Desktop\11pm_to_6am_calculations2.xlsm";
            Encounters encounters = new Encounters(path);
            /*for (int i = 1; i < 9; i ++)
            {
                encounters.analyzePatients(2019, i);
            }*/
            /*
            DateTime date = new DateTime(2019, 1, 1);
            DateTime endDate = new DateTime(2019, 9, 1);
            while (date < endDate)
            {
                encounters.getPatientsPerDay(date);
                date = date.AddDays(1);
            }*/
            encounters.getPatientsPerDay(new DateTime(2019, 8, 12));


            Console.Read();
            ;
        }
    }
}