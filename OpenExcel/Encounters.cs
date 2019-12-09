using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;

namespace OpenExcel
{

    

    class Encounters
    {
        Excel.Worksheet _sheet;

        public Encounters(string path) {
            var excelApp = new Excel.Application();
            excelApp.Visible = true;

            Excel.Workbooks books = excelApp.Workbooks;
            Excel.Workbook wb = books.Open(path);
            _sheet = wb.ActiveSheet;
        }

        public void analyzePatients(int year, int month) {
            
            for (int i = 1; i < 32; i ++) {
                string arrivalDate = month.ToString() + "/" + i.ToString() + "/" + year.ToString();
                findInWorksheet(arrivalDate);
            }
        
        }

        public int getPatientsPerDay(DateTime date) {
            Excel.Range range = null;
            Excel.Range currentFind = null;
            Excel.Range firstFind = null;
            int count = 0;
            range = _sheet.Columns[Columns.ADT_ARRIVAL_DATE];
            string dateString = date.ToString("M/d/yyyy");
            firstFind = range.Find(dateString, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, Type.Missing, Type.Missing, Type.Missing);

            if (firstFind != null) {

                count++;
                displayFindResult(firstFind);
                currentFind = range.FindNext(firstFind);
                if (currentFind != null) {
                          count++;
                    displayFindResult(currentFind);
                }
                while ((currentFind != null) & (currentFind.Address != firstFind.Address))  {
                    currentFind = range.FindNext(currentFind);
                    if (currentFind.Address != firstFind.Address) {
                        count++;
                        displayFindResult(currentFind);
                    }
                }
            }
            Console.WriteLine(dateString + ", " + count);


            return count;
        }

        /*void findInWorksheet(Excel.Range range, string searchTerm)
        {
            Excel.Range currentFind = null;
            Excel.Range firstFind = null;
            int count = 0;

            ArrayList readingTimes = new ArrayList();


            //the problem is that the date is matched in columns B, C, and AG, so we get three matches for every patient arrival. 
            //we need to restrict the range to column B
            range.va


            range = (Excel.Range)_sheet.Cells[2, 35000];

            //
            firstFind = range.Find(searchTerm, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, Type.Missing,Type.Missing, Type.Missing); 
            if (firstFind != null)
            {
                count++;
                //string firstFindString = (string)firstFind.Cells[firstFind.Row, Columns.ADT_ARRIVAL_TIME].Value;
                Console.WriteLine(firstFind.Address);
                //what was the read time for this event? It's in column Q
                //readingTimes.Add(sheetRange.Cells[firstFind.Row, "Q"].Value);
                currentFind = range.FindNext(firstFind);
                while ((currentFind != null) & (currentFind.Address != firstFind.Address))
                {
                    currentFind = range.FindNext(currentFind);
                    if (currentFind.Address != firstFind.Address)
                    {
                        //readingTimes.Add(sheetRange.Cells[currentFind.Row, "Q"].Value);
                        count++;
                        Console.WriteLine(currentFind.Address);
                    }
                }
            }
            Console.WriteLine("Number of patients on " + searchTerm + ": " + count);
            
            double sum = 0;
            int count = 0;
            foreach (Object time in readingTimes)
            {
                count++;
                sum += (double)time;
            }
            Console.WriteLine(sum / count);
            
        }*/

        public void findInWorksheet(string searchTerm) {
            Excel.Range range = null;
            Excel.Range currentFind = null;
            Excel.Range firstFind = null;
            int count = 0;

            ArrayList readingTimes = new ArrayList();

            //restrict range to our desired search column
            range = _sheet.Columns[Columns.ADT_ARRIVAL_DATE];

            firstFind = range.Find(searchTerm, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, Type.Missing, Type.Missing, Type.Missing);


            if (firstFind != null) {
                
                if (_sheet.Cells[firstFind.Row, Columns.ATTND_PROV_NAME].Value == Physicians.ruiz) {
                    count++;
                    displayFindResult(firstFind);
                }
                

                currentFind = range.FindNext(firstFind);
                if (_sheet.Cells[currentFind.Row, Columns.ATTND_PROV_NAME].Value == Physicians.ruiz) {
                    count++;
                    displayFindResult(currentFind);
                }
                    

                if (currentFind != null) {
                    if (_sheet.Cells[currentFind.Row, Columns.ATTND_PROV_NAME].Value == Physicians.ruiz) {
                        count++;
                        displayFindResult(currentFind);
                    }
                        
                }
                while ((currentFind != null) & (currentFind.Address != firstFind.Address)) {

                    currentFind = range.FindNext(currentFind);
                    if (currentFind.Address != firstFind.Address)  {
                        if ((string)_sheet.Cells[currentFind.Row, Columns.ATTND_PROV_NAME].Value == Physicians.ruiz) {
                            count++;
                            displayFindResult(currentFind);
                        }

                    }
                }
            }
            Console.WriteLine("Total, " + count);
     
        }

        private void displayFindResult(Excel.Range range)   {
            Console.Write(_sheet.Cells[range.Row, Columns.ADT_ARRIVAL_TIME].Value + ", ");
            if (_sheet.Cells[range.Row, Columns.ATTND_PROV_NAME].Value != null)  {
                Console.WriteLine(_sheet.Cells[range.Row, Columns.ATTND_PROV_NAME].Value);
            }            
        }
    }
}
