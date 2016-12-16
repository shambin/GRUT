using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;

namespace GRUT_1
{
    /*  Things to do with excel:
     *  
     *  1.  Create a roster as a text file from Argos report
     *  
     *  2.  Copy master gradesheet for each student
     *      Need standardized gradesheet
     *      or
     *      Embed format
     *      
     *  3.  Create master gradesheet for a class
     *      From xml?
     *       
    */


    public class ExcelApplication
    {
        Excel.Application mainApp = null;
        Excel.Workbooks mainAppWorkbooks = null;
        Excel.Workbook argosWorkbook = null;
        Excel.Sheets argosWorksheets = null;
        Excel.Worksheet aSheet = null;
        Excel.Range aRange = null;
        int argosRowCount = 0;
        int argosColumnCount = 0;
        string someFile = null;

        public ExcelApplication(String argosFile)
        {
            someFile = argosFile;
            helloWorld = "Hello World";
            mainApp = new Excel.Application();
            mainAppWorkbooks = mainApp.Workbooks;
            // opening workbook causes leak
                       argosWorkbook = mainAppWorkbooks.Open(argosFile, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                       argosWorksheets = argosWorkbook.Worksheets;
                       aSheet = argosWorksheets.get_Item(1);
                       aRange = aSheet.UsedRange;
                       argosRowCount = aRange.Rows.Count;
                       argosColumnCount = aRange.Columns.Count;
        }

        public void makeStudentCopies( ListBox viewer)
        {
            
//            for ( int i=1; i<=argosRowCount; i++ )
            {
                string line = null;
                for ( int j=1; j<=argosColumnCount; j++ )
                {
                    line += aSheet.Cells[/*i*/ 1, j].value2 + "  ";
                }
                viewer.Items.Add(line);
            }

        }

        public void createRoster(/*ListBox viewer,*/ String rosterFileName = null)
        {
            if ( rosterFileName == null)
            {
                MessageBox.Show("Trying to create roster but file name is null ");
                return;
            }
            MessageBox.Show("Creating roster from argos excel file " + rosterFileName );
            List<String> data = new List<string>();
            for (int row = 1; row <= argosRowCount; row++)
                if (aRange.Cells[row, 1].value2 != null && aRange.Cells[row, 1].value2.StartsWith("N0"))
                {
                    String nNumber = aRange.Cells[row, 1].value2;
                    String name = aRange.Cells[row, 2].value2;
                    int startIndex = 0;
                    int endIndex = name.LastIndexOf(",");
                    int length = endIndex - startIndex + 1;
                    String lastName = name.Substring(startIndex, endIndex);
                    String firstName = name.Substring(endIndex + 2);

                    String email = aRange.Cells[row, 9].value2;
                    endIndex = email.LastIndexOf("@");
                    String userID = email.Substring(startIndex, endIndex);
                    String line = nNumber + ":" + lastName + ":" + firstName + ":" + userID;
                    data.Add(line);
//                    viewer.Items.Add(line);                    
                }
            using (StreamWriter outputFile = new StreamWriter(rosterFileName))
            {
                foreach (String line in data)
                    outputFile.WriteLine(line);
            }

        }

        ~ExcelApplication()
        {
            if (aRange != null)
            {
                MessageBox.Show("returning range");
                Marshal.ReleaseComObject(aRange);
            }
            if (aSheet != null)
            {
                MessageBox.Show("returning sheet");
                Marshal.ReleaseComObject(aSheet);
            }
            if (argosWorksheets != null)
            {
                MessageBox.Show("Returning sheets");
                Marshal.ReleaseComObject(argosWorksheets);
            }
            if (argosWorkbook != null)
            {
                MessageBox.Show("returning workbook");
                argosWorkbook.Close(false, null, null);
                Marshal.ReleaseComObject(argosWorkbook);
            }
            if (mainAppWorkbooks != null)
            {
                MessageBox.Show("returning workbooks");
                mainAppWorkbooks.Close();
                Marshal.ReleaseComObject(mainAppWorkbooks);
            }
            if (mainApp != null)
            {
                MessageBox.Show("Returning main application");
                mainApp.Quit();
                Marshal.ReleaseComObject(mainApp);
            }
            MessageBox.Show("ExcelApplication DESTRUCTOR");

        }

        String helloWorld = null;
        public string word
        {
            get { return helloWorld; }
            set { helloWorld = value; }
        }

    }// class ExcelApplication

}// namespace GRUT_1
