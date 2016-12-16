using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;                // for files?
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace GRUT_1
{
    public partial class Form1 : Form
    {
        String baseExcelFilename = null;
        String baseExcelFileType = null;
        String rosterFileName = null;


        const String EXCEL = ".xls";
        const String TEXT = ".txt";

        public Form1()
        {
            InitializeComponent();
        }

        private void getArgosFile()
        {
//            MessageBox.Show(openFileDialog1.Filter);
            openFileDialog1.Filter = "Argos file|*.xls;*.txt";
            if(openFileDialog1.ShowDialog()==DialogResult.OK)
            {
                baseExcelFilename = openFileDialog1.FileName;
                baseExcelFileType = Path.GetExtension(baseExcelFilename);
                MessageBox.Show("The argos file is of type " + baseExcelFileType);
            }
        }
        private void createRosterFileName()
        {
            if (baseExcelFilename == null)
            {
                MessageBox.Show("No Argos Report File Selected");
                return;
            }
            rosterFileName = baseExcelFilename;
            rosterFileName = System.IO.Path.GetFileName(rosterFileName);
            rosterFileName = rosterFileName.Replace(baseExcelFileType, "_Roster.txt");
            MessageBox.Show("The rosterFileName is " + rosterFileName);
            folderBrowserDialog1.ShowDialog();
            String directory = folderBrowserDialog1.SelectedPath;
            rosterFileName = directory + "\\" + rosterFileName;
            MessageBox.Show("The rosterFileName is now " + rosterFileName);

        }
        private void createRoster()
        {

        }
        private void createRosterFromArgosExcel()
        {
            if (baseExcelFilename == null)
            {
                MessageBox.Show("OOPS, no argos file selected");
            }
            else if (baseExcelFileType == EXCEL)
            {
                ExcelApplication bob = new ExcelApplication(baseExcelFilename);
 //                   MessageBox.Show(bob.word);
 //                   bob.word = baseExcelFileType;
 //                   MessageBox.Show(bob.word);
                bob.createRoster(/*lstViewer,*/ rosterFileName);
            }
            else if (baseExcelFileType == TEXT)
            {
                MessageBox.Show("Processing argos text file ");
                createRosterFromTextFile();
            }
            else
                MessageBox.Show("OOPS, there be DRAGONS here ");
        }

        private void createRosterFromTextFile()
        {
            StreamReader argosTextFile = new StreamReader(baseExcelFilename);
            List<String> data = new List<string>();
            string line;
            while ( (line=argosTextFile.ReadLine()) != null )
            {
                line = line.Trim();
                if (line.StartsWith("N0"))
                {
                    string number = line.Substring( 0, 9 );
                    //                    line = line.Substring(line.IndexOf(' '));
                    //                    line = line.Trim();
                    int lastNameLength = line.IndexOf(',') - 13;
                    String lastName = line.Substring(13,lastNameLength);
                    //                    line = line.Substring(line.IndexOf(' '));
                    //                    line = line.Trim();
                    int firstNameLength = 42 - line.IndexOf(',');
                    String firstName = line.Substring(line.IndexOf(',')+1, firstNameLength);
                    firstName = firstName.Trim();
                    //                    line = line.Substring(line.IndexOf(' '));
                    //                    line = line.Trim();
                    String userID = line.Substring(95);
                    userID = userID.Substring(0, userID.IndexOf('@'));
                    String it = number + ":" + lastName + ":" + firstName + ":" + userID;
//                    lstViewer.Items.Add(it);
                    data.Add(it);
                }
                
            }
            using (StreamWriter outputFile = new StreamWriter(rosterFileName))
            {
                foreach (String thing in data)
                    outputFile.WriteLine(thing);
                outputFile.Close();
            }
            
            argosTextFile.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            MessageBox.Show("Closing");
            
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void runProcess()
        {
            System.IO.Directory.CreateDirectory(@"D:\Spring_2017\Test");
        }

        private void testCopy()
        {
            getArgosFile();
            ExcelApplication bob = new ExcelApplication(baseExcelFilename);
            bob.makeStudentCopies(lstViewer);
        }

        private void btnTryIt_Click(object sender, EventArgs e)
        {
            //            getArgosFile();
            //            createRosterFileName();
            //            createRoster();
            // runProcess();
            testCopy();
                      

           

            MessageBox.Show("FINISHED TryIt");

        }
    }
}
