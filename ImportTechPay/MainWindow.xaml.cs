/* Title:           Import Tech Pay
 * Date:            5-*22-19
 * Author:          Terry Holmes
 * 
 * Description:     This is used to import tech pay */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using DataValidationDLL;
using NewEventLogDLL;
using TechPayDLL;

namespace ImportTechPay
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        DataValidationClass TheDataValidationClass = new DataValidationClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        TechPayClass TheTechPayClass = new TechPayClass();

        //setting up the data
        ImportedTechPayDataSet TheImportedTechPayDataSet = new ImportedTechPayDataSet();
        FindTechPayItemByCodeDataSet TheFindTechPayItemByCodeDataSet = new FindTechPayItemByCodeDataSet();
        FindTechPayItemByDescriptionDataSet TheFindTechPayItemByDescriptionDataSet = new FindTechPayItemByDescriptionDataSet();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnImportExcel_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlDropOrder;
            Excel.Workbook xlDropBook;
            Excel.Worksheet xlDropSheet;
            Excel.Range range;

            string strTechPayCode;
            string strJobDescription;
            string strValueForValidation;
            decimal decTechPayPrice = 0;
            bool blnFatalError = false;
            int intCounter;
            int intNumberOfRecords;
            int intRecordsReturned;
            int intColumnRange;
            bool blnItemFound;

            try
            {
                TheImportedTechPayDataSet.importedtechpay.Rows.Clear();

                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.FileName = "Document"; // Default file name
                dlg.DefaultExt = ".xlsx"; // Default file extension
                dlg.Filter = "Excel (.xlsx)|*.xlsx"; // Filter files by extension

                // Show open file dialog box
                Nullable<bool> result = dlg.ShowDialog();

                // Process open file dialog box results
                if (result == true)
                {
                    // Open document
                    string filename = dlg.FileName;
                }

                PleaseWait PleaseWait = new PleaseWait();
                PleaseWait.Show();

                xlDropOrder = new Excel.Application();
                xlDropBook = xlDropOrder.Workbooks.Open(dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlDropSheet = (Excel.Worksheet)xlDropOrder.Worksheets.get_Item(1);

                range = xlDropSheet.UsedRange;
                intNumberOfRecords = range.Rows.Count;
                intColumnRange = range.Columns.Count;

                for (intCounter = 2; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strTechPayCode = Convert.ToString((range.Cells[intCounter, 1] as Excel.Range).Value2).ToUpper();
                    strJobDescription = Convert.ToString((range.Cells[intCounter, 2] as Excel.Range).Value2).ToUpper();
                    strValueForValidation = Convert.ToString((range.Cells[intCounter, 3] as Excel.Range).Value2).ToUpper();
                    blnItemFound = false;

                    blnFatalError = TheDataValidationClass.VerifyDoubleData(strValueForValidation);

                    if(blnFatalError == false)
                    {
                        decTechPayPrice = Convert.ToDecimal(strValueForValidation);
                    }

                    TheFindTechPayItemByCodeDataSet = TheTechPayClass.FindTechPayItemByCode(strTechPayCode);

                    intRecordsReturned = TheFindTechPayItemByCodeDataSet.FindTechPayItemByCode.Rows.Count;

                    if(intRecordsReturned > 0)
                    {
                        blnItemFound = true;
                    }

                    if(blnItemFound == false)
                    {
                        TheFindTechPayItemByDescriptionDataSet = TheTechPayClass.FindTechPayItemByDescription(strJobDescription);

                        intRecordsReturned = TheFindTechPayItemByDescriptionDataSet.FindTechPayItemByDescription.Rows.Count;

                        if (intRecordsReturned > 0)
                        {
                            blnItemFound = true;
                        }
                    }

                    if(blnItemFound == false)
                    {
                        ImportedTechPayDataSet.importedtechpayRow NewTechPayRow = TheImportedTechPayDataSet.importedtechpay.NewimportedtechpayRow();

                        NewTechPayRow.JobDescription = strJobDescription;
                        NewTechPayRow.TechPayCode = strTechPayCode;
                        NewTechPayRow.TechPayPrice = decTechPayPrice;

                        TheImportedTechPayDataSet.importedtechpay.Rows.Add(NewTechPayRow);
                    }
                }

                PleaseWait.Close();
                dgrResults.ItemsSource = TheImportedTechPayDataSet.importedtechpay;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Tech Pay Items // Import Excel Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void BtnProcess_Click(object sender, RoutedEventArgs e)
        {
            int intCounter;
            int intNumberOfRecords;
            string strTechPayCode;
            string strJobDescription;
            decimal decTechPayPrice;
            bool blnFatalError = false;

            try
            {
                intNumberOfRecords = TheImportedTechPayDataSet.importedtechpay.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strTechPayCode = TheImportedTechPayDataSet.importedtechpay[intCounter].TechPayCode;
                    strJobDescription = TheImportedTechPayDataSet.importedtechpay[intCounter].JobDescription;
                    decTechPayPrice = TheImportedTechPayDataSet.importedtechpay[intCounter].TechPayPrice;

                    blnFatalError = TheTechPayClass.InsertTechPayItem(strTechPayCode, strJobDescription, decTechPayPrice);

                    if (blnFatalError == true)
                        throw new Exception();
                }

                TheMessagesClass.InformationMessage("All the information has been imported");
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Tech Pay // Process Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
    }
}
