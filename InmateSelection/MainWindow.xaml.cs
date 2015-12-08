using ClosedXML.Excel;
using Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace CardAssignment
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private DataSet ExcelData
        {
            get; set;
        }

        private string FilePath
        {
            get; set;
        }

        private string NewSheetName
        {
            get; set;
        }

        private SolidColorBrush ErrorColor
        {
            get
            {
                return new SolidColorBrush(Color.FromRgb(255, 58, 14));
            }
        }

        private SolidColorBrush SuccessColor
        {
            get
            {
                return new SolidColorBrush(Color.FromRgb(66, 186, 42));
            }
        }

        public MainWindow()
        {   
            InitializeComponent();
        }

        #region "Page Events"

        /// <summary>
        /// Starts the loading process for the file that was selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            lblProcessing.Visibility = Visibility.Visible;

            fileDialog.Filter = "Excel files (.xlsx)|*.xlsx|All Files (*.*)|*.*";
            fileDialog.FilterIndex = 1;

            bool? userClickedOk = fileDialog.ShowDialog();
            try
            {
                if (userClickedOk == true)
                {
                    FilePath = fileDialog.FileName;
                    ExcelData = LoadExcel(FilePath);
                    lstSheets.Items.Clear();
                    foreach (DataTable table in ExcelData.Tables)
                    {
                        lstSheets.Items.Add(table.TableName);
                    }
                    lstSheets.Visibility = Visibility.Visible;
                    lblSheetList.Visibility = Visibility.Visible;
                }
                lblProcessing.Foreground = SuccessColor;
                lblProcessing.Content = "Success!";
            }
            catch (Exception ex)
            {
                lblProcessing.Foreground = ErrorColor;
                lstSheets.Visibility = Visibility.Hidden;
                lblSheetList.Visibility = Visibility.Hidden;
                if (ex.ToString().Contains("being used by another process"))
                {
                    lblProcessing.Content = "File is open. Close and try again.";
                }
                else
                {
                    lblProcessing.Content = "Error occurred!";
                }
            }
        } // btnSelectFile_Click

        /// <summary>
        /// Processes the selected sheet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnProcess_Click(object sender, RoutedEventArgs e)
        {
            lblCompleted.Content = "Processing...";
            NewSheetName = txtNewSheetName.Text.Trim();

            if (lstSheets.Items.Contains(NewSheetName))
            {
                lblError.Foreground = ErrorColor;
                lblError.Visibility = Visibility.Visible;
                lblError.Text = NewSheetName + " sheet already exists. Change the name of the new sheet and try again.";
            }
            else
            {
                try
                {
                    ProcessExcel(lstSheets.SelectedValue.ToString());
                    lblCompleted.Content = "Success!";
                }
                catch (Exception ex)
                {
                    lblCompleted.Foreground = ErrorColor;
                    lblCompleted.Content = "Error!";
                }
            }
        } // btnProcess_Click

        /// <summary>
        /// Captures the list selection change
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lstSheets_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            lblSheetName.Visibility = Visibility.Visible;
            txtNewSheetName.Visibility = Visibility.Visible;
            txtNewSheetName.Text = lstSheets.SelectedValue + "_SendList";

            if (lstSheets.Items.Contains(txtNewSheetName.Text))
            {
                int sendListCounter = 2;

                while (lstSheets.Items.Contains($"{txtNewSheetName.Text}{sendListCounter}"))
                {
                    sendListCounter++;
                }

                txtNewSheetName.Text = $"{txtNewSheetName.Text}{sendListCounter}";
            }

            btnProcess.Visibility = Visibility.Visible;
        } // lstSheets_SelectionChanged

        #endregion

        #region "Excel"

        /// <summary>
        /// Opens the excel sheet and loads the entire workbook into a dataset
        /// </summary>
        /// <param name="file">File path</param>
        /// <returns>Filled dataset with the workbook information</returns>
        private DataSet LoadExcel(string file)
        {
            FileStream excelSteam = File.Open(file, FileMode.Open, FileAccess.Read);

            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelSteam);

            excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet();

            excelReader.Close();

            return result;
        } // LoadExcel

        /// <summary>
        /// Processes the excel sheet, creating mom and child objects
        /// </summary>
        /// <param name="SheetName">Name that will be assigned to the sheet</param>
        private void ProcessExcel(string SheetName)
        {
            List<Mom> Moms = new List<Mom>();
            // create the list
            foreach (DataRow row in ExcelData.Tables[SheetName].Rows)
            {
                Mom newMom = new Mom(row);
                Moms.Add(newMom);
            }
            // assign each mom the requested number of cards
            Random rand = new Random();
            foreach (Mom mom in Moms)
            {
                for (int i = 0; i < mom.CardsRequested; i++)
                {
                    mom.ChildrenToSendCards.Add(SelectChild(Moms, mom, rand));
                }
            }
            WriteNewSheet(Moms);
            lblCompleted.Visibility = Visibility.Visible;
        } // ProcessExcel

        /// <summary>
        /// Updates the existing file with the new worksheet information
        /// </summary>
        /// <param name="Moms">List of completed Mom objects</param>
        private void WriteNewSheet(List<Mom> Moms)
        {
            XLWorkbook workbook = new XLWorkbook(FilePath);
            DataTable table = ConvertListToDataTable(Moms);
            workbook.Worksheets.Add(table);
            workbook.Save();
        } // WriteNewSheet

        #endregion

        /// <summary>
        /// Selects an individual child from the list of children
        /// </summary>
        /// <param name="Moms">List of Moms serialized from Excel sheet</param>
        /// <param name="currentMom">Mom currently being assigned a child</param>
        /// <param name="rand">Random object to use for selecting next child</param>
        /// <returns>Child object</returns>
        private Child SelectChild(List<Mom> Moms, Mom currentMom, Random rand)
        {
            List<Child> availableChildren = (from m in Moms where m.Child != null && m != currentMom select m.Child).ToList();

            int maxSelectedCount = (from c in availableChildren select c.SelectedCount).Max();
            int minSelectedCount = (from c in availableChildren select c.SelectedCount).Min();
            Child selected = null;
            
            // grab random child object, unless a child has already been selected this round
            if (maxSelectedCount == minSelectedCount)
            {
                selected = availableChildren[rand.Next(0, availableChildren.Count())];
            } else
            {
                // ensure that all children have a fair chance at being selected
                List<Child> tempChildren = (from c in availableChildren where c.SelectedCount == minSelectedCount select c).ToList();
                selected = tempChildren[rand.Next(0, tempChildren.Count())];
            }

            selected.SelectedCount++;
            return selected;
        } // SelectChild

        /// <summary>
        /// Converts a List of Moms into a DataTable that can be converted into an excel document
        /// </summary>
        /// <param name="Moms">List of previously filled out mom objects</param>
        /// <returns>DataTable</returns>
        private DataTable ConvertListToDataTable(List<Mom> Moms)
        {
            DataTable convertedTable = new DataTable();
            convertedTable.TableName = NewSheetName;
            convertedTable.Columns.Add("Cards Requested");
            convertedTable.Columns.Add("Mom");
            convertedTable.Columns.Add("Child");
            convertedTable.Columns.Add("DOC #");
            convertedTable.Columns.Add("Facility");
            convertedTable.Columns.Add("Address #1");
            convertedTable.Columns.Add("Address #2");
            convertedTable.Columns.Add("City");
            convertedTable.Columns.Add("State");
            convertedTable.Columns.Add("Zip");
            foreach (Mom mom in Moms)
            {
                // Add the basic mom information
                DataRow row = convertedTable.NewRow();
                row["Cards Requested"] = mom.CardsRequested;
                row["Mom"] = mom.Name;
                if (mom.Child == null)
                {
                    row["Child"] = string.Empty;
                } else
                {
                    row["Child"] = mom.Child.Name;
                }
                
                convertedTable.Rows.Add(row);
                // add a row for each child that was assigned to a mom
                foreach(Child child in mom.ChildrenToSendCards)
                {
                    DataRow childRow = convertedTable.NewRow();
                    childRow["Cards Requested"] = null;
                    childRow["Mom"] = mom.Name;
                    childRow["Child"] = child.Name;

                    childRow["DOC #"] = child.DOC;
                    childRow["Facility"] = child.Facility;
                    childRow["Address #1"] = child.Address1;
                    childRow["Address #2"] = child.Address2;
                    childRow["City"] = child.City;
                    childRow["State"] = child.State;
                    childRow["Zip"] = child.Zip;
                    convertedTable.Rows.Add(childRow);
                }
            }
            return convertedTable;
        } // ConvertListToDataTable
    }
}
