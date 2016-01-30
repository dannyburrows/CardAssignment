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
            //Clear screen
            lblSheetName.Visibility = Visibility.Hidden;
            txtNewSheetName.Visibility = Visibility.Hidden;
            btnProcess.Visibility = Visibility.Hidden;
            lblCompleted.Visibility = Visibility.Collapsed;
            lblError.Visibility = Visibility.Collapsed;

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
                if (lstSheets.Items.Count == 1)
                {
                    lstSheets.SelectedIndex = 0;
                }
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
            lblCompleted.Visibility = Visibility.Visible;
            lblCompleted.Content = "Processing...";
            NewSheetName = txtNewSheetName.Text.Trim();

            if (lstSheets.Items.Contains(NewSheetName))
            {
                string errorMessage = NewSheetName + " sheet already exists. Change the name of the new sheet and try again.";
                DisplayError(errorMessage);
            }
            else
            {
                try
                {
                    ProcessExcel(lstSheets.SelectedValue.ToString());
                    lblCompleted.Foreground = SuccessColor;
                    lblCompleted.Content = "Success!";
                }
                catch (Exception ex)
                {
                    string errorMessage;
                    if (ex.Message == "An item with the same key has already been added.")
                    {
                        errorMessage = NewSheetName + " sheet already exists. Change the name of the new sheet and try again.";
                    }
                    else if(ex.ToString().Contains("being used by another process"))
                    {
                        errorMessage = "File is open. Close and try again.";
                    }
                    else
                    {
                        errorMessage = "Error occurred!";
                    }

                    DisplayError(errorMessage);
                }
            }
        } // btnProcess_Click

        private void DisplayError(string errorMessage)
        {
            lblCompleted.Visibility = Visibility.Collapsed;
            lblError.Foreground = ErrorColor;
            lblError.Visibility = Visibility.Visible;
            lblError.Text = errorMessage;
        }

        /// <summary>
        /// Captures the list selection change
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lstSheets_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lstSheets.SelectedValue != null)
            {
                lblError.Visibility = Visibility.Collapsed;
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
            }
        } // lstSheets_SelectionChanged

        #endregion

        #region "Card Assignment"

        private static void AdjustNumberOfCardsRequested(List<Mom> Moms)
        {
            foreach (Mom currentMom in Moms)
            {
                //can't request more cards than there are available children
                int maxCardCount = Moms.Count(m => m.Name != currentMom.Name && m.HasParticipatingChild);

                if (currentMom.CardsRequested > maxCardCount)
                {
                    currentMom.CardsRequested = maxCardCount;
                }
            }
        }

        private static void SetNumberOfCardsNeeded(List<Mom> Moms)
        {
            foreach (Mom momWithParticipatingChild in Moms.Where(m => m.HasParticipatingChild))
            {
                //each child will receive at least one card
                momWithParticipatingChild.Child.CardsNeeded = Math.Max(1, momWithParticipatingChild.CardsRequested);
            }

            AdjustNumberOfCardsNeeded(Moms);
        }
        
        private static void AdjustNumberOfCardsNeeded(List<Mom> Moms)
        {
            while (Moms.Sum(m => m.CardsNeededForChild) > Moms.Sum(m => m.CardsRequested)
                && Moms.Any(m => m.CardsNeededForChild > 1))
            {
                HandleInsufficientCardsRequested(Moms, Moms.Sum(m => m.CardsNeededForChild) - Moms.Sum(m => m.CardsRequested));
            }

            while (Moms.Sum(m => m.CardsRequested) > Moms.Sum(m => m.CardsNeededForChild))
            {
                HandleExtraCardsRequested(Moms, Moms.Sum(m => m.CardsRequested) - Moms.Sum(m => m.CardsNeededForChild));
            }
        }

        private static void HandleInsufficientCardsRequested(List<Mom> Moms, int totalCardsToSubtract)
        {
            //children receiving the most cards will each "donate" one card to the children whose moms aren't sending cards
            List<Child> childrenDonatingCards = Moms.Where(m => m.CardsNeededForChild > 1)
                                                    .Select(m => m.Child)
                                                    .OrderByDescending(c => c.CardsNeeded)
                                                    .Take(totalCardsToSubtract)
                                                    .ToList();

            foreach (Child childDonatingCard in childrenDonatingCards)
            {
                childDonatingCard.CardsNeeded--;
            }
        }

        private static void HandleExtraCardsRequested(List<Mom> Moms, int totalCardsToAdd)
        {
            List<Child> childrenReceivingExtraCards = Moms.Where(m => m.HasParticipatingChild)
                                                            .Select(m => m.Child)
                                                            .OrderBy(c => c.CardsNeeded)
                                                            .Take(totalCardsToAdd)
                                                            .ToList();

            foreach (Child childReceivingExtraCard in childrenReceivingExtraCards)
            {
                childReceivingExtraCard.CardsNeeded++;
            }
        }

        private void AssignCards(List<Mom> Moms)
        {
            // assign each mom the requested number of cards
            foreach (Mom currentMom in Moms.OrderByDescending(m => Moms.Count(m2 => m2.Name == m.Name))
                                            .ThenByDescending(m => m.CardsRequested))
            {
                currentMom.ChildrenToSendCards.AddRange(SelectChildren(Moms, currentMom, currentMom.CardsRequested));
            }
        }

        /// <summary>
        /// Selects children from the list of children
        /// </summary>
        /// <param name="Moms">List of Moms serialized from Excel sheet</param>
        /// <param name="currentMom">Mom currently being assigned a child</param>
        /// <param name="numberToSelect">Number of children to select</param>
        /// <returns>Child object</returns>
        private List<Child> SelectChildren(List<Mom> Moms, Mom currentMom, int numberToSelect)
        {
            List<Child> availableChildren = Moms.Where(m => m.HasParticipatingChild && m.Name != currentMom.Name)
                                                .Select(m => m.Child)
                                                .Where(c => c.CardsNeeded > 0
                                                            && !Moms.Any(m => m.Name == currentMom.Name
                                                                                && m.ChildrenToSendCards.Any(cc => cc == c)))
                                                .ToList();

            List<Child> selectedChildren = availableChildren.OrderByDescending(c => c.CardsNeeded)
                                                            .ThenBy(c => new Guid())
                                                            .Take(numberToSelect)
                                                            .ToList();

            foreach (Child selectedChild in selectedChildren)
            {
                selectedChild.CardsNeeded--;
            }

            return selectedChildren;
        } // SelectChildren

        private void CheckForUnassignedCards(List<Mom> Moms)
        {
            if (Moms.Any(m => m.CardsNeededForChild > 0))
            {
                //this could be due to not enough cards being requested for the number of participating children
                DisplayError("Child(ren) with not enough cards assigned.");
            }
            else if (Moms.Any(m => m.CardsRequested > m.ChildrenToSendCards.Count))
            {
                DisplayError("Mom(s) with not enough cards assigned.");
            }
            else if (Moms.Any(m => m.CardsRequested < m.ChildrenToSendCards.Count))
            {
                DisplayError("Mom(s) with too many cards assigned.");
            }
        }

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
            List<Mom> Moms = GetMoms(SheetName);

            AdjustNumberOfCardsRequested(Moms);
            SetNumberOfCardsNeeded(Moms);
            AssignCards(Moms);
            CheckForUnassignedCards(Moms);
            
            WriteNewSheet(Moms);
        } // ProcessExcel

        private List<Mom> GetMoms(string SheetName)
        {
            List<Mom> Moms = new List<Mom>();
            
            foreach (DataRow row in ExcelData.Tables[SheetName].Rows)
            {
                Mom newMom = new Mom(row);
                Moms.Add(newMom);
            }

            return Moms;
        }

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
                foreach(Child child in mom.ChildrenToSendCards.OrderBy(c => c.Name))
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
