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
using System.Windows.Controls;
using Microsoft.Win32;
using System.IO;
using Excel;
using System.Data;
using System.Linq;
using System.Collections.Generic;
using ClosedXML;
using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;

namespace InmateSelection
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

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();

            fileDialog.Filter = "Excel files (.xls)|*.xlsx|All Files (*.*)|*.*";
            fileDialog.FilterIndex = 1;

            bool? userClickedOk = fileDialog.ShowDialog();

            if (userClickedOk == true)
            {
                string fileName = fileDialog.FileName;
                ExcelData = LoadExcel(fileName);
                foreach(DataTable table in ExcelData.Tables)
                {
                    lstSheets.Items.Add(table.TableName);
                }
                
            }
        }

        private void btnProcess_Click(object sender, RoutedEventArgs e)
        {
            ProcessExcel(lstSheets.SelectedValue.ToString());
        }

        private DataSet LoadExcel(string file)
        {
            FileStream excelSteam = File.Open(file, FileMode.Open, FileAccess.Read);

            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelSteam);

            excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet();

            excelReader.Close();

            return result;
        }

        private void ProcessExcel(string SheetName)
        {
            List<Mom> Moms = new List<Mom>();
            // create the list
            foreach(DataRow row in ExcelData.Tables[SheetName].Rows)
            {
                Mom newMom = new Mom(row);
                Moms.Add(newMom);
            }
            // assign each mom the requested number of cards
            foreach(Mom mom in Moms)
            {
                for(int i = 0; i < mom.CardsRequested; i++)
                {
                    mom.ChildrenToSendCards.Add(SelectChild(Moms));
                }
            }
            WriteNewSheet(Moms);
            lblFinished.Visibility = Visibility.Visible;
        }

        private void WriteNewSheet(List<Mom> Moms)
        {
            XLWorkbook workbook = new XLWorkbook();
            DataTable table = ConvertListToDataTable(Moms);
            workbook.Worksheets.Add(table);
            workbook.SaveAs("C:\\Users\\danny\\Downloads\\testsave.xlsx");
        }

        private Child SelectChild(List<Mom> Moms)
        {
            int maxSelectedCount = (from m in Moms select m.Child.SelectedCount).Max();
            int minSelectedCount = (from m in Moms select m.Child.SelectedCount).Min();
            Child selected = null;
            Random rand = new Random();

            if (maxSelectedCount == minSelectedCount)
            {
                selected = Moms[rand.Next(0, Moms.Count())].Child;
                selected.SelectedCount++;
            } else
            {
                List<Mom> tempMoms = (from m in Moms where m.Child.SelectedCount == minSelectedCount select m).ToList();
                selected = tempMoms[rand.Next(0, tempMoms.Count())].Child;
                selected.SelectedCount++;
            }

            return selected;
        }

        private DataTable ConvertListToDataTable(List<Mom> Moms)
        {
            DataTable convertedTable = new DataTable();
            convertedTable.TableName = "Send List";
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
                DataRow row = convertedTable.NewRow();
                row["Cards Requested"] = mom.CardsRequested;
                row["Mom"] = mom.Name;
                row["Child"] = mom.Child.Name;
                convertedTable.Rows.Add(row);
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
        }
    }
}
