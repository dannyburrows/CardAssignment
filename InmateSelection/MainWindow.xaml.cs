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

namespace InmateSelection
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
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
                DataSet data = LoadExcel(fileName);
                //ProcessExcel(ref data);
            }
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

        private void ProcessExcel(ref DataSet Data)
        {
            List<Mom> Moms = new List<Mom>();
            // create the list
            foreach(DataRow row in Data.Tables[0].Rows)
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
    }
}
