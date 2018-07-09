using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ExcelProcess
{
    public class DataViewModel : INotifyPropertyChanged
    {
        DataModel dataModel;
        public DataViewModel()
        {
            dataModel = new DataModel();
            BrowsCommand = new RelayCommand(BrowseClicked, param=>true);
            StartClick = new RelayCommand(StartProcessing, param => true);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        

        public string ExcelPath
        {
            get => dataModel.ExcelPath;
            set
            {
                if (value != dataModel.ExcelPath)
                {
                    dataModel.ExcelPath = value;
                    NotifyPropertyChanged();
                }
            }
         
        }

        private int columnRow;
        public int ColumnRow
        {
            get => dataModel.ColumnRow;
            set
            {
                if (value != dataModel.ColumnRow)
                {
                    dataModel.ColumnRow = value;
                    NotifyPropertyChanged();
                }

            }
        }

        private int dataStartRow;
        public int DataStartRow
        {
            get => dataModel.DataStartRow;
            set
            {
                if (value != dataModel.DataStartRow)
                {
                    dataModel.DataStartRow = value;
                    NotifyPropertyChanged();
                }

            }
        }

        private int dataEndRow;
        public int DataEndRow
        {
            get => dataModel.DataEndRow;
            set
            {
                if (value != dataModel.DataEndRow)
                {
                    dataModel.DataEndRow = value;
                    NotifyPropertyChanged();
                }

            }
        }

        private ICommand browseCommand;

        public ICommand BrowsCommand
        {
            get;set;
        }

        private void BrowseClicked(object obg)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Multiselect = false;
                //openFileDialog.Filter = "Text files (*.xls)|*.xls|All files (*.*)|*.*";
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (openFileDialog.ShowDialog() == true)
                {
                    this.ExcelPath = openFileDialog.FileName;
                      
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                
            }

        }

        public ICommand StartClick
        {
            get;set;
        }

        private int timeInterval;
        public int TimeInterval
        {
            get => timeInterval;
            set
            {
                if(timeInterval != value)
                {
                    timeInterval = value;
                    NotifyPropertyChanged();
                }
            }
        }

        private void StartProcessing(object obj)
        {
            ExcelReader excelReader = new ExcelReader(ExcelPath, sheetName: "Sheet1", ncolumnStartIndex: DataStartRow, timeInterval: TimeInterval);
            excelReader.StartProcessing();
        }

    }
}
