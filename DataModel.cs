using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProcess
{
   public class DataModel
    {

        private string excelPath;
        public string ExcelPath
        {
            get => excelPath;
            set => excelPath = value;
        }

        private int columnRow;
        public int ColumnRow
        {
            get => columnRow;
            set => columnRow = value;
        }

        private int dataStartRow;
        public int DataStartRow
        {
            get => dataStartRow;
            set => dataStartRow = value;
        }

        private int dataEndRow;
        public int DataEndRow
        {
            get => dataEndRow;
            set => dataEndRow = value;
        }

    }
}
