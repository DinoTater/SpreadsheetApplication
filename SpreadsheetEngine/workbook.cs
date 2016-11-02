using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.ComponentModel;

namespace SpreadsheetEngine
{
    public class Workbook
    {
        private List<Spreadsheet> _book = new List<Spreadsheet>();
        private int _activeSheet = 0;
        public UndoRedo u_r = new UndoRedo();
        public event PropertyChangedEventHandler WorkbookChanged;

        public Spreadsheet Active
        {
            get
            {
                return _book[_activeSheet];
            }
        }

        public Workbook()
            : this(10, 10)
        {

        }

        public Workbook(int row, int column)
        {
            AddSheet(new Spreadsheet(row, column));
        }

        public void AddSheet(Spreadsheet sheet)
        {
            _book.Add(sheet);
            sheet.SpreadsheetCellPropertyChanged += SpreadsheetPropertyChanged;
        }

        public void SpreadsheetPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "Value")
                WorkbookChanged(sender, e);
            else if (e.PropertyName == "BackColor")
                WorkbookChanged(sender, e);
        }
    }
}
