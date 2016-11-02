using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Drawing;

namespace SpreadsheetEngine
{
    public abstract class Cell : INotifyPropertyChanged
    {
        private string _name = "";
        private int _row = 0;
        private int _column = 0;
        protected string _text = "";
        protected string _value = "";
        protected int _backColor = -1;
        protected float _textSize = 11;
        protected string _fontName = "Arial";
        protected string _fontStyle = "Normal";
        

        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        public Cell(int RowIndex, int ColumnIndex)
        {
            _row = RowIndex;
            _column = ColumnIndex;
            _name += Convert.ToChar('A' + ColumnIndex);
            _name += (RowIndex + 1).ToString();
        }

        public Cell()
        {
        }

        public string Name
        {
            get { return _name; }
        }

        // Row index property to retrieve row index
        public int RowIndex
        {
            get { return _row; }
        }

        // Column index property to retrieve column index
        public int ColumnIndex
        {
            get { return _column; }
        }

        // Value property
        public string Value
        {
            get { return _value; }
            internal set { }
        }

        // Text property
        public string Text
        {
            get { return _text; }

            set
            {
                if (value != _text)
                {
                    _text = value;
                    PropertyChanged(this, new PropertyChangedEventArgs("Text"));
                }
            }
        }

        // Text property
        public float TextSize
        {
            get { return _textSize; }

            set
            {
                if (value != _textSize)
                {
                    _textSize = value;
                    PropertyChanged(this, new PropertyChangedEventArgs("TextSize"));
                }
            }
        }

        // Text property
        public string FontName
        {
            get { return _fontName; }

            set
            {
                if (value != _fontName)
                {
                    _fontName = value;
                    PropertyChanged(this, new PropertyChangedEventArgs("FontName"));
                }
            }
        }

        // Text property
        public string Style
        {
            get { return _fontStyle; }

            set
            {
                if (value != _fontStyle)
                {
                    _fontStyle = value;
                    PropertyChanged(this, new PropertyChangedEventArgs("FontStyle"));
                }
            }
        }

        public int BackColor
        {
            get { return _backColor; }

            set
            {
                // Check for any changes. If so, fire changed event
                if (value != _backColor)
                {
                    _backColor = value;
                    PropertyChanged(this, new PropertyChangedEventArgs("BackColor"));
                }
            }
        }

        public bool Defaults
        {
            get
            {
                if (BackColor == -1 && string.IsNullOrEmpty(Text) && TextSize == 11)
                {
                    return true;
                }
                return false;
            }
        }

        public void Clear()
        {
            _text = "";
            _backColor = -1;
            _textSize = 11;
            _fontName = "Arial";
            _fontStyle = "Normal";
            PropertyChanged(this, new PropertyChangedEventArgs("Text"));
            PropertyChanged(this, new PropertyChangedEventArgs("BackColor"));
        }

    } // End of Cell class
}
