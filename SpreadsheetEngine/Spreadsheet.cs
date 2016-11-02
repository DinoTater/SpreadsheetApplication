// Adam Berenter
// 11440727

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.IO;
using System.Xml;
using System.Xml.Linq;


namespace SpreadsheetEngine
{
    public class Spreadsheet : Cell
    {
        public event PropertyChangedEventHandler SpreadsheetCellPropertyChanged = delegate { };
        private Cell[,] CellArray;
        private Dictionary<string, HashSet<string>> _cellDependencies;
        public UndoRedo u_r = new UndoRedo();

        // Need a way to instantiate a cell, which is otherwise abstract
        private class CellInstance : Cell
        {
            public CellInstance(int row, int col)
                : base(row, col)
            {
            }

            public void SetValue(string value)
            {
                _value = value;
            }
        }


        //Constructor
        public Spreadsheet(int rows, int columns)
        {
            CellArray = new Cell[rows, columns];
            _cellDependencies = new Dictionary<string, HashSet<string>>();

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    CellInstance newCell = new CellInstance(i, j);
                    newCell.PropertyChanged += CellPropertyChanged;
                    CellArray[i, j] = newCell;
                }
            }
        }

        // What happens when a cell changes?
        private void CellPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            CellInstance c = sender as CellInstance;

            if (e != null)
            {
                if (e.PropertyName == "Text")
                {
                    removeDependencies(c.Name);

                    if (c.Text != "" && c.Text[0] == '=' && c.Text.Length > 1) // Cell contains expression
                    {
                        ExpTree newTree = new ExpTree(c.Text.Substring(1));
                        addDependencies(c.Name, newTree.GetVariables());
                    }

                    evaluateCell(sender as Cell);

                    SpreadsheetCellPropertyChanged(c, e);
                }
                else if (e.PropertyName == "TextSize")
                {
                    SpreadsheetCellPropertyChanged(sender, new PropertyChangedEventArgs("TextSize"));
                }
                else if (e.PropertyName == "BackColor")
                {
                    SpreadsheetCellPropertyChanged(sender, new PropertyChangedEventArgs("BackColor"));
                }
            }
        }

        // Retrieve cell using row/column values
        public Cell GetCell(int RowIndexValue, int ColumnIndexValue)
        {
            return CellArray[RowIndexValue, ColumnIndexValue];
        }

        // Find a cell by variable name
        public Cell GetCell(string name)
        {
            Cell found;

            char column = name[0];
            if (Char.IsLetter(column) == false)
                return null;

            int row;
            if (int.TryParse(name.Substring(1), out row) == false)
                return null;

            try
            {
                found = GetCell(row - 1, column - 'A');
            }
            catch
            {
                return null;
            }

            // Found it!
            return found;
        }

        // Number of rows in the spreadsheet
        public int RowCount
        {
            get { return CellArray.GetLength(0); }
        }

        // Number of columns in the spreadsheet
        public int ColumnCount
        {
            get { return CellArray.GetLength(1); }
        }

        private void SetExpressionVariable(ExpTree current, string variable)
        {
            Cell variableCell = GetCell(variable);
            double value;
            if (string.IsNullOrEmpty(variableCell.Value))
            {
                //...set empty cells to 0.
                current.SetVar(variableCell.Name, 0);
            }
            else if (!double.TryParse(variableCell.Value, out value))
            {
                //...set non-value cells to 0.
                current.SetVar(variable, 0);
            }
            else
            {
                //...just set normally otherwise.
                current.SetVar(variable, value);
            }
        }

        private void evaluateCell(Cell cell)
        {
            CellInstance c = cell as CellInstance;


            if (string.IsNullOrEmpty(c.Text))
            {
                c.SetValue("");
                SpreadsheetCellPropertyChanged(cell, new PropertyChangedEventArgs("Value"));
            }
            else if (c.Text[0] == '=')
            {
                bool error = false;
                // Get string value without '=' in order to evaluate expression as an expression
                string exp = cell.Text.Substring(1);

                ExpTree newTree = new ExpTree(exp);

                // Need to get all variables from string and find values
                string[] variables = newTree.GetVariables();

                foreach (string variable in variables)
                {
                    // Check for self reference
                    if (variable == c.Name)
                    {
                        DisplayError(c, variable, "SELF REFERENCE");
                        error = true;
                        break;
                    }

                    // Check if cell exists
                    if (GetCell(variable) == null)
                    {
                        DisplayError(c, variable, "FALSE ID");
                        error = true;
                        break;
                    }

                    // Check if there is a 'multi-step' circular reference
                    if(circularReference(variable, c.Name))
                    {
                        DisplayError(c, variable, "CIRCULAR REFERENCE");
                        error = true;
                        break;
                    }

                    // Set variable into m_vars
                    SetExpressionVariable(newTree, variable);
                }
                
                if (error)
                {
                    return;
                }

                c.SetValue(newTree.Eval().ToString());
                SpreadsheetCellPropertyChanged(cell, new PropertyChangedEventArgs("Value"));
            }
            else // Not an expression, just a value
            {
                c.SetValue(c.Text);
                SpreadsheetCellPropertyChanged(cell, new PropertyChangedEventArgs("Value"));
            }

            // Check all dependencies
            if (_cellDependencies.ContainsKey(c.Name))
            {
                foreach (string dep in _cellDependencies[c.Name])
                {
                    evaluateCell(dep);
                }
            }
        }

        private void evaluateCell(string gridLoc)
        {
            evaluateCell(GetCell(gridLoc));
        }

        // Check if there is circular reference in expression
        public bool circularReference (string startName, string currName)
        {
            if (startName == currName)
            {
                return true;
            }

            if (!_cellDependencies.ContainsKey(currName))
            {
                return false;
            }

            // Recursive check for any 'multi-step' circular reference
            {
                foreach(string dep in _cellDependencies[currName])
                {
                    if(circularReference(startName, dep))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private void addDependencies(string name, string[] variables)
        {
            // Check expression for all dependencies
            foreach (string varName in variables)
            {
                if (_cellDependencies.ContainsKey(varName) == false)
                {
                    _cellDependencies[varName] = new HashSet<string>();
                }
                // Add the cell to dependency dictionary
                _cellDependencies[varName].Add(name);
            }
        }

        private void removeDependencies(string name)
        {
            // Need to remove from all cells that this cell is associated with
            List<string> keys = new List<string>();

            foreach (string key in _cellDependencies.Keys)
            {
                if (_cellDependencies[key].Contains(name))
                    keys.Add(key);
            }

            //Removal
            foreach (string key in keys)
            {
                HashSet<string> remKeys = _cellDependencies[key];

                if (remKeys.Contains(name))
                    remKeys.Remove(name);
            }
        }

        public bool Load(Stream stream)
        {
            XDocument document = null;

            try
            {
                document = XDocument.Load(stream);
            }
            catch(Exception)
            {
                return false;
            }

            if (document == null)
                return false;

            ClearSheet();

            XElement root = document.Root;
            foreach (XElement child in root.Elements("Spreadsheet"))
            {
                LoadHelper(child);
            }

            u_r.Clear();

            return true;
        }

        public void LoadHelper(XElement sElement)
        {
            foreach(XElement child in sElement.Elements("Cell"))
            {
                Cell c = GetCell(child.Attribute("Name").Value);

                // Check if we got anything
                if (c == null)
                    continue;

                // Load element, set in spreadsheet
                var textElement = child.Element("Text");
                if (textElement != null)
                {
                    c.Text = textElement.Value;
                }

                /*var sizeElement = child.Element("TextSize");
                if (textElement != null)
                {
                    c.TextSize = float.Parse(sizeElement.Value);
                }*/

                var backElement = child.Element("BackColor");
                if (backElement != null)
                {
                    c.BackColor = int.Parse(backElement.Value);
                }
            }
        }

        public bool Save(Stream stream)
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = Encoding.UTF8;
            settings.NewLineOnAttributes = false;
            settings.Indent = true;
            settings.NewLineChars = "\r\n";

            XmlWriter sWriter = XmlWriter.Create(stream, settings);
            if (sWriter == null)
            {
                return false;
            }

            sWriter.WriteStartElement("Spreadsheet");

            SaveHelper(sWriter);

            sWriter.WriteEndElement();
            sWriter.Close();
            
            return true;
        }

        public void SaveHelper(XmlWriter sWriter)
        {
            sWriter.WriteStartElement("Spreadsheet");

            var write = from Cell c in CellArray where (!c.Defaults) select c;

            foreach(Cell c in write)
            {
                sWriter.WriteStartElement("Cell");
                sWriter.WriteAttributeString("Name", c.Name);
                sWriter.WriteElementString("Text", c.Text);
                //sWriter.WriteElementString("TextSize", c.TextSize.ToString());
                sWriter.WriteElementString("BackColor", c.BackColor.ToString());

                sWriter.WriteEndElement();
            }

            sWriter.WriteEndElement();
        }

        public void ClearSheet()
        {
            for (int i = 0; i < RowCount; i++)
            {
                for (int j = 0; j < ColumnCount; j++)
                {
                    CellArray[i, j].Clear();
                }
            }
        }

        private void DisplayError(CellInstance location, string violator, string violation)
        {
            location.SetValue("#" + violation + " at cell " + violator);
            SpreadsheetCellPropertyChanged(location as Cell, new PropertyChangedEventArgs("Value"));
        }

    } //End of spreadsheet class
}