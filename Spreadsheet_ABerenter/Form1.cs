// Adam Berenter
// 11440727

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using SpreadsheetEngine;

namespace Spreadsheet_ABerenter
{
    public partial class Butter : Form
    {
        private Spreadsheet cellSheet = new Spreadsheet (50, 26);
        private bool changes = false;

        public Butter()
        {
            InitializeComponent();
            cellSheet.SpreadsheetCellPropertyChanged += DataCellChanged;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        // Load form with 26 columns labelled A-Z, and rows from 1-50
        private void Form1_Load(object sender, EventArgs e)
        {
            for (int i = 0; i < 26; i++)
            {
                int next = i + 65;
                string header = ((char)next).ToString();
                dataGridView1.Columns.Add(header, header);
            }

            dataGridView1.Rows.Add(50);
            for (int i = 1; i <= 50; i++)
            {
                dataGridView1.Rows[i - 1].HeaderCell.Value = i.ToString();
            }

            foreach (FontFamily family in FontFamily.Families)
            {
                comboBox1.Items.Add(family.Name);
            }

            int[] ptArray = new int[] { 6,7,8,9,10,11,12,14,16,18,20,22,24,26,28,36,72,128};

            //int[] ptArray = Enumerable.Range(6, 7).Select(i => (int)i).ToArray();
            foreach (int num in ptArray)
            {
                comboBox2.Items.Add(num);
            }

            updateToolStrip();
            dataGridView1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int i = 0;
            Random rnd = new Random();

            // Print "Hello World!" in approximately 50 cells. Set to 60 in case it is in column A or B for parts 2/3 in this section
            for (i = 0; i < 60; i++)
            {
                int iRow = rnd.Next(0, 50);
                int iCol = rnd.Next(0, 26);
                Cell c = cellSheet.GetCell(iRow, iCol);
                string newText = "Hello World!";
                c.Text = newText;
                dataGridView1.Rows[iRow].Cells[iCol].Value = cellSheet.GetCell(iRow, iCol).Value;
            }

            // Set all cells in column B to say "This is cell B#"
            for (i = 0; i < 50; i++)
            {
                Cell c = cellSheet.GetCell(i, 1);
                string newText = "This is cell " + ((char)(c.ColumnIndex + 'A')).ToString() + (c.RowIndex + 1).ToString();
                c.Text = newText;
                dataGridView1.Rows[i].Cells[1].Value = c.Value;
            }

            // Copy cells from B to it's equivalent A cell
            for (i = 0; i < 50; i++)
            {
                Cell c = cellSheet.GetCell(i, 0);
                string newText = "=B" + (c.RowIndex).ToString();
                c.Text = newText;
                dataGridView1.Rows[i].Cells[0].Value = c.Value;
            }

        }

        private void DataCellChanged (object sender, PropertyChangedEventArgs e)
        {
            Cell c = sender as Cell;
            if (c != null)
            {
                if (e.PropertyName == "Value")
                {
                    string value = c.Value;
                    dataGridView1[c.ColumnIndex, c.RowIndex].Value = value;
                    changes = true;
                }
                else if (e.PropertyName == "Font")
                {
                    changes = true;
                    //FontFamily family = c.FontName;
                    float size = c.TextSize;
                    //FontStyle style = c.Style;
                    // dataGridView1.Rows[c.RowIndex].Cells[c.ColumnIndex].Style.Font = new Font(name, size,  style);
                    //dataGridView1[c.ColumnIndex, c.RowIndex].Style.Font = new Font(family, size);
                }
                else if (e.PropertyName == "BackColor")
                {
                    changes = true;
                    dataGridView1.Rows[c.RowIndex].Cells[c.ColumnIndex].Style.BackColor = Color.FromArgb(c.BackColor);
                }
            }
        }

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            Cell c = cellSheet.GetCell(e.RowIndex, e.ColumnIndex);
            dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = c.Text;
        }
        
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            Cell c = cellSheet.GetCell(e.RowIndex, e.ColumnIndex);
                        
           string newText;

            // Get text from cell, if nonexistent use ""
            try
            {
                newText = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            }
            catch (NullReferenceException)
            {
                newText = "";
            }

            // Instantiate undos command
            IUndoRedoCmd[] undos = new IUndoRedoCmd[1];

            //only one undo for restore text
            undos[0] = new RestoreText(c.Text, c.Name);
            
            // Set text in spreadsheet, add undo into undo/redo system and write cell to spreadsheet
            // Don't forget to update tool strip!!
            c.Text = newText;
            cellSheet.u_r.addUndo(new UndoRedoCollection(undos, "cell text change"));
            updateToolStrip();
            dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = c.Value;
        }

        private void dataGridView1_BackgroundColorChanged(object sender, EventArgs e)
        {
            
        }

        // update tool strip to read for appropriate changes
        private void updateToolStrip()
        {
            ToolStripMenuItem group = menuStrip1.Items[1] as ToolStripMenuItem;

            foreach (ToolStripItem item in group.DropDownItems)
            {
                if (item.Text.Substring(0,4) == "Undo")
                {
                    button4.Enabled = cellSheet.u_r.undo_poss;
                    item.Enabled = cellSheet.u_r.undo_poss;
                    item.Text = "Undo " + cellSheet.u_r.undoAction;
                }
                else if (item.Text.Substring(0, 4) == "Redo")
                {
                    button5.Enabled = cellSheet.u_r.redo_poss;
                    item.Enabled = cellSheet.u_r.redo_poss;
                    item.Text = "Redo " + cellSheet.u_r.redoAction;
                }
            }
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            cellSheet.u_r.undo(cellSheet);
            updateToolStrip();
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            cellSheet.u_r.redo(cellSheet);
            updateToolStrip();
        }

        // Use color dialog to set background color
        private void chooseBackgroundColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            List<IUndoRedoCmd> undos = new List<IUndoRedoCmd>();
           ColorDialog colorDialog = chooseBackgroundColorToolStripMenuItem.Tag as ColorDialog;

            //Put tag here to maintain customized color... i.e. keep old dialog box
            if (null == colorDialog)
            {
                colorDialog = new ColorDialog();
                chooseBackgroundColorToolStripMenuItem.Tag = colorDialog;
            }

            if (colorDialog.ShowDialog(this) == DialogResult.OK)
            {
                // Get the chosen color from user as int
                int chosen = colorDialog.Color.ToArgb();

                // Make changes for all cells selected
                foreach (DataGridViewCell dgc in dataGridView1.SelectedCells)
                {
                    Cell c = cellSheet.GetCell(dgc.RowIndex, dgc.ColumnIndex);
                    // Add undo for each cell selected so one undo/redo command changes all cells
                    undos.Add(new RestoreBackColor(c.BackColor, c.Name));

                    //Set color
                    c.BackColor = chosen;
                }

                // Officially add undo to collection with description of action that will occur
                cellSheet.u_r.addUndo(new UndoRedoCollection(undos, "background color change"));
                updateToolStrip();
            }
            else // User hit cancel
                return;
        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (changes == true)
                saveChangesDialog(sender, e);

            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "XML files (*.xml)|*.xml";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                Stream stream = new FileStream(openFileDialog.FileName, FileMode.Open, FileAccess.Read);
                cellSheet.Load(stream);
                stream.Dispose();
            }
            changes = false;
            updateToolStrip();
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XML files (*.xml)|*.xml";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Stream stream = new FileStream(saveFileDialog.FileName, FileMode.Create, FileAccess.Write);
                cellSheet.Save(stream);
                stream.Dispose();
            }
            changes = false;
        }

        private void chooseFontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var fontDialog = new FontDialog();
            List<IUndoRedoCmd> undos = new List<IUndoRedoCmd>();

            if (fontDialog.ShowDialog() == DialogResult.OK)
            {
                float chosen = fontDialog.Font.SizeInPoints;

                foreach (DataGridViewCell dgc in dataGridView1.SelectedCells)
                {
                    Cell c = cellSheet.GetCell(dgc.RowIndex, dgc.ColumnIndex);

                    // Add undo for each cell selected so one undo/redo command changes all cells
                    undos.Add(new RestoreSize(c.TextSize, c.Name));

                    //Set size
                    c.TextSize = chosen;
                }

                // Officially add undo to collection with description of action that will occur
                cellSheet.u_r.addUndo(new UndoRedoCollection(undos, "background color change"));
                updateToolStrip();
            }
            else // User hit cancel
                return;
        }

        private void toolStripContainer1_TopToolStripPanel_Click(object sender, EventArgs e)
        {

        }

        private void toolStripContainer1_ContentPanel_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            loadToolStripMenuItem_Click(sender, e);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            saveToolStripMenuItem_Click(sender, e);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            undoToolStripMenuItem_Click(sender, e);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            redoToolStripMenuItem_Click(sender, e);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<IUndoRedoCmd> undos = new List<IUndoRedoCmd>();

            int newFont = comboBox1.SelectedIndex;

            // Get the chosen font from user as int
            string chosen = FontFamily.Families[newFont].ToString();

            // Make changes for all cells selected
            foreach (DataGridViewCell dgc in dataGridView1.SelectedCells)
            {
                Cell c = cellSheet.GetCell(dgc.RowIndex, dgc.ColumnIndex);
                // Add undo for each cell selected so one undo/redo command changes all cells
                //undos.Add(new RestoreBackColor(c.BackColor, c.Name));

                //Set font
                c.FontName = chosen;
            }

            // Officially add undo to collection with description of action that will occur
            cellSheet.u_r.addUndo(new UndoRedoCollection(undos, "font change"));
            updateToolStrip();
            
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (changes == true)
                saveChangesDialog(sender, e);
        }

        private void saveChangesDialog(object sender, EventArgs e)
        {
            string message = "Would you like to save current spreadsheet?";
            string title = "Save";
            MessageBoxButtons buttons = MessageBoxButtons.YesNoCancel;

            DialogResult result;
            result = MessageBox.Show(message, title, buttons);

            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                saveToolStripMenuItem_Click(sender, e);
                cellSheet.ClearSheet();
            }
            else if(result == System.Windows.Forms.DialogResult.No)
            {
                cellSheet.ClearSheet();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            newToolStripMenuItem_Click(sender, e);
        }
    }
}
