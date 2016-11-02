using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetEngine
{
    public interface ICopyPasteCmd
    {
        ICopyPasteCmd Execute(Spreadsheet ssheet);
    }

    // Collect undos and redos for easy access to each action
    public class CopyPasteCollection
    {
        public string _action;
        private ICopyPasteCmd[] _actions;

        public CopyPasteCollection()
        {
        }

        // Collection entered as array (used with text changes)
        public CopyPasteCollection(ICopyPasteCmd[] actions, string action)
        {
            _actions = actions;
            _action = action;
        }

        // Collection with list, used with multiple cells changed (used with background color for now)
        public CopyPasteCollection(List<ICopyPasteCmd> cmds, string action)
        {
            _actions = cmds.ToArray();
            _action = action;
        }

        public CopyPasteCollection Execute(Spreadsheet ssheet)
        {
            List<ICopyPasteCmd> copyList = new List<ICopyPasteCmd>();

            foreach (ICopyPasteCmd cmd in _actions)
            {
                copyList.Add(cmd.Execute(ssheet));
            }

            return new CopyPasteCollection(copyList.ToArray(), this._action);
        }
    } // End of collection

    /*
    // Actual implementation of copy/paste
    public class CopyPaste
    {
        private Stack<CopyPasteCollection> _pastes = new Stack<CopyPasteCollection>();

        //check if paste is possible
        public bool paste_poss
        {
            get { return _pastes.Count != 0; }
        }

        // Push last action onto undo stack
        // Clear redo stack at this point (eliminates possibility of redo)
        public void addUndo(CopyPasteCollection copies)
        {
            _pastes.Push(copies);
        }

        // Perform undo
        public void undo(Spreadsheet ssheet)
        {
            CopyPasteCollection actions = _pastes.Pop();
            _redos.Push(actions.Execute(ssheet));
        }

        // Perform redo
        public void redo(Spreadsheet ssheet)
        {
            CopyPasteCollection actions = _redos.Pop();
            _pastes.Push(actions.Execute(ssheet));
        }

        //clear both stacks
        public void Clear()
        {
            _pastes.Clear();
            _redos.Clear();
        }
    } // End of Undo/Redo class

    // Restore text
    public class RestoreText : ICopyPasteCmd
    {
        private string _text, _name;

        public RestoreText(string newText, string newName)
        {
            _text = newText;
            _name = newName;
        }

        public ICopyPasteCmd Execute(Spreadsheet ssheet)
        {
            Cell c = ssheet.GetCell(_name);
            string old = c.Text;
            c.Text = _text;
            return new RestoreText(old, _name);
        }
    }

    // Restore background color
    public class RestoreBackColor : ICopyPasteCmd
    {
        private int _color;
        private string _name;

        public RestoreBackColor(int newColor, string name)
        {
            _color = newColor;
            _name = name;
        }

        public ICopyPasteCmd Execute(Spreadsheet ssheet)
        {
            Cell c = ssheet.GetCell(_name);
            int old = c.BackColor;
            c.BackColor = _color;
            return new RestoreBackColor(old, _name);
        }
    }*/
}
