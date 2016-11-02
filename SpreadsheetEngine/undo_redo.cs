using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetEngine
{
    public interface IUndoRedoCmd
    {
        IUndoRedoCmd Execute(Spreadsheet ssheet);
    }

    // Collect undos and redos for easy access to each action
    public class UndoRedoCollection
    {
        public string _action;
        private IUndoRedoCmd[] _actions;

        public UndoRedoCollection()
        {
        }

        // Collection entered as array (used with text changes)
        public UndoRedoCollection(IUndoRedoCmd[] actions, string action)
        {
            _actions = actions;
            _action = action;
        }

        // Collection with list, used with multiple cells changed (used with background color for now)
        public UndoRedoCollection(List<IUndoRedoCmd> cmds, string action)
        {
            _actions = cmds.ToArray();
            _action = action;
        }

        public UndoRedoCollection Execute(Spreadsheet ssheet)
        {
            List<IUndoRedoCmd> actionList = new List<IUndoRedoCmd>();

            foreach (IUndoRedoCmd cmd in _actions)
            {
                actionList.Add(cmd.Execute(ssheet));
            }

            return new UndoRedoCollection(actionList.ToArray(), this._action);
        }
    } // End of collection

    // Actual implementation of undo/redo
    public class UndoRedo
    {
        private Stack<UndoRedoCollection> _undos = new Stack<UndoRedoCollection>();
        private Stack<UndoRedoCollection> _redos = new Stack<UndoRedoCollection>();

        //check if undo is possible
        public bool undo_poss
        {
            get { return _undos.Count != 0; }
        }

        //check if redo is possible
        public bool redo_poss
        {
            get { return _redos.Count != 0; }
        }

        //Get action of undo
        public string undoAction
        {
            get
            {
                if (undo_poss)
                    return _undos.Peek()._action;
                return "";
            }
        }

        //Get action of redo
        public string redoAction
        {
            get
            {
                if (redo_poss)
                    return _redos.Peek()._action;
                return "";
            }
        }

        // Push last action onto undo stack
        // Clear redo stack at this point (eliminates possibility of redo)
        public void addUndo(UndoRedoCollection undos)
        {
            _undos.Push(undos);
            _redos.Clear();
        }

        // Perform undo
        public void undo(Spreadsheet ssheet)
        {
            UndoRedoCollection actions = _undos.Pop();
            _redos.Push(actions.Execute(ssheet));
        }

        // Perform redo
        public void redo(Spreadsheet ssheet)
        {
            UndoRedoCollection actions = _redos.Pop();
            _undos.Push(actions.Execute(ssheet));
        }

        //clear both stacks
        public void Clear()
        {
            _undos.Clear();
            _redos.Clear();
        }
    } // End of Undo/Redo class

    // Restore text
    public class RestoreText : IUndoRedoCmd
    {
        private string _text, _name;

        public RestoreText(string newText, string newName)
        {
            _text = newText;
            _name = newName;
        }

        public IUndoRedoCmd Execute(Spreadsheet ssheet)
        {
            Cell c = ssheet.GetCell(_name);
            string old = c.Text;
            c.Text = _text;
            return new RestoreText(old, _name);
        }
    }

    // Restore size
    public class RestoreSize : IUndoRedoCmd
    {
        private string _name;
        private float _size;

        public RestoreSize(float newSize, string newName)
        {
            _size = newSize;
            _name = newName;
        }

        public IUndoRedoCmd Execute(Spreadsheet ssheet)
        {
            Cell c = ssheet.GetCell(_name);
            float old = c.TextSize;
            c.TextSize = _size;
            return new RestoreSize(old, _name);
        }
    }

    // Restore background color
    public class RestoreBackColor : IUndoRedoCmd
    {
        private int _color;
        private string _name;

        public RestoreBackColor(int newColor, string name)
        {
            _color = newColor;
            _name = name;
        }

        public IUndoRedoCmd Execute(Spreadsheet ssheet)
        {
            Cell c = ssheet.GetCell(_name);
            int old = c.BackColor;
            c.BackColor = _color;
            return new RestoreBackColor(old, _name);
        }
    }
}
