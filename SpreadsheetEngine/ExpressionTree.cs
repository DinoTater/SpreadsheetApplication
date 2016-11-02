using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetEngine
{
    public class ExpTree
    {
        private Dictionary<string, double> m_vars = new Dictionary<string, double>();
        private Node m_root;
        private string _expression;
        private static char[] _ops = { '+', '-', '*', '/', '^' };

        public ExpTree(string expression)
        {
            m_root = Compile(expression);
            _expression = expression;
        }

        public void clearVariables()
        {
            m_vars.Clear();
        }

        public void SetExpression(string exp)
        {
            _expression = exp;
            m_vars.Clear();
            m_root = Compile(_expression);
        }

        private Node Compile(string exp)
        {
            // builds and returns root node
            int i = exp.Length - 1;

            if (string.IsNullOrEmpty(exp))
                return null;

            // Discover all sets of parenthesis
            if (exp[0] == '(')
            {
                int counter = 0;

                // Find end parenthesis, but take into account other sets of beginning parens
                for (int j = 0; j < exp.Length; j++)
                {
                    if (exp[j] == '(')
                        counter++;
                    else if (exp[j] == ')')
                    {
                        counter--;

                        if (counter == 0)
                        {
                            if (exp.Length - 1 != j) // Last parenthesis in that substring
                                break;
                            else // continue looking for parenthesis
                                return Compile(exp.Substring(1, exp.Length - 2));
                        }
                    }
                }
            }
            // Compile inside parenthesis. Reaching here means parenthesis are broken down and can compile within
            // Compile in reverse order of precedence, start with plus and move through to divide
            foreach (char op in _ops)
            {
                Node n = Compile(exp, op);
                if (n != null)
                    return n;
            }

            // Get here if we are at a leaf node
            return BuildSimple(exp);
        }

        // Compile around operator
        private Node Compile(string exp, char op)
        {
            int count = 0, i = exp.Length - 1;
            bool end = false;
            opNode on = null;

            // Check if right-associative (^)
            bool rightAssoc = false;
            if (op == '^')
            {
                rightAssoc = true;
                i = 0;
            }

            // Compile for that specific operator
            // May be parens involved, in which case find and use Compile call without operator to break it down
            while (!end)
            {
                if (exp[i] == '(')
                {
                    if (rightAssoc)
                        count--;
                    count++;
                }
                else if (exp[i] == ')')
                {
                    if (rightAssoc)
                        count++;
                    count--;
                }

                if (count == 0)
                {
                    // Found operator inside parens
                    if (exp[i] == op)
                    {
                        on = new opNode(exp[i]);
                        on.Left = Compile(exp.Substring(0, i));
                        on.Right = Compile(exp.Substring(i + 1));
                        return on;
                    }
                }

                if (rightAssoc)
                {
                    if (i == exp.Length - 1)
                        end = true;
                    i++;
                }
                else
                {
                    if (i == 0)
                        end = true;
                    i--;
                }
            }

            if (count != 0)
                throw new Exception();

            return null;
        }

        // Which operator are we using?
        public int GetOpIndex(string exp)
        {
            for (int i = exp.Length - 1; i > 0; i--)
            {
                if (_ops.Contains(exp[i]))
                    return i;
            }
            return -1;
        }

        // Used if exp is only one constant or only one variable
        private Node BuildSimple(string exp)
        {
            double num;
            if (double.TryParse(exp, out num))
            {
                return new constNode((double)num);
            }

            varNode vn = new varNode(exp);
            m_vars[exp] = 0.0;
            return vn;
        }

        class Node { }

        class constNode : Node
        {
            public double Value;

            public constNode(double value)
            {
                Value = value;
            }
        }

        class varNode : Node
        {
            public string VarName;

            public varNode(string varName)
            {
                VarName = varName;
            }
        }

        class opNode : Node
        {
            public char Op;
            public Node Left, Right;

            public opNode() { }

            public opNode(char op)
            {
                Op = op;
                Left = Right = null;
            }
        }

        // Sets the specified variable within the ExpTree variables dictionary
        public void SetVar(string varName, double varValue)
        {
            m_vars[varName] = varValue;
        }

        public string[] GetVariables()
        {
            return m_vars.Keys.ToArray();
        }

        public double Eval()
        {
            //ExpTree newTree = new ExpTree(_expression);
            double result = Eval(m_root);
            return result;
        }

        // Implement this member function with no parameters that evaluates the expression to a double value
        private double Eval(Node n)
        {
            constNode ConNode = n as constNode;
            if (ConNode != null)
                return ConNode.Value;

            varNode VarNode = n as varNode;
            if (VarNode != null)
                return m_vars[VarNode.VarName];

            opNode oNode = n as opNode;
            if (oNode != null)
            {
                switch (oNode.Op)
                {
                    case '+':
                        return Eval(oNode.Left) + Eval(oNode.Right);
                    case '-':
                        return Eval(oNode.Left) - Eval(oNode.Right);
                    case '*':
                        return Eval(oNode.Left) * Eval(oNode.Right);
                    case '/':
                        return Eval(oNode.Left) / Eval(oNode.Right);
                    case '^':
                        return Math.Pow(Eval(oNode.Left), Eval(oNode.Right));
                }
                throw new NullReferenceException();
            }
            else
                throw new NullReferenceException();
        }
    } // End of ExpTree
}
