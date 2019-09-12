using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace SpreadsheetEngine
{
    public class ExpTree
    {
        private IExpressionTreeNode _root;
        private Tuple<List<char>, List<string>> _parsed;
        private static Dictionary<string, Tuple<Cell, int>> _values; 
        public ExpTree(string expression)
        {
            if (_values == null)
                _values = new Dictionary<string, Tuple<Cell, int>>();
            _parsed = ParseString(expression);
            CreateTree();
        }

        public static void SetValueDictionary(List<string> keys, List<Cell> values, List<int> uses)
        {
            for (int i = 0; i < keys.Count; i++)
            {
                _values.Add(keys[i], new Tuple<Cell, int>(values[i], uses[i]));
            }
        }

        public static void AddCellToDictionary(string key, Cell value)
        {
            if (_values != null)
            {
                if (_values.ContainsKey(key))
                    _values[key] = new Tuple<Cell, int>(_values[key].Item1, _values[key].Item2 + 1);
                else
                    _values[key] = new Tuple<Cell, int>(value, 1);
            }
            else
            {
                _values = new Dictionary<string, Tuple<Cell, int>>();
                _values[key] = new Tuple<Cell, int>(value, 1);
            }
        }

        public static void RemoveFromDictionary(string key)
        {
            if (_values != null)
            {
                if (_values.ContainsKey(key))
                {
                    if (_values[key].Item2 == 1)
                        _values.Remove(key);
                    else
                        _values[key] = new Tuple<Cell, int>(_values[key].Item1, _values[key].Item2 - 1);
                }
            }
        }

        public double Eval()
        {
            if (_root != null)
                return _root.EvaluateNode();
            else
                return 0.0;
        }

        private int GetLeastImportantOpIndex(List<char> ops)
        {
            int index;
            if ((index = ops.FindLastIndex(p => p == '-' || p == '+')) >= 0)
            {
            }
            else if ((index = ops.FindLastIndex(p => p == '/' || p == '*' || p == '%')) >= 0)
            {
            }
            else if ((index = ops.FindLastIndex(p => p == '^')) >= 0)
            {
            }
            else
            {
                index = -1;
            }
            return index;
        }


        private bool NextValueIsUnaryNegetive(List<char> ops, int index)
        {
            if (ops.Count > index + 1 && ops[index + 1] == '_')
            {
                return true;
            }
            else
                return false;
        }

        private void OverwriteUsedOps(List<char> ops, List<int> removalIndexes)
        {
            foreach (int currentIndex in removalIndexes)
            {
                ops[currentIndex] = ' ';
            }
        }


        private List<Tuple<char, int>> GetOrderedOps()
        {
            List<Tuple<char, int>> opsOrdered = new List<Tuple<char, int>>();
            List<char> ops = _parsed.Item1;

            for (int i = 0; i < _parsed.Item1.Count; i++)
            {
                int index = GetLeastImportantOpIndex(ops);
                if (index >= 0)
                {
                    List<int> usedOpIndexes = new List<int>();
                    opsOrdered.Add(new Tuple<char, int>(ops[index], index));

                    if (NextValueIsUnaryNegetive(ops, index))
                    {
                        opsOrdered.Add(new Tuple<char, int>('_', index+1));
                        usedOpIndexes.Add(index+1);
                    }
                    usedOpIndexes.Add(index);
                    OverwriteUsedOps(ops, usedOpIndexes);
                }
            }

            return opsOrdered;
        }


        private void InsertOps(List<Tuple<char, int>> opsOrdered)
        {
            foreach (var cur in opsOrdered)
            {
                InsertNode((IOpNode)_root, ExpressionTreeNodeFactory.GetNode(cur.Item1, cur.Item2));
            }
        }


        private void RemoveBadData()
        {
            for (int i = 0; i < _parsed.Item2.Count; i++)
            {
                if (_parsed.Item2[i] == "")
                {
                    _parsed.Item2.RemoveAt(i--);
                    _parsed.Item2.RemoveAt(i--);
                }
            }
        }


        private void CreateTree()
        {
           
            List<Tuple<char, int>> opsOrdered = GetOrderedOps();
            InsertOps(opsOrdered);
            RemoveBadData();
            InsertDataNodes((IOpNode)_root);
        }


        private void InsertDataNodes(IOpNode curNode)
        {
            if (_root == null)
                _root = ExpressionTreeNodeFactory.GetNode(_parsed.Item2[0]);
            if (curNode == null)
                return;
            if (curNode.Left != null)
            {
                InsertDataNodes((IOpNode)curNode.Left);
            }
            else
            {
                curNode.Left = ExpressionTreeNodeFactory.GetNode(_parsed.Item2[0]);
                _parsed.Item2.RemoveAt(0);
            }

            if (curNode.Right != null)
            {
                InsertDataNodes((IOpNode)curNode.Right);
            }                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             
            else
            {                         
                curNode.Right = ExpressionTreeNodeFactory.GetNode(_parsed.Item2[0]);
                _parsed.Item2.RemoveAt(0);
             }
        }


        //inserts node based off of index, nodes should already be in operator order
        private void InsertNode(IOpNode curNode, IOpNode newNode)
        {
            if (_root == null)
            {
                _root = newNode;
                return;
            }
            if (newNode.Position > curNode.Position)
            {
                if (curNode.Right != null)
                    InsertNode((IOpNode) curNode.Right, newNode);
                else
                    curNode.Right = newNode;
            }
            else
            {
                if (curNode.Left != null)
                    InsertNode((IOpNode)curNode.Left, newNode);
                else
                    curNode.Left = newNode;
            }
        }


        private Tuple<List<char>, List<string>> ParseString(string str)
        {
            char[] ops = { '/', '*', '^', '%', '+', '-', '(', ')', '_' };
            List<char> opList = new List<char>();
            List<string> numList = new List<string>();
            string current = "";
            string parenTree = "";

            int lParen = 0;
            for (int i=0; i < str.Length; i++)
            {
                char c = str[i];
                if (ops.Contains(c))
                {
                    if (c == '(')
                    {
                        lParen++;
                        if (lParen > 1)
                            parenTree += c;
                    }
                    else if (c == ')')
                    {
                        lParen--;
                        if (lParen == 0)
                        {
                            numList.Add(parenTree);
                        }
                        else
                            parenTree += c;
                    }
                    else if (lParen > 0)
                        parenTree += c;
                    else
                    {
                        if (current != "")
                            numList.Add(current);
                        opList.Add(c);
                        current = "";
                        if (str.Length > i+1 && str[i+1] == '-')
                        {
                            str = str.Remove(i + 1, 1);
                            str = str.Insert(i + 1, "_");
                            numList.Add("0");
                        }
                    }
                }
                else
                {
                    if (lParen == 0)
                        current += c;
                    else
                        parenTree += c;
                }
            }
            if (current != "")
                numList.Add(current);
            return (new Tuple<List<char>, List<string>>(opList, numList));
        }


        private interface IExpressionTreeNode
        {
            double EvaluateNode();
        }


        private class MiniExpTree : ExpTree, IExpressionTreeNode
        {
            public MiniExpTree(string tree) : base(tree)
            {
            }
            public double EvaluateNode()
            {
                return base.Eval();
            }
        }


        private class ExpTreeDataNode : IExpressionTreeNode
        {
            private readonly double _data;
            public ExpTreeDataNode(double data)
            {
                _data = data;
            }

            public double EvaluateNode()
            {
                return _data;
            }
        }


        private class TreeVariableNode : IExpressionTreeNode
        {
            private string _key;
            public TreeVariableNode(string key)
            {
                _key = key;
            }

            public double EvaluateNode()
            {
                return Convert.ToDouble(_values[_key].Item1.Value);
            }

            private Tuple<int, int> ParseOutComma(string key)
            {
                string[] values = key.Split(',');
                if (values.Length == 2)
                {
                    return new Tuple<int, int>(Convert.ToInt32(values[0]), Convert.ToInt32(values[1]));
                }
                else
                {
                    return new Tuple<int, int>(-1,-1);
                }
            }
        }


        private class AddNode : IOpNode
        {
            public IExpressionTreeNode Right { get; set; }
            public IExpressionTreeNode Left { get; set; }
            public int Position { get; }
            public AddNode(int position)
            {
                Position = position;
            }
            public double EvaluateNode()
            {
                return Left.EvaluateNode() + Right.EvaluateNode();
            }
        }


        private class SubNode : IOpNode
        {
            public IExpressionTreeNode Right { get; set; }
            public IExpressionTreeNode Left { get; set; }
            public int Position { get; }
            public SubNode(int position)
            {
                Position = position;
            }
            public double EvaluateNode()
            {
                return Left.EvaluateNode() - Right.EvaluateNode();
            }
        }


        private class MultNode : IOpNode
        {
            public IExpressionTreeNode Right { get; set; }
            public IExpressionTreeNode Left { get; set; }
            public int Position { get; }
            public MultNode(int position)
            {
                Position = position;
            }
            public double EvaluateNode()
            {
                return Left.EvaluateNode() * Right.EvaluateNode();
            }
        }


        private class DivNode : IOpNode
        {
            public IExpressionTreeNode Right { get; set; }
            public IExpressionTreeNode Left { get; set; }
            public int Position { get; }
            public DivNode(int position)
            {
                Position = position;
            }
            public double EvaluateNode()
            {
                return Left.EvaluateNode() / Right.EvaluateNode();
            }
        }


        private class ModNode : IOpNode
        {
            public IExpressionTreeNode Right { get; set; }
            public IExpressionTreeNode Left { get; set; }
            public int Position { get; }
            public ModNode(int position)
            {
                Position = position;
            }
            public double EvaluateNode()
            {
                return Left.EvaluateNode() % Right.EvaluateNode();
            }
        }


        private class ExpNode : IOpNode
        {
            public IExpressionTreeNode Right { get; set; }
            public IExpressionTreeNode Left { get; set; }
            public int Position { get; }

            public ExpNode(int position)
            {
                Position = position;
            }
            public double EvaluateNode()
            {
                double right = Right.EvaluateNode();
                double left = Left.EvaluateNode();
                double total = left;
                for (int i = 0; i < right - 1; i++)
                {
                    total = total*left;
                }
                return total;
            }
        }


        private interface IOpNode : IExpressionTreeNode
        {
            IExpressionTreeNode Right { get; set; }
            IExpressionTreeNode Left { get; set; }
            int Position { get; }
        };


        private static class ExpressionTreeNodeFactory
        {
            public static IExpressionTreeNode GetNode(string data)
            {
                try
                {
                    return new ExpTreeDataNode(Convert.ToDouble(data));
                }
                catch
                {
                    var match = data.IndexOfAny(new char[] { '/', '*', '^', '%', '+', '-', '(', ')', '_' });

                    if (match <= 0)
                        return new TreeVariableNode(data);
                    else
                        return new MiniExpTree(data);
                }
            }
            public static IOpNode GetNode(char op, int position)
            {
                IOpNode node;
                switch (op)
                {
                    case '/':
                        node = new DivNode(position);
                        break;
                    case '*':
                        node = new MultNode(position);
                        break;
                    case '+':
                        node = new AddNode(position);
                        break;
                    case '_':
                    case '-':
                        node = new SubNode(position);
                        break;
                    case '^':
                        node = new ExpNode(position);
                        break;
                    case '%':
                        node = new ModNode(position);
                        break;
                    default:
                        node = null;
                        break;
                }
                return node;
            }
        }
    }
}
