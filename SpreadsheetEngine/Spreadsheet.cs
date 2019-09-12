using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Xml.Serialization;


namespace SpreadsheetEngine
{
    public class Spreadsheet
    {
        private class SpreadsheetCell : Cell
        {
            public List<SpreadsheetCell> FollowedCells;
            public bool IsFollowing => FollowedCells.Count > 0;
            private ExpTree _expTree;

            public SpreadsheetCell(int row, int column) : base(row, column)
            {
                FollowedCells = new List<SpreadsheetCell>();
            }

            public void FollowedCellPropertyChanged(object sender, PropertyChangedEventArgs e)
            {
                if (e.PropertyName == "Value")
                    EvaluateExpTree();
            }

            public void SetValue(string newValue)
            {
                base.Value = newValue;
            }

            public void CreateNewTree(string exp)
            {
                if (exp == "")
                    _expTree = null;
                else
                    _expTree = new ExpTree(exp);
            }

            public void EvaluateExpTree()
            {
                try
                {
                    if (_expTree != null)
                        base.Value = _expTree.Eval().ToString();
                    else if (FollowedCells.Count > 0)
                        base.Value = FollowedCells[0].Value;
                    else
                        base.Value = base.Text;
                }
                catch
                {
                    base.Value = "#REF!";
                }
            }

            public bool HasCircularReference()
            {
                return circularReferenceHelper(this, new List<Tuple<int, int>>());
            }

            private bool circularReferenceHelper(SpreadsheetCell curCell, List<Tuple<int, int>> previousCalls)
            {
                Tuple<int, int> thisCell = new Tuple<int, int>(curCell.RowIndex + 1, curCell.ColumnIndex);
                if (previousCalls.Contains(thisCell))
                    return true;

                previousCalls.Add(thisCell);
                foreach (var cur in curCell.FollowedCells)
                {
                    bool hasCircularReference = circularReferenceHelper(cur, previousCalls);
                    if (hasCircularReference)
                        return true;
                }

                previousCalls.Remove(thisCell);
                return false;
            }
        }

        public int ColumnCount { get; }
        public int RowCount { get; }
        public Cell[,] CellArray;

        public Spreadsheet(int rows, int columns)
        {
            CellArray = new Cell[rows, columns];
            InitializeCells(rows, columns);
            ColumnCount = columns;
            RowCount = rows;
        }

        private string GetCellName(int row, int column)
        {
            string str = "";
            str += char.ConvertFromUtf32('A' + column);
            str += (row + 1).ToString();
            return str;
        }

        private Tuple<int, int> GetCellLocationByName(string name)
        {
            Tuple<string, int> splitName = ParseTillNumber(name);
            int row = CalculateRow(splitName.Item2);
            int column = CalculateColumn(splitName.Item1);
            return new Tuple<int, int>(row, column);
        }

        private void InitializeCells(int rows, int columns)
        {
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    CellArray[i, j] = new SpreadsheetCell(i, j);
                    CellArray[i, j].PropertyChanged += CellPropertyChangedEvent;
                }
            }
        }

        private void RemoveAllFollowedCells(SpreadsheetCell cell)
        {
            if (cell.IsFollowing)
            {
                foreach (var tempCell in cell.FollowedCells)
                {
                    ExpTree.RemoveFromDictionary(GetCellName(tempCell.RowIndex, tempCell.ColumnIndex));
                    tempCell.PropertyChanged -= cell.FollowedCellPropertyChanged;
                }

                cell.FollowedCells.Clear();
            }
        }

        private List<string> ExtractCellNames(string expression)
        {
            if (expression[0] == '=')
                expression = expression.Substring(1);
            List<string> cellNames = new List<string>();
            char[] ops = { '/', '*', '^', '%', '+', '-', '(', ')' };
            string tempCellName = "";
            foreach (char cur in expression)
            {
                if (char.IsLetter(cur))
                {
                    tempCellName += cur;
                }
                else if (char.IsDigit(cur) && tempCellName.Length > 0)
                {
                    tempCellName += cur;
                }
                else if (ops.Contains(cur) && tempCellName.Length > 0)
                {
                    cellNames.Add(tempCellName);
                    tempCellName = "";
                }
            }

            if (tempCellName != "")
                cellNames.Add(tempCellName);
            return cellNames;
        }

        private int CalculateColumn(string str)
        {
            int column = 0;
            foreach (char c in str)
                column += char.ToUpper(c) - 65;
            return column;
        }

        private int CalculateRow(int preAdjusted)
        {
            return preAdjusted - 1;
        }

        private void FollowCellsInString(string expression, SpreadsheetCell cell)
        {
            List<string> cellNames = ExtractCellNames(expression);
            foreach (var cellName in cellNames)
            {
                Tuple<int, int> cellLocation = GetCellLocationByName(cellName);
                SpreadsheetCell foundCell = (SpreadsheetCell)GetCell(cellLocation.Item1, cellLocation.Item2);
                ExpTree.AddCellToDictionary(cellName, foundCell);
                foundCell.PropertyChanged += cell.FollowedCellPropertyChanged;
                cell.FollowedCells.Add(foundCell);
            }
        }

        public PropertyChangedEventHandler CellChangedEventSender;

        private void CellPropertyChangedEvent(object sender, PropertyChangedEventArgs e)
        {
            SpreadsheetCell cell = (SpreadsheetCell)sender;
            if (e.PropertyName == "Text")
            {
                RemoveAllFollowedCells(cell);
                string expression = cell.Text;

                if (string.IsNullOrEmpty(expression))
                {
                    cell.SetValue("");
                }
                else
                {
                    if (expression[0] == '=')
                        AddExpressionTree(cell);
                    else
                        cell.SetValue(expression);
                }

                CellChangedEventSender(cell, new PropertyChangedEventArgs("Text"));
            }
            else if (e.PropertyName == "Value")
            {
                CellChangedEventSender(cell, new PropertyChangedEventArgs("Value"));
            }
        }

        private void AddExpressionTree(SpreadsheetCell cell)
        {
            string expression = cell.Text.Substring(1);

            FollowCellsInString(expression, cell);
            if (cell.HasCircularReference())
            {
                cell.SetValue("Circular Reference!!!");
            }
            else
            {
                if (ExpressionIsJustACell(expression))
                    cell.CreateNewTree("");
                else
                    cell.CreateNewTree(expression);
                cell.EvaluateExpTree();
            }
        }

        private bool ExpressionIsJustACell(string expression)
        {
            if (ExtractCellNames(expression).Count == 1 && OperatorCount(expression) == 0)
                return true;
            return false;
        }

        private int OperatorCount(string expression)
        {
            if (expression[0] == '=')
                expression = expression.Substring(1);
            int amount = 0;
            char[] ops = { '/', '*', '^', '%', '+', '-', '(', ')' };
            foreach (var cur in expression)
            {
                if (ops.Contains(cur))
                    amount++;
            }

            return amount;
        }

        //take a string and returns a Tuple with the left being the column and the right being the row (not adjusted for 0)
        private Tuple<string, int> ParseTillNumber(string str)
        {
            if (string.IsNullOrEmpty(str))
                return null;
            for (int i = 0; i < str.Length; i++)
            {
                if (str[i] >= '0' && str[i] <= '9')
                    return new Tuple<string, int>(str.Substring(0, i), Convert.ToInt32(str.Substring(i)));
            }

            return null;
        }

        public Cell GetCell(int row, int column)
        {
            if (row >= RowCount || column >= ColumnCount || row < 0 || column < 0)
                return null;
            return CellArray[row, column];
        }

        private List<Cell> GetUsedCells()
        {
            List<Cell> usedCellList = new List<Cell>();
            foreach (Cell cell in CellArray)
            {
                if (!string.IsNullOrEmpty(cell.Text))
                    usedCellList.Add(cell);
            }

            return usedCellList;
        }

        public void SaveToXml(Stream stream)
        {
            List<Cell> usedCells = GetUsedCells();
            XDocument xdoc = new XDocument(
                new XDeclaration("1.0", "utf-8", null),
                new XElement("Root",
                    from cell in usedCells
                    select new XElement("Cell",
                        new XElement("Row", cell.RowIndex.ToString()),
                        new XElement("Column", cell.ColumnIndex.ToString()),
                        new XElement("Text", cell.Text))));
            xdoc.Save(stream);
        }

        public void ReadFromXml(Stream stream)
        {
            XDocument xdoc = XDocument.Load(stream);
            var cells = xdoc.Descendants("Cell");

            foreach (XElement cell in cells)
            {
                int cellRow = (int)cell.Descendants("Row").First();
                int cellColumn = (int)cell.Descendants("Column").First();
                string text = (string)cell.Descendants("Text").First();
                Cell tempCell = GetCell(cellRow, cellColumn);
                tempCell.Text = text;
            }
        }

        public void ClearSpreadsheet()
        {
            foreach (Cell cell in CellArray)
            {
                if (!string.IsNullOrEmpty(cell.Text))
                {
                    cell.Text = "";
                }
            }
        }
    }
}
