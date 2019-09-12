using System;

using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using SpreadsheetEngine;

namespace Spreadsheet
{
    public partial class Form1 : Form
    {
        public SpreadsheetEngine.Spreadsheet Spreadsheet;
        private Cell _cellBeingEdited;
        public Form1()
        {
            InitializeComponent();
            GenerateGrid();
            Spreadsheet = new SpreadsheetEngine.Spreadsheet(50, 26);
            Spreadsheet.CellChangedEventSender += CellPropertyChangedHandler;
        }

        private void CellPropertyChangedHandler(object sender, PropertyChangedEventArgs e)
        {
            if (sender is Cell cell)
            {
                if (e.PropertyName == "Text" || e.PropertyName == "Value")
                    dataGridView1[cell.ColumnIndex, cell.RowIndex].Value = cell.Value;
            }
        }
        private void GenerateGrid()
        {
            for (int i = 65; i < 91; i++)
            {
                DataGridViewColumn c = new DataGridViewTextBoxColumn();
                c.HeaderText = char.ConvertFromUtf32(i);
                c.Name = char.ConvertFromUtf32(i);
                dataGridView1.Columns.Add(c);
            }
            dataGridView1.Rows.Add(50);
            for (int i = 0; i < 50; i++)
            {
                dataGridView1.Rows[i].HeaderCell.Value = (i + 1).ToString();
            }
        }

        private void DemoButton_Click(object sender, EventArgs e)
        {
            Random rand = new Random();
            
            for (int i = 0; i < 50; i++)
            {
                Spreadsheet.GetCell(i, 1).Text = "This is B" + (i + 1).ToString();
                Spreadsheet.GetCell(i, 0).Text = "=B" + (i+1).ToString();
                Spreadsheet.GetCell(rand.Next(49), rand.Next(24)+2).Text = "Hello World";
            }
        }

        private void CellValueChangedEvent(object sender, DataGridViewCellEventArgs e)
        {
            Cell cell = Spreadsheet.GetCell(e.RowIndex, e.ColumnIndex);
            try
            {
                cell.Text = dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString();
            }
            catch
            {
                cell.Text = "";
            }
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Cell cell = Spreadsheet.GetCell(e.RowIndex, e.ColumnIndex);
            if (cell != null)
            {
                _cellBeingEdited = cell;
                topTextBox.Text = cell.Text;
            }
            topTextBox.Focus();
        }

        private void topTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                _cellBeingEdited.Text = topTextBox.Text;
            }
        }

        private void saveToXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog
            {
                InitialDirectory = @"c:\",
                Filter = @"xml files (*.xml)|*.xml|All files (*.*)|*.*",
                DefaultExt = "xml",
                FilterIndex = 2,
                AddExtension = true,
                RestoreDirectory = true
            };
            Stream stream;

            if (dialog.ShowDialog() == DialogResult.OK && (stream = dialog.OpenFile()) != null)
            {
                Spreadsheet.SaveToXml(stream);
                stream.Close();
            }
        }

        private void loadFromXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {

            OpenFileDialog dialog = new OpenFileDialog
            {
                InitialDirectory = @"c:\",
                Filter = @"xml files (*.xml)|*.xml|All files (*.*)|*.*",
                DefaultExt = "xml",
                FilterIndex = 2,
                AddExtension = true,
                RestoreDirectory = true
            };

            Stream stream;

            if (dialog.ShowDialog() == DialogResult.OK && (stream = dialog.OpenFile()) != null)
            {
                Spreadsheet.ClearSpreadsheet();
                Spreadsheet.ReadFromXml(stream);
                stream.Close();
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var frm = new About();
            frm.Location = this.Location;
            frm.StartPosition = FormStartPosition.Manual;
            frm.FormClosing += delegate { this.Show(); };
            frm.Show();
            this.Hide();
        }
    }
}
