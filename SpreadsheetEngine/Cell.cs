using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace SpreadsheetEngine
{
    public abstract class Cell : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public int RowIndex { get; private set; }
        public int ColumnIndex { get; private set; }
        private string _value;
        protected Cell(int nRowIndex, int nColumnIndex)
        {
            RowIndex = nRowIndex;
            ColumnIndex = nColumnIndex;
        }

        public string Value
        {
            get => _value;
            protected set
            {
                if (_value != value)
                {
                    _value = value;
                    OnPropertyChanged();
                }
            }   
        }
        private string _text;
        public string Text
        {
            get => _text;
            set
            {
                if (_text != value)
                {
                    _text = value;
                    OnPropertyChanged();
                }
            }
        }

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
