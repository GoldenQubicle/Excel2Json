using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Json
{
    public class ValueGrid
    {
        private Dictionary<int, Dictionary<int, dynamic>> values = new Dictionary<int, Dictionary<int, dynamic>>();

        public int RowCount
        {
            get
            {
                return values
                    .Keys
                    .Max()
                    + 1;
            }
        }

        public int ColumnCount
        {
            get
            {
                return values
                    .Max(r => r.Key)
                    + 1;
            }
        }

        // The indexer
        public dynamic this[int row, int column]
        {
            get { return GetValue(row, column); }
            set { SetValue(row, column, value); }
        }

        public void Add(int row, int column, dynamic value)
        {
            SetValue(row, column, value);
        }

        private dynamic GetValue(int row, int column)
        {
            if (!values.ContainsKey(row))
            {
                return null;
            }

            if (!values[row].ContainsKey(column))
            {
                return null;
            }

            return values[row][column];
        }

        private void SetValue(int row, int column, dynamic value)
        {
            if (!values.ContainsKey(row))
            {
                values.Add(row, new Dictionary<int, dynamic>());
            }

            values[row][column] = value;
        }
    }
}
