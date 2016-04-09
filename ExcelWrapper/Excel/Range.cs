using System.Collections;
using System.Collections.Generic;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace ComWrapper.Excel
{
    /// <summary>
    /// 
    /// </summary>
    public class Range : WrapperBase<MsExcel.Range>
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public Range this[int rowIndex, int columnIndex]
        {
            get { return new Range(Parent, ComObject[rowIndex, columnIndex]); }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cell1"></param>
        /// <param name="cell2"></param>
        /// <returns></returns>
        public Range this[Range cell1, Range cell2]
        {
            get { return new Range(Parent, ComObject[cell1.ComObject, cell2.ComObject]); }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public Range this[string str]
        {
            get { return new Range(Parent, ComObject[str]); }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Column
        {
            get { return ComObject.Column; }
        }

        /// <summary>
        /// 
        /// </summary>
        public Range Columns
        {
            get { return new Range(Parent, ComObject.Columns); }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Row
        {
            get { return ComObject.Row; }
        }

        /// <summary>
        /// 
        /// </summary>
        public Range Rows
        {
            get { return new Range(Parent, ComObject.Rows); }
        }

        /// <summary>
        /// 
        /// </summary>
        public object NumberFormat
        {
            get { return ComObject.NumberFormat; }
            set { ComObject.NumberFormat = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public object NumberFormatLocal
        {
            get { return ComObject.NumberFormatLocal; }
            set { ComObject.NumberFormatLocal = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public object Value
        {
            get { return ComObject.Value; }
            set { ComObject.Value = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public object Value2
        {
            get { return ComObject.Value2; }
            set { ComObject.Value2 = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="obj"></param>
        public Range(WrapperBase parent, MsExcel.Range obj) : base(parent, obj)
        { }

        /// <summary>
        /// 
        /// </summary>
        public void Clear()
        {
            ComObject.Clear();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public IEnumerable<Range> GetEnumerator()
        {
            IEnumerator e = ComObject.GetEnumerator();
            while (e.MoveNext())
                yield return new Range(this, (MsExcel.Range)e.Current);
        }
    }
}
