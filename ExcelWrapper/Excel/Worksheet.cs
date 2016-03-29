using MsExcel = Microsoft.Office.Interop.Excel;

namespace ComWrapper.Excel
{
    /// <summary>
    /// 
    /// </summary>
    public class Worksheet : WrapperBase<MsExcel.Worksheet>
    {
        /// <summary>
        /// 
        /// </summary>
        public Range Cells
        {
            get { return new Range(this, ComObject.Cells); }
        }

        /// <summary>
        /// 
        /// </summary>
        public Range Columns
        {
            get { return new Range(this, ComObject.Columns); }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Index
        {
            get { return ComObject.Index; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Name
        {
            get { return ComObject.Name; }
        }

        /// <summary>
        /// 
        /// </summary>
        public Range Rows
        {
            get { return new Range(this, ComObject.Rows); }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="obj"></param>
        public Worksheet(WrapperBase parent, MsExcel.Worksheet obj) : base(parent, obj)
        { }

        /// <summary>
        /// 
        /// </summary>
        public void Delete()
        {
            ComObject.Delete();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public Range Range(object x, object y = null)
        {
            return new Range(this, ComObject.Range[x, y]);
        }
    }
}
