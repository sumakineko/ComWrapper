using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace ComWrapper.Excel
{
    /// <summary>
    /// 
    /// </summary>
    public class Workbooks : WrapperBase<MsExcel.Workbooks>
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="idx"></param>
        /// <returns></returns>
        public Workbook this[int idx]
        {
            get { return new Workbook(this, ComObject[idx]); }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="comObject"></param>
        public Workbooks(WrapperBase parent, MsExcel.Workbooks comObject) : base(parent, comObject)
        { }

        /// <summary>
        /// 
        /// </summary>
        public void Add()
        {
            ComObject.Add();
        }

        /// <summary>
        /// 
        /// </summary>
        public void Close()
        {
            ComObject.Close();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public IEnumerable<Workbook> GetEnumerator()
        {
            IEnumerator e = ComObject.GetEnumerator();
            while (e.MoveNext())
                yield return new Workbook(this, (MsExcel.Workbook)e.Current);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="readOnly"></param>
        /// <returns></returns>
        public Workbook Open(string filename, bool readOnly = false)
        {
            return new Workbook(this, ComObject.Open(Filename: filename, ReadOnly: readOnly));
        }
    }
}
