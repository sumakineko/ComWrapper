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
    public class Sheets : WrapperBase<MsExcel.Sheets>
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="idx"></param>
        /// <returns></returns>
        public Worksheet this[object idx]
        {
            get { return new Worksheet(this, ComObject[idx]); }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="obj"></param>
        public Sheets(WrapperBase parent, MsExcel.Sheets obj) : base(parent, obj)
        { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="before"></param>
        /// <param name="after"></param>
        /// <param name="count"></param>
        public void Add(Worksheet before = null, Worksheet after = null, object count = null)
        {
            ComObject.Add(before == null ? null : before.ComObject, after == null ? null : after.ComObject, count == null ? null : count);
        }

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
        /// <returns></returns>
        public IEnumerable<Worksheet> GetEnumerator()
        {
            IEnumerator e = ComObject.GetEnumerator();
            while (e.MoveNext())
                yield return new Worksheet(this, (MsExcel.Worksheet)e.Current);
        }
    }
}
