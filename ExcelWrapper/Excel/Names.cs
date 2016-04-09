using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace ComWrapper.Excel
{
    public class Names : WrapperBase<MsExcel.Names>, IEnumerable<Name>
    {
        public int Count
        {
            get { return ComObject.Count; }
        }

        public Name this[int idx]
        {
            get { return new Name(this, ComObject.Item(idx)); }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="comObject"></param>
        public Names(WrapperBase parent, MsExcel.Names comObject) : base(parent, comObject)
        { }
        
        public IEnumerator<Name> GetEnumerator()
        {
            IEnumerator e = ComObject.GetEnumerator();
            while (e.MoveNext())
                yield return new Name(this, (MsExcel.Name)e.Current);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
