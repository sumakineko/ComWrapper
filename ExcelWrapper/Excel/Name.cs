using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace ComWrapper.Excel
{
    public class Name : WrapperBase<MsExcel.Name>
    {
        public string _Name
        {
            get { return ComObject.Name; }
            set { ComObject.Name = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="comObject"></param>
        public Name(WrapperBase parent, MsExcel.Name comObject) : base(parent,comObject)
        { }
    }
}
