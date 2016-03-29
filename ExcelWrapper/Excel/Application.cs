using System;
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
    public class Application : WrapperBase<MsExcel.Application>
    {
        /// <summary>
        /// 
        /// </summary>
        public bool Visible
        {
            get { return ComObject.Visible; }
            set { ComObject.Visible = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public bool DisplayAlerts
        {
            get { return ComObject.DisplayAlerts; }
            set { ComObject.DisplayAlerts = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public bool EnableEvents
        {
            get { return ComObject.EnableEvents; }
            set { ComObject.EnableEvents = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public Application() : base(new MsExcel.Application())
        { }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public Workbooks WorkBooks()
        {
            var workbooks = ComObject.Workbooks;
            return new Workbooks(this, workbooks);
        }

        /// <summary>
        /// 
        /// </summary>
        public void Quit()
        {
            ComObject.Quit();
        }
    }
}
