using MsExcel = Microsoft.Office.Interop.Excel;

namespace ComWrapper.Excel
{
    /// <summary>
    /// 
    /// </summary>
    public class Workbook : WrapperBase<MsExcel.Workbook>
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="comObject"></param>
        public Workbook(WrapperBase parent, MsExcel.Workbook comObject) : base(parent, comObject)
        { }

        /// <summary>
        /// 
        /// </summary>
        public void Save()
        {
            ComObject.Save();
        }
    }
}
