namespace ComWrapper
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class WrapperBase<T> : WrapperBase
    {
        /// <summary>
        /// 
        /// </summary>
        protected internal new T ComObject
        {
            get { return (T)base.ComObject; }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="obj"></param>
        protected WrapperBase(WrapperBase parent, T obj) : base(parent, obj)
        { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="comObject"></param>
        protected WrapperBase(T comObject) : base(comObject)
        { }
    }
}
