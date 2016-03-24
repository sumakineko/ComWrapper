using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ComWrapper
{
    /// <summary>
    /// 
    /// </summary>
    public class WrapperBase : IDisposable
    {
        private bool _disposed = false;

        private Dictionary<Guid, WrapperBase> _children = new Dictionary<Guid, WrapperBase>();

        /// <summary>
        /// 
        /// </summary>
        public Guid Id { get; private set; }

        /// <summary>
        /// 
        /// </summary>
        protected dynamic ComObject { get; private set; }

        /// <summary>
        /// 
        /// </summary>
        protected WrapperBase Parent { get; set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="parent">親のラッパーオブジェクト</param>
        /// <param name="com">対象ComObject</param>
        protected WrapperBase(WrapperBase parent, dynamic com)
        {
            Id = new Guid();
            Parent = parent;
            ComObject = com;

            if (parent != null) parent.AddChild(this);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="obj"></param>
        protected WrapperBase(object obj) : this(null, obj)
        { }

        /// <summary>
        /// 管理対象の子ラッパーオブジェクトを追加
        /// </summary>
        /// <param name="child">追加する子ラッパーオブジェクト</param>
        protected void AddChild(WrapperBase child)
        {
            _children.Add(child.Id, child);
        }

        /// <summary>
        /// 子オブジェクトを削除
        /// </summary>
        /// <param name="child">削除する子ラッパーオブジェクト</param>
        protected void RemoveChild(WrapperBase child)
        {
            _children.Remove(child.Id);
        }

        /// <summary>
        /// Dispose
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            if (disposing)
            {
                List<Guid> keys = new List<Guid>(_children.Keys);
                foreach (var key in keys)
                {
                    var obj = _children[key];
                    if (obj != null) obj.Dispose();
                }

                Marshal.ReleaseComObject(ComObject);
                ComObject = null;

                Parent.RemoveChild(this);
            }

            _disposed = true;
        }

        /// <summary>
        /// 
        /// </summary>
        ~WrapperBase()
        {
            Dispose(false);
        }
    }
}
