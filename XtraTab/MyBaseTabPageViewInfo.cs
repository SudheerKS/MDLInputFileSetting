using System;
using System.Drawing;
using DevExpress.XtraTab.ViewInfo;

namespace WindowsApplication1
{
    public class MyBaseTabPageViewInfo : BaseTabPageViewInfo
    {
        public MyBaseTabPageViewInfo(DevExpress.XtraTab.IXtraTabPage page)
            : base(page) { }

        public override bool HasImage {
            get {
                return true;
            }
        }

        public override Size ImageSize {
            get {
                return new Size(16, 16);
            }
        }
    }
}
