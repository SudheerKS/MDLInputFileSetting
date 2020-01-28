using System;
using DevExpress.XtraTab.ViewInfo;

namespace WindowsApplication1
{
    public class MySkinTabHeaderViewInfo : SkinTabHeaderViewInfo
    {
        public MySkinTabHeaderViewInfo(BaseTabControlViewInfo viewInfo)
            : base(viewInfo) { }

        protected override BaseTabPageViewInfo CreatePage(DevExpress.XtraTab.IXtraTabPage page)
        {
            return new MyBaseTabPageViewInfo(page);
        }
    }
}
