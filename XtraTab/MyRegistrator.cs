using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraTab.Registrator;
using DevExpress.XtraTab.Drawing;
using DevExpress.XtraTab.ViewInfo;
using DevExpress.XtraTab;
using WindowsApplication1;

namespace MDL
{
    public class MyRegistrator : SkinViewInfoRegistrator
    {
        public MyRegistrator() { }

        public override string ViewName {
            get {
                return "MyStyle";
            }
        }

        public override DevExpress.XtraTab.Drawing.BaseTabPainter CreatePainter(DevExpress.XtraTab.IXtraTab tabControl)
        {
            return new MySkinTabPainter(tabControl);
        }

        public override BaseTabHeaderViewInfo CreateHeaderViewInfo(BaseTabControlViewInfo viewInfo)
        {
            return new MySkinTabHeaderViewInfo(viewInfo);
        }
    }
}