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
using DevExpress.XtraTab;
using DevExpress.Utils;
using DevExpress.XtraTab.ViewInfo;
using DevExpress.XtraEditors.Repository;

namespace WindowsApplication1
{
    public class MySkinTabPainter : SkinTabPainter
    {
        public static RepositoryItemCheckEdit HeaderEdit = new RepositoryItemCheckEdit();
        public MySkinTabPainter(DevExpress.XtraTab.IXtraTab tabControl)
            : base(tabControl) { }

        protected override void DrawHeaderPageImage(TabDrawArgs e, BaseTabPageViewInfo pInfo)
        {
            XtraTabPage page = pInfo.Page as XtraTabPage;
            page.Tag = pInfo.Image;
            bool value = false;
            (page.TabControl.Tag as Dictionary<XtraTabPage, bool>).TryGetValue(page, out value);
            DrawEditorHelper.DrawEdit(e.Graphics, HeaderEdit, pInfo.Image, value);
        }
    }
}
