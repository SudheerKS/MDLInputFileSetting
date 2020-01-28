using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.Utils.Drawing;
using DevExpress.Skins;
using DevExpress.Utils;
using DevExpress.XtraEditors.Drawing;
using DevExpress.XtraEditors.ViewInfo;
using DevExpress.XtraEditors.Repository;

namespace WindowsApplication1
{
    public static class DrawEditorHelper
    {
        public static void DrawEdit(Graphics g, RepositoryItem edit, Rectangle r, object value)
        {
            BaseEditViewInfo info = edit.CreateViewInfo() as BaseEditViewInfo;
            BaseEditPainter painter = edit.CreatePainter();
            info.EditValue = value;
            info.Bounds = r;
            info.CalcViewInfo(g);
            ControlGraphicsInfoArgs args = new ControlGraphicsInfoArgs(info, new GraphicsCache(g), r);
            painter.Draw(args);
            args.Cache.Dispose();
        }
    }
}
