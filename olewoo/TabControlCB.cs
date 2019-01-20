/**************************************
 *
 * Part of OLEWOO - http://www.benf.org
 *
 * CopyLeft, but please credit.
 *
 */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Drawing2D;

namespace olewoo_cs
{
    public partial class TabControlCB : TabControl
    {
        public TabControlCB()
        {
            InitializeComponent();
        }

        protected override void OnMouseEnter(EventArgs e) => base.OnMouseEnter(e);

        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            if (e.Bounds != RectangleF.Empty)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

                RectangleF tabTextArea = RectangleF.Empty;

                for (int nIndex = 0; nIndex < this.TabCount; nIndex++)
                {
                    tabTextArea = this.GetTabRect(nIndex);
                    if (nIndex == this.SelectedIndex)
                    {
                        var icon = new Rectangle((int)tabTextArea.X, (int)tabTextArea.Y, (int)tabTextArea.Width, (int)tabTextArea.Height);
                        e.Graphics.DrawImageUnscaled(this.ImageList.Images[0], new Point(icon.Left + icon.Width - 16, icon.Top + icon.Height - 16));
                    }
                    else
                    {
                          var _Path = new GraphicsPath();
                          _Path.AddRectangle(tabTextArea);
                          using (var _Brush = new LinearGradientBrush(tabTextArea, SystemColors.Control, SystemColors.ControlLight, LinearGradientMode.Vertical))
                          {
                              var _ColorBlend = new ColorBlend(3);
                              _ColorBlend.Colors = new Color[]{SystemColors.ControlLightLight, 
                                                          Color.FromArgb(255,SystemColors.ControlLight),SystemColors.ControlDark,
                                                          SystemColors.ControlLightLight};

                              _ColorBlend.Positions = new float[] { 0f, .4f, 0.5f, 1f };
                              _Brush.InterpolationColors = _ColorBlend;

                              e.Graphics.FillPath(_Brush, _Path);
                              using (var pen = new Pen(SystemColors.ActiveBorder))
                              {
                                  e.Graphics.DrawPath(pen, _Path);
                              }
                          }
                          _Path.Dispose();
                    }
                    string str = this.TabPages[nIndex].Text;
                    var stringFormat = new StringFormat();
                    stringFormat.Alignment = StringAlignment.Near;
                    tabTextArea.Offset(2, 2);
                    e.Graphics.DrawString(str, this.Font, new SolidBrush(this.TabPages[nIndex].ForeColor), tabTextArea, stringFormat);                    
                }
            }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
                var tabTextArea = (RectangleF)this.GetTabRect(SelectedIndex);
                var icon = new Rectangle((int)tabTextArea.X, (int)tabTextArea.Y, (int)tabTextArea.Width, (int)tabTextArea.Height);
                tabTextArea =
                    new RectangleF(tabTextArea.X + tabTextArea.Width - 16, 4, 16,16);
                var pt = new Point(e.X, e.Y);
                if (tabTextArea.Contains(pt))
                {
                    var icu = this.SelectedTab.Tag as IClearUp; // IDispose not appropriate.
                    if (icu != null) icu.ClearUp();
                    this.TabPages.Remove(this.SelectedTab);
                    if (_itr != null) _itr();
                }
        }

        public InformTabRemoved InformTabRemoved
        {
            set
            {
                _itr = value;
            }
        }
        InformTabRemoved _itr;
    }

    public delegate void InformTabRemoved();
}
