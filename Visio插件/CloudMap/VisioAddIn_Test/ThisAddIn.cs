using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Interop.Visio;

namespace VisioAddIn_Test
{
    public partial class ThisAddIn
    {
        private bool isKeyPressed = false;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 订阅图形选择事件
            this.Application.SelectionChanged += Application_SelectionChanged;
            // 订阅键盘按键事件
            this.Application.KeyUp += Application_KeyUp;
        }

        static private void Draw(double Smfactor, double Bifactor, double gap)
        {
            // 获取当前Visio应用程序实例
            Visio.Application visioApp = Globals.ThisAddIn.Application;
            // 获取当前文档
            Visio.Document doc = visioApp.ActiveDocument;
            // 获取当前页面
            //Visio.Page page = doc.Pages[1];
            Visio.Page page = doc.Application.ActivePage;
            // 获取选中的图形
            Visio.Selection selectedShapes = visioApp.ActiveWindow.Selection;
            if (selectedShapes.Count > 0)
            {
                foreach (Visio.Shape shape in selectedShapes)
                {
                    double X1 = (shape.get_Cells("PinX").ResultIU - shape.CellsU["Width"].ResultIU / 2) * Smfactor;
                    double Y1 = (shape.get_Cells("PinY").ResultIU - shape.CellsU["Height"].ResultIU / 2) * Smfactor;
                    double X2 = (shape.get_Cells("PinX").ResultIU + shape.CellsU["Width"].ResultIU / 2) * Bifactor;
                    double Y2 = (shape.get_Cells("PinY").ResultIU + shape.CellsU["Height"].ResultIU / 2) * Bifactor;

                    // 创建一个动态数组用于保存矩形上的点坐标
                    List<Tuple<double, double>> rectanglePoints = new List<Tuple<double, double>>();                 
                    // 在选中图形位置上 创建并放置一个新形状 - 矩形                
                    Visio.Shape shapeYun = page.DrawRectangle(X1, Y1, X2, Y2);
                    shapeYun.CellsU["LineColor"].FormulaU = "RGB(255, 0, 0)";  // 红色
                    shapeYun.CellsU["FillPattern"].FormulaU = "0";  // 无填充
                    // 获取当前文档的图层集合                                                                
                    Visio.Layers layers = page.Layers;                   
                    // 创建一个新图层
                    Visio.Layer newLayer = layers.Add("修改");                 
                }                  
            }          
        }

        private void Application_KeyUp(int KeyCode, int KeyButtonState, ref bool CancelDefault)
        {
            if (KeyCode == 82)
            {
                isKeyPressed = true;
            }
        }
        public void Application_SelectionChanged(Visio.Window window)
        {
            // 当图形选择发生变化时触发                        
            if (window.Selection.Count > 0 && isKeyPressed)
            {
                // 执行按下Ctrl键后的操作
                MessageBox.Show("进入修订模式，");
                ThisAddIn.Draw(1, 1, 0.1);
                isKeyPressed = false;
                window.DeselectAll();
            }
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.Application.SelectionChanged -= Application_SelectionChanged;
            // 订阅键盘按键事件
            this.Application.KeyUp -= Application_KeyUp;
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
