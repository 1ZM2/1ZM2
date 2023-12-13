using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Visio;
using System.Windows.Forms;

namespace VisioAddInTest3
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
                //MessageBox.Show("键被按下，并且有图形被选中。");
                ThisAddIn.DrawYunMap(0.98, 1.02, 0.1);
                isKeyPressed = false;
                window.DeselectAll();
            }
        }
        static private void DrawYunMap(double Smfactor, double Bifactor, double gap)
        {
            // 获取当前Visio应用程序实例
            Visio.Application visioApp = Globals.ThisAddIn.Application;
            // 获取当前文档
            Visio.Document doc = visioApp.ActiveDocument;
            // 获取当前页面
            //Visio.Page page = doc.Pages[1];
            //Visio.Page activePage = doc.Application.ActivePage;
            //string pageName = activePage.Name;

            // 获取选中的图形
            Visio.Selection selectedShapes = visioApp.Window.Selection;
            int countCloudmap = 0;
            countCloudmap += 1;
            foreach (Visio.Shape shape in selectedShapes)
            {//将高度乘以0.418/10.5
                Visio.Page shapePage = shape.ContainingPage;
                string curname = shapePage.Name;
                int pageCount = doc.Pages.Count;
                for (int pageIndex = 1; pageIndex <= pageCount; pageIndex++)
                {
                    Visio.Page page = doc.Pages[pageIndex];
                    if (curname == page.Name)
                    {
                        Visio.Page shapePageCurr = doc.Pages[pageIndex];

                        double X1 = (getShapeValue(shape, "PinX") - getShapeValue(shape, "Width") / 2) * Smfactor;
                        double Y1 = (getShapeValue(shape, "PinY") - getShapeValue(shape, "Height") / 2) * Smfactor;
                        double X2 = (getShapeValue(shape, "PinX") - getShapeValue(shape, "Width") / 2) * Bifactor;
                        double Y2 = (getShapeValue(shape, "PinY") - getShapeValue(shape, "Height") / 2) * Bifactor;
                        // 计算矩形的宽和高
                        double epsilon = 1e-5;  // 10的负5次方
                        double width = Math.Abs(X2 - X1);
                        double height = Math.Abs(Y2 - Y1);

                        // 创建一个动态数组用于保存矩形上的点坐标
                        List<Tuple<double, double>> rectanglePoints = new List<Tuple<double, double>>();

                        //得到矩形上的点坐标
                        GetRectanglePoints(X1, Y1, X2, Y2, gap, rectanglePoints);
                        //Visio.Shape selectedShape = selectedShapes[1];

                        // 在选中图形位置上 创建并放置一个新形状 - 矩形                
                        Visio.Shape shapeYun = shapePageCurr.DrawRectangle(X1, Y1, X2, Y2);
                        shapeYun.Name = "图框" + countCloudmap.ToString("0000");
                        //使用上述坐标画弧
                        double x1 = 0;
                        double y1 = 0;
                        int rowi = 1;
                        //删除矩形的2到5行的geometry值
                        for (int i = 0; i < 4; i++)
                        {
                            shapeYun.DeleteRow((short)Visio.VisSectionIndices.visSectionFirstComponent, 2);
                        }
                        shapeYun.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, 1, 0].FormulaU = "\"" + 0 + "\"";
                        shapeYun.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, 1, 1].FormulaU = "\"" + 0 + "\"";

                        foreach (var point in rectanglePoints)
                        {
                            double x2 = point.Item1;
                            double y2 = point.Item2;
                            if (Math.Abs(x2 - x1) < epsilon && Math.Abs(y2 - y1) < epsilon)
                            {
                                continue;
                            }
                            rowi++;
                            DrawArcMap(shapeYun, rowi, ref x1, x2, ref y1, y2, width, height);
                            continue;
                        }
                        visioApp.Window.DeselectAll();
                    }
                }
            }
        }
        static private List<Tuple<double, double>> GetRectanglePoints(double X1, double Y1, double X2, double Y2, double gap, List<Tuple<double, double>> rectanglePoints)
        {
            // 计算矩形上的点坐标并添加到数组中
            // 计算矩形四条边上的点坐标并添加到数组中
            for (double x = Math.Min(X1, X2); x < Math.Max(X1, X2); x += gap)
            {
                rectanglePoints.Add(new Tuple<double, double>(x - X1, Y1 - Y1)); // 下边

            }
            for (double y = Math.Min(Y1, Y2); y < Math.Max(Y1, Y2); y += gap)
            {
                rectanglePoints.Add(new Tuple<double, double>(X2 - X1, y - Y1)); // 右边

            }
            for (double x = Math.Max(X1, X2); x > Math.Min(X1, X2); x -= gap)
            {
                rectanglePoints.Add(new Tuple<double, double>(x - X1, Y2 - Y1)); // 上边

            }
            for (double y = Math.Max(Y1, Y2); y > Math.Min(Y1, Y2); y -= gap)
            {
                rectanglePoints.Add(new Tuple<double, double>(X1 - X1, y - Y1)); // 左边                    
            }
            rectanglePoints.Add(new Tuple<double, double>(X1 - X1, Y1 - Y1)); // 回到原点
            return rectanglePoints;
        }
        static private void DrawArcMap(Visio.Shape shape, int row_i, ref double x_1, double x_2, ref double y_1, double y_2, double W_width, double H_height)
        {

            double widthNum = x_2 / W_width;
            double heightNum = y_2 / H_height;
            //"SQRT((Geometry1.X10-Geometry1.X9)^2+(Geometry1.Y10-Geometry1.Y9)^2)*0.06"
            double distance = Math.Sqrt(Math.Pow(x_1 - x_2, 2) + Math.Pow(y_1 - y_2, 2));
            shape.AddRow((short)Visio.VisSectionIndices.visSectionFirstComponent, (short)(row_i), (short)Visio.VisRowTags.visTagArcTo);
            shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, (short)(row_i), 0].FormulaU = "\"" + "Width*" + widthNum + "\"";
            shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, (short)(row_i), 1].FormulaU = "\"" + "Height*" + heightNum + "\"";
            shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, (short)(row_i), 2].FormulaU = "\"" + distance * 0.5 + "\"";
            x_1 = x_2;
            y_1 = y_2;
            //shape.CellsU["FillForegnd"].FormulaU = "RGB(255, 255, 255)";  // 白色
            //shape.CellsU["FillBkgnd"].FormulaU = "RGB(255, 255, 255)";  // 白色
            shape.CellsU["LineColor"].FormulaU = "RGB(255, 0, 0)";  // 红色
            shape.CellsU["FillPattern"].FormulaU = "0";  // 无填充
        }
        static private double getShapeValue(Visio.Shape shape, string propName)
        {
            double value = shape.get_Cells(propName).ResultIU;
            return value;
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
