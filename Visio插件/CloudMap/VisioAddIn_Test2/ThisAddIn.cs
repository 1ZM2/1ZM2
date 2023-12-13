using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Visio;
using System.Windows.Forms;


namespace VisioAddIn_Test2
{
    public static class GlobalVariables
    {
        public static int executeCount = 0;
    }
    public partial class ThisAddIn
    {
        private Ribbon1 ribbon;
        private Ribbon1 ribbon1;
        private bool isKeyPressed = false;
        private bool isS_Pressed = false;
        private bool isCtrlKeyHasDown = false;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 订阅图形选择事件
            this.Application.SelectionChanged += Application_SelectionChanged;
            // 订阅键盘按键事件            
            this.Application.KeyUp += Application_KeyUp;
            this.Application.KeyUp += Application_KeyUp1;
            this.Application.KeyUp += Application_KeyUp2;
            this.Application.KeyDown += Application_KeyDown;
            //订阅Ribbon事件
            // 初始化 ribbon 对象
            ribbon = Globals.Ribbons.Ribbon1;
            ribbon1 = Globals.Ribbons.Ribbon1;
            // 订阅 MyCustomEvent 事件
            ribbon.MyCustomEvent += ribbon_MyCustomEvent;
            ribbon1.MyCustomEvent1 += ribbon_MyCustomEvent1;
            
            //this.Application.KeyPress += Application_KeyPress;            
        }


        private void ribbon_MyCustomEvent1(object sender, EventArgs e)
        {
            MessageBox.Show("已关闭修订模式Button");
            isS_Pressed = false;
        }
        private void ribbon_MyCustomEvent(object sender, EventArgs e)
        {
            MessageBox.Show("已启动修订模式Button");
            isS_Pressed = true;
        }


        private void Application_KeyDown(int KeyCode, int KeyButtonState, ref bool CancelDefault)
        {
            if (KeyCode == 17)
            {
                //MessageBox.Show("已按下ctrl");
                isCtrlKeyHasDown = true;
                CancelDefault = true; // 阻止默认操作
            }
        }

        private void Application_KeyUp(int KeyCode, int KeyButtonState, ref bool CancelDefault)
        {              //r is 82、 crtl is 17     F4 键的键码是 115
            if (KeyCode == 115)
            {
                //MessageBox.Show("画云图Key");
                isKeyPressed = true;
                CancelDefault = true; // 阻止默认操作
                if (isKeyPressed && HasApplication_Selection)
                {
                    //MessageBox.Show("键被按下，并且有图形被选中4。");
                    GlobalVariables.executeCount++;
                    ThisAddIn.DrawYunMap(0.98, 1.02, 0.1);
                    isKeyPressed = false;

                }
            }
        }
        private void Application_KeyUp1(int KeyCode, int KeyButtonState, ref bool CancelDefault)
        {   //r is 82、 crtl is 17、 s id 83 F4 键的键码是 115                 
            if (KeyCode == 82 && isCtrlKeyHasDown)
            {
                MessageBox.Show("已启动修订模式Key");
                isS_Pressed = true;
                CancelDefault = true; // 阻止默认操作
                isCtrlKeyHasDown = false;
            }            
        }
        private void Application_KeyUp2(int KeyCode, int KeyButtonState, ref bool CancelDefault)
        {   //r is 82 crtl is 17 s id 83 Esc is 27
            if (KeyCode == 27)
            {
                MessageBox.Show("退出云图");
                isS_Pressed = false;
            }
        }
        private bool HasApplication_Selection = false;
        public void Application_SelectionChanged(Visio.Window window)
        {

            //if (window.Selection.Count > 0 && isKeyPressed && isS_Pressed)
            //{
            //    //if (isKeyPressed && isS_Pressed)
            //    {
            //        //MessageBox.Show("键被按下，并且有图形被选中4。");
            //        GlobalVariables.executeCount++;
            //        ThisAddIn.DrawYunMap(0.98, 1.02, 0.1);
            //        isKeyPressed = false;
            //        window.DeselectAll();
            //    }
            //}
            if (window.Selection.Count > 0 && isS_Pressed)
            {   MessageBox.Show("并且有图形被选中4。");
                HasApplication_Selection = true;
            }
            //window.DeselectAll();
            // 当图形选择发生变化时触发                        
        }
        static private void DrawYunMap(double Smfactor, double Bifactor, double gap)
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
            int shapeCount = selectedShapes.Count;
            if (shapeCount >= 2)
            {
                List<Tuple<double, double>> SumPoints0 = new List<Tuple<double, double>>();
                foreach (Visio.Shape shape in selectedShapes)
                {
                   
                    short IsExist = Math.Abs(shape.CellExists["Controls.TextPosition.X", (int)Visio.VisUnitCodes.visNumber]);
                    if (IsExist == 1)
                    {
                        //short rowCount = shape.RowCount[(short)Visio.VisSectionIndices.visSectionFirstComponent];
                        //// 假设三个点的坐标为 A(x1, y1), B(x2, y2), C(x3, y3)                       
                        //for (short i = 0; i < rowCount; i++)
                        //{
                        //    double X = shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, i, 0].ResultIU;
                        //    double Y = shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, i, 0].ResultIU;
                        //    SumPoints0.Add(Tuple.Create(X, Y));
                        //}
                        double X = shape.get_Cells("BeginX").ResultIU;
                        double Y = shape.get_Cells("BeginY").ResultIU;
                        SumPoints0.Add(Tuple.Create(X, Y));
                        X = shape.get_Cells("EndX").ResultIU;
                        Y = shape.get_Cells("EndY").ResultIU;
                        SumPoints0.Add(Tuple.Create(X, Y));

                    }
                    else
                    {
                        double X = (shape.get_Cells("PinX").ResultIU - shape.CellsU["Width"].ResultIU / 2) ;
                        double Y = (shape.get_Cells("PinY").ResultIU - shape.CellsU["Height"].ResultIU / 2) ;
                        SumPoints0.Add(Tuple.Create(X, Y));
                        X = (shape.get_Cells("PinX").ResultIU + shape.CellsU["Width"].ResultIU / 2) ;
                        Y = (shape.get_Cells("PinY").ResultIU + shape.CellsU["Height"].ResultIU / 2);
                        SumPoints0.Add(Tuple.Create(X, Y));
                    }
                }
                // 查找最小和最大 X 坐标值
                double minX = SumPoints0.Min(point => point.Item1);
                double maxX = SumPoints0.Max(point => point.Item1);

                // 查找最小和最大 Y 坐标值
                double minY = SumPoints0.Min(point => point.Item2);
                double maxY = SumPoints0.Max(point => point.Item2);

                
                 MessageBox.Show("图形");
                double X1 = minX;
                double Y1 = minY;
                double X2 = maxX;
                double Y2 = maxY;
                // 计算矩形的宽和高
                double epsilon = 1e-5;  // 10的负5次方
                double width = Math.Abs(X2 - X1);
                double height = Math.Abs(Y2 - Y1);

                // 创建一个动态数组用于保存矩形上的点坐标
                List<Tuple<double, double>> rectanglePoints = new List<Tuple<double, double>>();
                //得到矩形上的点坐标
                GetRectanglePoints(X1, Y1, X2, Y2, gap, rectanglePoints);

                // 在选中图形位置上 创建并放置一个新形状 - 矩形                
                Visio.Shape shapeYun = page.DrawRectangle(X1, Y1, X2, Y2);
                shapeYun.Name = GlobalVariables.executeCount.ToString("0000");
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
                //画出前后两个点的弧形
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
            if (shapeCount == 1)
            {
                foreach (Visio.Shape shape in selectedShapes)
                //foreach (Visio.Shape shape in page.Shapes)
                {//将高度乘以0.418/10.5

                    short IsExist = Math.Abs(shape.CellExists["Controls.TextPosition.X", (int)Visio.VisUnitCodes.visNumber]);
                    if (IsExist == 1)
                    {
                        MessageBox.Show("线条");
                        short rowCount = shape.RowCount[(short)Visio.VisSectionIndices.visSectionFirstComponent];
                        // 假设三个点的坐标为 A(x1, y1), B(x2, y2), C(x3, y3)
                        List<Tuple<double, double>> LinePoints0 = new List<Tuple<double, double>>();
                        for (short i = 0; i < rowCount; i++)
                        {
                            double X = shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, i, 0].ResultIU;
                            double Y = shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, i, 0].ResultIU;
                            LinePoints0.Add(Tuple.Create(X, Y));
                        }



                        double X1 = shape.get_Cells("BeginX").ResultIU * Smfactor;
                        double Y1 = shape.get_Cells("BeginY").ResultIU * Smfactor;
                        double X2 = shape.get_Cells("EndX").ResultIU * Bifactor;
                        double Y2 = shape.get_Cells("EndY").ResultIU * Bifactor;
                        // 计算矩形的宽和高
                        double epsilon = 1e-5;  // 10的负5次方
                        double width = Math.Abs(X2 - X1);
                        double height = Math.Abs(Y2 - Y1);

                        // 创建一个动态数组用于保存矩形上的点坐标
                        List<Tuple<double, double>> rectanglePoints = new List<Tuple<double, double>>();
                        //得到矩形上的点坐标
                        GetRectanglePoints(X1, Y1, X2, Y2, gap, rectanglePoints);


                        int undoScopeID = visioApp.BeginUndoScope("Create and Modify Shape");
                        // 在选中图形位置上 创建并放置一个新形状 - 矩形                
                        Visio.Shape shapeYun = page.DrawRectangle(X1, Y1, X2, Y2);
                        shapeYun.Name = GlobalVariables.executeCount.ToString("0000");
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
                        //画出前后两个点的弧形
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
                        // 结束撤销操作范围
                        visioApp.EndUndoScope(undoScopeID, true);
                    }
                    else
                    {
                        MessageBox.Show("图形");
                        double X1 = (shape.get_Cells("PinX").ResultIU - shape.CellsU["Width"].ResultIU / 2) * Smfactor;
                        double Y1 = (shape.get_Cells("PinY").ResultIU - shape.CellsU["Height"].ResultIU / 2) * Smfactor;
                        double X2 = (shape.get_Cells("PinX").ResultIU + shape.CellsU["Width"].ResultIU / 2) * Bifactor;
                        double Y2 = (shape.get_Cells("PinY").ResultIU + shape.CellsU["Height"].ResultIU / 2) * Bifactor;
                        // 计算矩形的宽和高
                        double epsilon = 1e-5;  // 10的负5次方
                        double width = Math.Abs(X2 - X1);
                        double height = Math.Abs(Y2 - Y1);

                        // 创建一个动态数组用于保存矩形上的点坐标
                        List<Tuple<double, double>> rectanglePoints = new List<Tuple<double, double>>();
                        //得到矩形上的点坐标
                        GetRectanglePoints(X1, Y1, X2, Y2, gap, rectanglePoints);

                        // 在选中图形位置上 创建并放置一个新形状 - 矩形                
                        Visio.Shape shapeYun = page.DrawRectangle(X1, Y1, X2, Y2);
                        shapeYun.Name = GlobalVariables.executeCount.ToString("0000");
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
                        //画出前后两个点的弧形
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

                    }

                }
                visioApp.Window.DeselectAll();
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

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 释放订阅图形选择事件
            this.Application.SelectionChanged -= Application_SelectionChanged;
            // 释放订阅键盘按键事件
            this.Application.KeyUp -= Application_KeyUp;
            this.Application.KeyUp -= Application_KeyUp1;
            this.Application.KeyUp -= Application_KeyUp2;
            ribbon.MyCustomEvent -= ribbon_MyCustomEvent;
            ribbon.MyCustomEvent -= ribbon_MyCustomEvent1;
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
