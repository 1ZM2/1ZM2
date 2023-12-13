using Microsoft.Office.Interop.Visio;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioAddIn_Test
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void 云图_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("start");
            // 获取当前Visio应用程序实例
            Visio.Application visioApp = Globals.ThisAddIn.Application;

            // 获取当前文档
            Visio.Document doc = visioApp.ActiveDocument;

            // 获取当前页面
            Visio.Page page = doc.Pages[1];
            foreach (Visio.Shape shape in page.Shapes)
            {//将高度乘以0.418/10.5
                double Smfactor = 0.98;
                double Bifactor = 1.02;
                //double changeFactor = 0.418 / 10.5;
                double X1 = (shape.get_Cells("PinX").ResultIU - shape.CellsU["Width"].ResultIU / 2) * Smfactor;
                double Y1 = (shape.get_Cells("PinY").ResultIU - shape.CellsU["Height"].ResultIU / 2) * Smfactor;
                double X2 = (shape.get_Cells("PinX").ResultIU + shape.CellsU["Width"].ResultIU / 2) * Bifactor;
                double Y2 = (shape.get_Cells("PinY").ResultIU + shape.CellsU["Height"].ResultIU / 2) * Bifactor;

                // 创建一个动态数组用于保存矩形上的点坐标
                List<Tuple<double, double>> rectanglePoints = new List<Tuple<double, double>>();

                // 计算矩形的宽和高
                double epsilon = 1e-5;  // 10的负5次方
                double width = Math.Abs(X2 - X1);
                double height = Math.Abs(Y2 - Y1);
                double gap = 0.1;
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


                // 创建一个新形状 - 矩形
                //Visio.Shape shapeYun = page.DrawCircularArc(X1, Y1, .00001);
                Visio.Shape shapeYun = page.DrawRectangle(X1, Y1, X2, Y2);

                //使用上述坐标画弧
                double x1 = 0;
                double y1 = 0;
                int rowi = 1;

                for (int i = 0; i < 4; i++)
                {
                    shapeYun.DeleteRow((short)VisSectionIndices.visSectionFirstComponent, 2);
                }
                Console.WriteLine("3");
                shapeYun.CellsSRC[(short)VisSectionIndices.visSectionFirstComponent, 1, 0].FormulaU = "\"" + 0 + "\"";
                shapeYun.CellsSRC[(short)VisSectionIndices.visSectionFirstComponent, 1, 1].FormulaU = "\"" + 0 + "\"";
                foreach (var point in rectanglePoints)
                {
                    //(point.Item1, point.Item2);

                    double x2 = point.Item1;
                    double y2 = point.Item2;
                    //shapeYun.CellsSRC[(short)VisSectionIndices.visSectionFirstComponent, 1, 0].ResultIU = x1;
                    //shapeYun.CellsSRC[(short)VisSectionIndices.visSectionFirstComponent, 1, 1].ResultIU = y1;
                    if (Math.Abs(x2 - x1) < epsilon && Math.Abs(y2 - y1) < epsilon)
                    {
                        continue;
                    }
                    if (x1 <= width && x2 <= width && Math.Abs(y2 - 0) < epsilon && Math.Abs(y1 - 0) < epsilon)
                    {//下
                        rowi++;

                        DrawMap(shapeYun, rowi, ref x1, x2, ref y1, y2);
                        continue;
                    }

                    if (Math.Abs(x1 - width) < epsilon && Math.Abs(x2 - width) < epsilon && y1 <= height && y2 <= height)
                    {//右
                        rowi++;

                        DrawMap(shapeYun, rowi, ref x1, x2, ref y1, y2);
                        continue;

                    }
                    if (x1 <= width && x2 <= width && Math.Abs(y1 - height) < epsilon && Math.Abs(y2 - height) < epsilon)
                    {//上
                        rowi++;

                        DrawMap(shapeYun, rowi, ref x1, x2, ref y1, y2);
                        continue;
                    }
                    if (Math.Abs(x1 - 0) < epsilon && Math.Abs(x2 - 0) < epsilon && y1 <= height && y2 <= height)
                    {//左
                        rowi++;

                        DrawMap(shapeYun, rowi, ref x1, x2, ref y1, y2);
                        continue;
                    }

                }
            }
            // 保存Visio文档
            //doc.Save();
            //doc.Close();
            //visioApp.Quit();

            // 清除选择
            visioApp.ActiveWindow.DeselectAll();
        }
        static public void DrawMap(Visio.Shape shape, int rowi, ref double x1, double x2, ref double y1, double y2)
        {

            // 创建一个新形状 - 圆弧                     
            //Shape shape = page.DrawArcByThreePoints(x1, y1, x2, y2, x3, y3);


            //"SQRT((Geometry1.X10-Geometry1.X9)^2+(Geometry1.Y10-Geometry1.Y9)^2)*0.06"
            double distance = Math.Sqrt(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2));
            shape.AddRow((short)VisSectionIndices.visSectionFirstComponent, (short)(rowi), (short)VisRowTags.visTagArcTo);
            shape.CellsSRC[(short)VisSectionIndices.visSectionFirstComponent, (short)(rowi), 0].FormulaU = "\"" + x2 + "\"";
            shape.CellsSRC[(short)VisSectionIndices.visSectionFirstComponent, (short)(rowi), 1].FormulaU = "\"" + y2 + "\"";
            shape.CellsSRC[(short)VisSectionIndices.visSectionFirstComponent, (short)(rowi), 2].FormulaU = "\"" + distance * 0.5 + "\"";

            x1 = x2;
            y1 = y2;
            // 设置填充为无颜色

            shape.CellsU["FillPattern"].FormulaU = "0";  // 无填充
            shape.CellsU["LineColor"].FormulaU = "RGB(255, 0, 0)";  // 红色
        }
    }      
}

    

            

