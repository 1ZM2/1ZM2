using System;
using System.Collections.Generic;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Visio;
using Visio=Microsoft.Office.Interop.Visio;
class Program
{
    static void Main(string[] args)
    {

        // 创建Visio应用程序对象
        Visio.Application visioApp = new Visio.Application();

        // 打开Visio文档
        Visio.Document doc = visioApp.Documents.Open(@"F:\CGNPT\C#云图\修改属性\绘图1.vsdx");

        Console.WriteLine("1");
        var page = doc.Pages[1];
        Visio.Shape shape = page.Shapes[1];

        //foreach (Visio.Shape shape in page.Shapes)
        //将高度乘以0.418/10.5
        //short IsExist = Math.Abs(shape.CellExists["Controls.TextPosition.X", (int)Visio.VisUnitCodes.visNumber]);
        short IsExist = Math.Abs(shape.CellExists["Width", (int)Visio.VisUnitCodes.visNumber]);
        if (IsExist == 1)
        {
            //double distance0 = 0.1;
            //short rowCount = shape.RowCount[(short)Visio.VisSectionIndices.visSectionFirstComponent];
            //// 假设三个点的坐标为 A(x1, y1), B(x2, y2), C(x3, y3)
            //Console.WriteLine($"函数坐标点行数：({rowCount})");
            //List<Tuple<double, double>> LinePoints0 = new List<Tuple<double, double>>();
            //for (short i = 1; i < rowCount; i++)
            //{
            //    double X = shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, i, 0].ResultIU;
            //    double Y = shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, i, 1].ResultIU;
            //    LinePoints0.Add(Tuple.Create(X, Y));
            //}
            //foreach (var point in LinePoints0)
            //{
            //    Console.WriteLine("VisioLine函数坐标点：({0}, {1})", point.Item1, point.Item2);
            //}
            ////识别到有geometry几个点 对于结构点*2
            //int geometryPoints = LinePoints0.Count * 2;
            ////按照框架点的个数 生成占位点
            //List<Tuple<double, double>> LinestucPoints = new List<Tuple<double, double>>();
            //for (int i = 0; i < geometryPoints; i++)
            //{
            //    LinestucPoints.Add(Tuple.Create(0.0, 0.0));
            //}
            //List<Tuple<double, double>> LinestucPoints1 = new List<Tuple<double, double>>();
            ////LinestucPoints1 = SubPoints(x1, y1, x2, y2, x3, y3);
            ////LinestucPoints[2] = LinestucPoints1[0];
            ////LinestucPoints[3] = LinestucPoints1[1];

            //for (int i = 0; i < LinePoints0.Count - 2; i++)
            //{
            //    //Console.WriteLine($"1");
            //    LinestucPoints1 = SubPoints(LinePoints0[i].Item1, LinePoints0[i].Item2, LinePoints0[i + 1].Item1, LinePoints0[i + 1].Item2, LinePoints0[i + 2].Item1, LinePoints0[i + 2].Item2);
            //    LinestucPoints[i + 1] = LinestucPoints1[0];
            //    LinestucPoints[geometryPoints - (i + 1) - 1] = LinestucPoints1[1];

            //}

            //if (LinePoints0[0].Item1 == LinePoints0[1].Item1)
            //{
            //    double y = LinePoints0[0].Item2;
            //    double x = LinePoints0[0].Item1 - distance0;
            //    //LineSubPonits.Insert(2,new Tuple<double, double>(x, y));
            //    LinestucPoints[0] = Tuple.Create(x, y);
            //    x = LinePoints0[0].Item1 + distance0;
            //    LinestucPoints[geometryPoints - 1] = Tuple.Create(x, y);

            //}
            //else
            //{
            //    double x = LinePoints0[0].Item1;
            //    double y = LinePoints0[0].Item2 - distance0;
            //    //LineSubPonits.Insert(2,new Tuple<double, double>(x, y));
            //    LinestucPoints[0] = Tuple.Create(x, y);
            //    y = LinePoints0[0].Item2 + distance0;
            //    LinestucPoints[geometryPoints - 1] = Tuple.Create(x, y);
            //}

            //if (LinePoints0[LinePoints0.Count - 2].Item1 == LinePoints0[LinePoints0.Count-1].Item1)
            //{
            //    double y = LinePoints0[LinePoints0.Count - 1].Item2;
            //    double x = LinePoints0[LinePoints0.Count - 1].Item1 + distance0;
            //    //LineSubPonits.Insert(2,new Tuple<double, double>(x, y));
            //    LinestucPoints[geometryPoints / 2 - 1] = Tuple.Create(x, y);
            //    x = LinePoints0[LinePoints0.Count-1].Item1 - distance0;
            //    LinestucPoints[geometryPoints / 2 ] = Tuple.Create(x, y);
            //}
            //else
            //{
            //    double x = LinePoints0[LinePoints0.Count - 1].Item1;
            //    double ya = LinePoints0[LinePoints0.Count - 1].Item2 + distance0;
            //    double yb = LinePoints0[LinePoints0.Count - 1].Item2 - distance0;
            //    if (Math.Abs(ya- LinestucPoints[geometryPoints / 2 - 2].Item2)< Math.Abs(yb - LinestucPoints[geometryPoints / 2 - 2].Item2) )
            //    {
            //        LinestucPoints[geometryPoints / 2 - 1] = Tuple.Create(x, ya);
            //        LinestucPoints[geometryPoints / 2] = Tuple.Create(x, yb);
            //    }
            //    else
            //    {
            //        LinestucPoints[geometryPoints / 2 - 1] = Tuple.Create(x, yb);
            //        LinestucPoints[geometryPoints / 2] = Tuple.Create(x, ya);
            //    }

            //    //LineSubPonits.Insert(2,new Tuple<double, double>(x, y));

            //}
            //// 复制第一个数据
            //Tuple<double, double> firstData = LinestucPoints[0];
            //// 在列表末尾插入复制的数据
            //LinestucPoints.Add(firstData);
            //// 打印排序后的列表          
            //foreach (var point in LinestucPoints)
            //{
            //    Console.WriteLine("{0}, {1}", point.Item1, point.Item2);
            //}
            List<Tuple<double, double>> LinestucPoints = new List<Tuple<double, double>>();
            LinestucPoints = GetLinestucPoints(shape);
            foreach (var point in LinestucPoints)
            {
                Console.WriteLine("{0}, {1}", point.Item1, point.Item2);
            }
            double X1 = (shape.get_Cells("PinX").ResultIU - shape.CellsU["Width"].ResultIU / 2);
            double Y1 = (shape.get_Cells("PinY").ResultIU - shape.CellsU["Height"].ResultIU / 2) ;
            double X2 = (shape.get_Cells("PinX").ResultIU + shape.CellsU["Width"].ResultIU / 2);
            double Y2 = (shape.get_Cells("PinY").ResultIU + shape.CellsU["Height"].ResultIU / 2) ;
            // 计算矩形的宽和高
            double epsilon = 1e-5;  // 10的负5次方
            double width = Math.Abs(X2 - X1);
            double height = Math.Abs(Y2 - Y1);

            // 创建一个动态数组用于保存矩形上的点坐标
            
            //得到矩形上的点坐标
            

            // 在选中图形位置上 创建并放置一个新形状 - 矩形                
            Visio.Shape shapeYun = page.DrawRectangle(X1, Y1, X2, Y2);
           
            //使用上述坐标画弧
            double x1 = 0;
            double y1 = 0;
            int rowi = 1;
            //删除矩形的2到5行的geometry值
            for (int i = 0; i < 5; i++)
            {
                shapeYun.DeleteRow((short)Visio.VisSectionIndices.visSectionFirstComponent, 2);
            }
            shapeYun.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, 1, 0].FormulaU = "\"" + LinestucPoints[0].Item1 + "\"";
            shapeYun.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, 1, 1].FormulaU = "\"" + LinestucPoints[0].Item2 + "\"";
            //画出前后两个点的弧形
            foreach (var point in LinestucPoints)
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
    private static void DrawArcMap(Visio.Shape shape, int row_i, ref double x_1, double x_2, ref double y_1, double y_2, double W_width, double H_height)
    {

        double widthNum = x_2 / W_width;
        double heightNum = y_2 / H_height;
        //"SQRT((Geometry1.X10-Geometry1.X9)^2+(Geometry1.Y10-Geometry1.Y9)^2)*0.06"
        double distance = Math.Sqrt(Math.Pow(x_1 - x_2, 2) + Math.Pow(y_1 - y_2, 2));
        //shape.AddRow((short)Visio.VisSectionIndices.visSectionFirstComponent, (short)(row_i), (short)Visio.VisRowTags.visTagArcTo);
        shape.AddRow((short)Visio.VisSectionIndices.visSectionFirstComponent, (short)(row_i), 193);
        shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, (short)(row_i), 0].FormulaU = "\"" + "Width*" + widthNum + "\"";
        shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, (short)(row_i), 1].FormulaU = "\"" + "Height*" + heightNum + "\"";
        //shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, (short)(row_i), 2].FormulaU = "\"" + distance * 0.5 + "\"";
        x_1 = x_2;
        y_1 = y_2;
        //shape.CellsU["FillForegnd"].FormulaU = "RGB(255, 255, 255)";  // 白色
        //shape.CellsU["FillBkgnd"].FormulaU = "RGB(255, 255, 255)";  // 白色
        shape.CellsU["LineColor"].FormulaU = "RGB(255, 0, 0)";  // 红色
        shape.CellsU["FillPattern"].FormulaU = "0";  // 无填充
    }
    public static List<Tuple<double, double>> SubPoints(double x1, double y1, double x2, double y2, double x3, double y3, double distance = 0.1)
    {
        // 步骤2：求边界BC的中点D
        double xD = (x1 + x3) / 2;
        double yD = (y1 + y3) / 2;
        double slope; // 斜率
        double intercept; // 截距
                          //识别到有geometry几个点 对于结构点*2
        int geometryPoints = 2;

        //按照框架点的个数 生成占位点
        List<Tuple<double, double>> LineSubPonits = new List<Tuple<double, double>>();
        for (int i = 0; i < geometryPoints; i++)
        {
            LineSubPonits.Add(Tuple.Create(0.0, 0.0));
        }


        //以中线为准 取点
        if (x2 == xD)
        {
            // 斜率不存在，直线垂直于 x 轴
            //Console.WriteLine($"直线方程为 x = {x2}");
            double x = x2;
            double y = y2 - 1;
            //LineSubPonits.Insert(2,new Tuple<double, double>(x, y));
            LineSubPonits[0] = Tuple.Create(x, y);
            y = y2 + 1;
            //LineSubPonits.Insert(3, new Tuple<double, double>(x, y));
            LineSubPonits[1] = Tuple.Create(x, y);

            return LineSubPonits;
        }
        else
        {

            // 计算斜率和截距
            slope = (yD - y2) / (xD - x2);
            intercept = y2 - slope * x2;

            Console.WriteLine($"直线方程为 y = {slope}x + {intercept}");


            for (double x = x2 - 2; x < x2 + 2; x += 0.3)
            {
                if (Math.Pow(x - x2, 2) + Math.Pow(slope * x + intercept - y2, 2) - distance < 0.05)
                {
                    //x = x1 + Math.Sqrt(distance - Math.Pow(slope * x + intercept - y1, 2));
                    //LineSubPonits.Insert(2, new Tuple<double, double>(x, slope * x + intercept));
                    LineSubPonits[0] = Tuple.Create(x, slope * x + intercept);
                    //Tuple<double, double> point1 = Tuple.Create(x, slope * x + intercept);
                }

            }
            for (double x = x2 + 2; x > x2 - 2; x -= 0.3)
            {
                if (Math.Pow(x - x2, 2) + Math.Pow(slope * x + intercept - y2, 2) - distance < 0.05)
                {
                    //x = x1 + Math.Sqrt(distance - Math.Pow(slope * x + intercept - y1, 2));
                    LineSubPonits[1] = Tuple.Create(x, slope * x + intercept);
                    //Tuple<double, double> point1 = Tuple.Create(x, slope * x + intercept);
                }

            }

            // 提取第二个点及其 x 和 y 坐标
            double dis1 = Math.Pow(x1 - LineSubPonits[0].Item1, 2) + Math.Pow(y1 - LineSubPonits[0].Item2, 2)+ Math.Pow(x3 - LineSubPonits[0].Item1, 2) + Math.Pow(y3 - LineSubPonits[0].Item2, 2);
            double dis2 = Math.Pow(x1 - LineSubPonits[1].Item1, 2) + Math.Pow(y1 - LineSubPonits[1].Item2, 2)+ Math.Pow(x3 - LineSubPonits[1].Item1, 2) + Math.Pow(y3 - LineSubPonits[1].Item2, 2);
            if (dis1 < dis2)
            {
                Tuple<double, double> temp = LineSubPonits[0];
                LineSubPonits[0] = LineSubPonits[1];
                LineSubPonits[1] = temp;
            }

            // 打印排序后的列表
            foreach (var point in LineSubPonits)
            {
                Console.WriteLine("子函数坐标点：({0}, {1})", point.Item1, point.Item2);
            }
            return LineSubPonits;
        }
    }
    public static void DrawMap(Visio.Shape shape, int rowi, ref double x1, double x2, ref double y1, double y2, double width, double height)
    {

        // 创建一个新形状 - 圆弧                     
        //Shape shape = page.DrawArcByThreePoints(x1, y1, x2, y2, x3, y3);

        double widthNum = x2 / width;
        double heightNum = y2 / height;
        int beforRowi = rowi - 1;
        //"SQRT((Geometry1.X10-Geometry1.X9)^2+(Geometry1.Y10-Geometry1.Y9)^2)*0.06"
        //double distance = Math.Sqrt(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2));
        shape.AddRow((short)VisSectionIndices.visSectionFirstComponent, (short)(rowi), (short)VisRowTags.visTagArcTo);
        shape.CellsSRC[(short)VisSectionIndices.visSectionFirstComponent, (short)(rowi), 0].FormulaU = "\"" + "Width*" + widthNum + "\"";
        shape.CellsSRC[(short)VisSectionIndices.visSectionFirstComponent, (short)(rowi), 1].FormulaU = "\"" + "Height*" + heightNum + "\"";
        shape.CellsSRC[(short)VisSectionIndices.visSectionFirstComponent, (short)(rowi), 2].FormulaU = "\"" + "SQRT((Geometry1.X" + rowi + "-Geometry1.X" + beforRowi + ")^2+(Geometry1.Y" + rowi + "-Geometry1.Y" + beforRowi + ")^2)*0.5" + "\"";

        x1 = x2;
        y1 = y2;
        //shape.CellsU["FillForegnd"].FormulaU = "RGB(255, 255, 255)";  // 白色
        //shape.CellsU["FillBkgnd"].FormulaU = "RGB(255, 255, 255)";  // 白色
        shape.CellsU["LineColor"].FormulaU = "RGB(255, 0, 0)";  // 红色
        shape.CellsU["FillPattern"].FormulaU = "0";  // 无填充

    }
    public static List<Tuple<double, double>> GetLinestucPoints(Visio.Shape shape, double distance0=0.1)
    {
        //distance0 = 0.1;
        short rowCount = shape.RowCount[(short)Visio.VisSectionIndices.visSectionFirstComponent];
        // 假设三个点的坐标为 A(x1, y1), B(x2, y2), C(x3, y3)
        Console.WriteLine($"函数坐标点行数：({rowCount})");
        List<Tuple<double, double>> LinePoints0 = new List<Tuple<double, double>>();
        for (short i = 1; i < rowCount; i++)
        {
            double X = shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, i, 0].ResultIU;
            double Y = shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, i, 1].ResultIU;
            LinePoints0.Add(Tuple.Create(X, Y));
        }
        foreach (var point in LinePoints0)
        {
            Console.WriteLine("VisioLine函数坐标点：({0}, {1})", point.Item1, point.Item2);
        }
        //识别到有geometry几个点 对于结构点*2
        int geometryPoints = LinePoints0.Count * 2;
        //按照框架点的个数 生成占位点
        List<Tuple<double, double>> LinestucPoints = new List<Tuple<double, double>>();
        for (int i = 0; i < geometryPoints; i++)
        {
            LinestucPoints.Add(Tuple.Create(0.0, 0.0));
        }
        List<Tuple<double, double>> LinestucPoints1 = new List<Tuple<double, double>>();
        //LinestucPoints1 = SubPoints(x1, y1, x2, y2, x3, y3);
        //LinestucPoints[2] = LinestucPoints1[0];
        //LinestucPoints[3] = LinestucPoints1[1];
        //中间点处理
        for (int i = 0; i < LinePoints0.Count - 2; i++)
        {
            //Console.WriteLine($"1");
            LinestucPoints1 = SubPoints(LinePoints0[i].Item1, LinePoints0[i].Item2, LinePoints0[i + 1].Item1, LinePoints0[i + 1].Item2, LinePoints0[i + 2].Item1, LinePoints0[i + 2].Item2);
            LinestucPoints[i + 1] = LinestucPoints1[0];
            LinestucPoints[geometryPoints - (i + 1) - 1] = LinestucPoints1[1];

        }
        //首尾点处理之首点
        if (LinePoints0[0].Item1 == LinePoints0[1].Item1)
        {
            double y = LinePoints0[0].Item2;
            double x = LinePoints0[0].Item1 - distance0;
            //LineSubPonits.Insert(2,new Tuple<double, double>(x, y));
            LinestucPoints[0] = Tuple.Create(x, y);
            x = LinePoints0[0].Item1 + distance0;
            LinestucPoints[geometryPoints - 1] = Tuple.Create(x, y);

        }
        else
        {
            double x = LinePoints0[0].Item1;
            double y = LinePoints0[0].Item2 + distance0;
            //LineSubPonits.Insert(2,new Tuple<double, double>(x, y));
            LinestucPoints[0] = Tuple.Create(x, y);
            y = LinePoints0[0].Item2 - distance0;
            LinestucPoints[geometryPoints - 1] = Tuple.Create(x, y);
        }
        //首尾点处理之尾点
        if (LinePoints0[LinePoints0.Count - 2].Item1 == LinePoints0[LinePoints0.Count - 1].Item1)
        {
            double y = LinePoints0[LinePoints0.Count - 1].Item2;
            double x = LinePoints0[LinePoints0.Count - 1].Item1 + distance0;
            //LineSubPonits.Insert(2,new Tuple<double, double>(x, y));
            LinestucPoints[geometryPoints / 2 - 1] = Tuple.Create(x, y);
            x = LinePoints0[LinePoints0.Count - 1].Item1 - distance0;
            LinestucPoints[geometryPoints / 2] = Tuple.Create(x, y);
        }
        else
        {
            double x = LinePoints0[LinePoints0.Count - 1].Item1;
            double ya = LinePoints0[LinePoints0.Count - 1].Item2 + distance0;
            double yb = LinePoints0[LinePoints0.Count - 1].Item2 - distance0;
            if (Math.Abs(ya - LinestucPoints[geometryPoints / 2 - 2].Item2) < Math.Abs(yb - LinestucPoints[geometryPoints / 2 - 2].Item2))
            {
                LinestucPoints[geometryPoints / 2 - 1] = Tuple.Create(x, ya);
                LinestucPoints[geometryPoints / 2] = Tuple.Create(x, yb);
            }
            else
            {
                LinestucPoints[geometryPoints / 2 - 1] = Tuple.Create(x, yb);
                LinestucPoints[geometryPoints / 2] = Tuple.Create(x, ya);
            }

            //LineSubPonits.Insert(2,new Tuple<double, double>(x, y));

        }
        // 复制第一个数据
        Tuple<double, double> firstData = LinestucPoints[0];
        // 在列表末尾插入复制的数据
        LinestucPoints.Add(firstData);
        return LinestucPoints;
        // 打印排序后的列表          
 
    }
}







