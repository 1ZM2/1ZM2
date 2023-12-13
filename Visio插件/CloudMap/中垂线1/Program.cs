using System;
using static System.Formats.Asn1.AsnWriter;

namespace AngleBisector
{
    class Program
    {

        static void Main(string[] args)
        {
            // 假设三个点的坐标为 A(x1, y1), B(x2, y2), C(x3, y3)
            List<Tuple<double, double>> LinePoints0 = new List<Tuple<double, double>>();
            
            // 步骤1：选择顶点A，相邻点B和C
            double x1 = 1, y1 = 1;
            double x2 = 2, y2 = 3;
            double x3 = 5, y3 = 1;
            double x4 = 3, y4 = 6;
            double x5 = 6, y5 = 6;
            LinePoints0.Add(Tuple.Create(x1, y1));
            LinePoints0.Add(Tuple.Create(x2, y2));
            LinePoints0.Add(Tuple.Create(x3, y3));
            LinePoints0.Add(Tuple.Create(x4, y4));
            LinePoints0.Add(Tuple.Create(x5, y5));
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

            for (int i = 0; i < LinePoints0.Count - 2 ; i++)
            {
                Console.WriteLine($"1");
                LinestucPoints1 = SubPoints(LinePoints0[i].Item1, LinePoints0[i].Item2, LinePoints0[i + 1].Item1, LinePoints0[i + 1].Item2, LinePoints0[i + 2].Item1, LinePoints0[i + 2].Item2);
                LinestucPoints[i+1] = LinestucPoints1[0];
                LinestucPoints[geometryPoints -(i+1)-1] = LinestucPoints1[1];

            }

            if (x1==x2)
            {
                double y = y1;
                double x = x1 + 1;
                //LineSubPonits.Insert(2,new Tuple<double, double>(x, y));
                LinestucPoints[0] = Tuple.Create(x, y);
                x= x1 - 1;
                LinestucPoints[geometryPoints-1] = Tuple.Create(x, y);

            }
            else
            {
                double x = x1;
                double y = y1 - 1;
                //LineSubPonits.Insert(2,new Tuple<double, double>(x, y));
                LinestucPoints[0] = Tuple.Create(x, y);
                y = y1 + 1;
                LinestucPoints[geometryPoints - 1] = Tuple.Create(x, y);
            }

            if (x3==x4)
            {
                double y = y4;
                double x = x4 - 1;
                //LineSubPonits.Insert(2,new Tuple<double, double>(x, y));
                LinestucPoints[geometryPoints / 2 - 1] = Tuple.Create(x, y);
                x = x4 + 1;
                LinestucPoints[geometryPoints / 2 - 2] = Tuple.Create(x, y);
            }
            else    
            {
                double x = x4;
                double y = y4 - 1;
                //LineSubPonits.Insert(2,new Tuple<double, double>(x, y));
                LinestucPoints[geometryPoints / 2 - 1] = Tuple.Create(x, y);
                y = y4 + 1;
                LinestucPoints[geometryPoints / 2] = Tuple.Create(x, y);
            }

            // 打印排序后的列表
            foreach (var point in LinestucPoints)
            {
                Console.WriteLine("函数坐标点：({0}, {1})", point.Item1, point.Item2);
            }
        }
        public static List<Tuple<double, double>> SubPoints(double x1, double y1, double x2, double y2, double x3, double y3, double distance = 1)
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
                Console.WriteLine($"直线方程为 x = {x2}");
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
                    if (Math.Pow(x - x2, 2) + Math.Pow(slope * x + intercept - y2, 2) - distance < 0.1)
                    {
                        //x = x1 + Math.Sqrt(distance - Math.Pow(slope * x + intercept - y1, 2));
                        //LineSubPonits.Insert(2, new Tuple<double, double>(x, slope * x + intercept));
                        LineSubPonits[0] = Tuple.Create(x, slope * x + intercept);
                        Console.WriteLine($"{x}");
                        Console.WriteLine($"2.1");
                        //Tuple<double, double> point1 = Tuple.Create(x, slope * x + intercept);
                    }

                }
                for (double x = x2 + 2; x > x2 - 2; x -= 0.3)
                {
                    if (Math.Pow(x - x2, 2) + Math.Pow(slope * x + intercept - y2, 2) - distance < 0.1)
                    {
                        //x = x1 + Math.Sqrt(distance - Math.Pow(slope * x + intercept - y1, 2));
                        LineSubPonits[1] = Tuple.Create(x, slope * x + intercept);
                        //Tuple<double, double> point1 = Tuple.Create(x, slope * x + intercept);
                    }

                }
                
                // 提取第二个点及其 x 和 y 坐标
                Tuple<double, double> secondPoint = LineSubPonits[0];
                double y_second = secondPoint.Item2;
                Tuple<double, double> thirdPoint = LineSubPonits[1];
                double y_third = thirdPoint.Item2;

                if (y_second > y_third)
                {
                    Tuple<double, double> temp = secondPoint;
                    secondPoint = thirdPoint;
                    thirdPoint = temp;
                }

                // 打印排序后的列表
                foreach (var point in LineSubPonits)
                {
                    Console.WriteLine("子函数坐标点：({0}, {1})", point.Item1, point.Item2);
                }
                return LineSubPonits;
            }
        }
    }
}