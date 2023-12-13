using System;

class MainClass
{
    public static void Main(string[] args)
    {
        double epsilon = 1e-2;  // 10的负5次方
        // 创建一个动态数组用于保存线上的点坐标
        List<Tuple<double, double>> LineStruPoints = new List<Tuple<double, double>>();
        List<Tuple<double, double>> LinePoints = new List<Tuple<double, double>>();
        List<Tuple<double, double>> LinePointsSum = new List<Tuple<double, double>>();
        LineStruPoints.Add(new Tuple<double, double>(0, -0.1));
        LineStruPoints.Add(new Tuple<double, double>(0.7693, 0.2559));
        LineStruPoints.Add(new Tuple<double, double>(0.5693, 0.2559));
        LineStruPoints.Add(new Tuple<double, double>(0.5693, 0.0382));
        LineStruPoints.Add(new Tuple<double, double>(0, 0.1));
        LineStruPoints.Add(new Tuple<double, double>(0, -0.1));
        //LineStruPoints.Add(new Tuple<double, double>(6, 3));
        double x1 = LineStruPoints[0].Item1;
        double y1 = LineStruPoints[0].Item2;
        foreach (var point in LineStruPoints)
        {
            double x2 = point.Item1;
            double y2 = point.Item2;
            if (Math.Abs(x2 - x1) < epsilon && Math.Abs(y2 - y1) < epsilon)
            {
                continue;
            }
            LinePoints = GetInterPonits(ref x1, x2,ref y1, y2);
            LinePointsSum.AddRange(LinePoints);
          
        }
        foreach (var point in LinePointsSum)
        {
            Console.WriteLine($"({point.Item1}, {point.Item2})");
        }
    }
    public static List<Tuple<double, double>> GetInterPonits(ref double x1, double x2, ref double y1, double y2)
    {
        double doubleerval = .1;   // 间隔 

        // 计算两个点的距离
        double distance = Math.Sqrt(Math.Pow(x2 - x1, 2) + Math.Pow(y2 - y1, 2));

        // 计算在距离上的间隔个数
        double count = Convert.ToInt32(distance / doubleerval);

        // 计算两个点的斜率
        double dX = x2 - x1;
        double dY = y2 - y1;

        List<Tuple<double, double>> Points1 = new List<Tuple<double, double>>();
        double K = dY / dX;
        if (x1 < x2 || y1 < y2)
        {
            //double x1 = 1, y1 = 1; // 第一个点的坐标
            //double x2 = 5, y2 = 4; // 第二个点的坐标
            
            if (double.IsInfinity(K))
            {
                // 生成每个间隔的坐标
                for (double y = y1; y <= y2; y += doubleerval)
                {
                    
                    double x = x1;
                    Points1.Add(new Tuple<double, double>(x, y));
                    Console.WriteLine($"{x}, {y}");
                }
                x1 = x2;
                y1 = y2;
                return Points1;
            }
            else
            {
                // 生成每个间隔的坐标
                for (double x = x1; x <= x2; x += doubleerval)
                {
                    double b = y1 - K * x1;
                    double y = x * K + b;
                    Points1.Add(new Tuple<double, double>(x, y));
                    Console.WriteLine($"{x}, {y}");
                }
                x1 = x2;
                y1 = y2;
                return Points1;
            }
        }
        else
        {
            if (double.IsInfinity(K))
            {
                // 生成每个间隔的坐标
                for (double y =y1 ; y >= y2; y -= doubleerval)
                {
                   
                    double x = x1;
                    Points1.Add(new Tuple<double, double>(x, y));
                    Console.WriteLine($"{x}, {y}");
                }
                x1 = x2;
                y1 = y2;
                return Points1;
            }
            else
            {
                // 生成每个间隔的坐标
                for (double x = x1; x >= x2; x -= doubleerval)
                {
                    double b = y1 - K * x1;
                    double y = x * K + b;
                    Points1.Add(new Tuple<double, double>(x, y));
                    Console.WriteLine($"{x}, {y}");
                }
                x1 = x2;
                y1 = y2;
                return Points1;
            }
        }         
    }
}