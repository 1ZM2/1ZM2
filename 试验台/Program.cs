using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Claims;

class Program
{
    static void Main(string[] args)
    {
        List<int> lst = new List<int> { 15, 6, 10, 2, 7,1, 2,  3, 5, 4,
                                         1, 1, 12, 1 ,4,15, 6, 10, 2, 7,
                                         1, 2,  3, 5, 4,1, 1, 12, 1 ,4,
                                         15, 6, 10, 2, 7, 1, 2,  3, 5, 4,
                                         1, 1, 12, 1 ,4};
        int groupSize = 40;
        int IndexCount = 2;
        int oneFill = groupSize * IndexCount - lst.Count;
        lst.AddRange(Enumerable.Repeat(1, oneFill));
        

        List<int> newLst = new List<int>();

        for (int i = 0; i < lst.Count; i += groupSize)
        {
            List<int> group = lst.Skip(i).Take(groupSize).ToList();
            int maxNum = group.Max();
            newLst.Add(maxNum);
        }

        Console.WriteLine(string.Join(", ", newLst));  // 输出: 15, 10, 13, 16, 18
        
        List<T> myTList = new List<T>
        {
            new T{Version="A"},
            new T{Version="G"},
            new T{Version="Z"},
        }; // 假设你已经有一个T类的列表

        List<string> versionList = myTList.Select(t => t.Version).ToList();

        //List<string> versionList = new List<string> { "A", "A", "AB", "A", "Z" };
        List<int> columnList = new List<int>();

        foreach (string str in versionList)
        {
            int column = GetColumnNumber(str);
            columnList.Add(column);
        }
        Console.WriteLine(columnList[2]);
    }
    public static int GetColumnNumber(string columnName)
    {
        int columnNumber = 0;
        columnName = columnName.ToUpper(); // 将输入的字符串转换为大写

        for (int i = 0; i < columnName.Length; i++)
        {
            char c = columnName[i];
            columnNumber = columnNumber * 26 + (c - 'A' + 1);
        }

        return columnNumber;
    }
    public static List<int> GetMaxVersionList(List<int> lst, int groupSize, int IndexCount)
    {
        
      
        int oneFill = groupSize * IndexCount - lst.Count;
        lst.AddRange(Enumerable.Repeat(1, oneFill));


        List<int> newLst = new List<int>();

        for (int i = 0; i < lst.Count; i += groupSize)
        {
            List<int> group = lst.Skip(i).Take(groupSize).ToList();
            int maxNum = group.Max();
            newLst.Add(maxNum);
        }
        return newLst;
    }

}
class T
{
    public string Version{get; set; }
}