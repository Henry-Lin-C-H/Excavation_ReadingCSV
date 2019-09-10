using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace csv_reader
{
    public class csv_read
    {
        public double Wall_length = 0;
        public double Wall_width = 0;
        public double depth = 0;
        public List<Tuple<double, double>> excaRange = new List<Tuple<double, double>>(); //X, Y
        public List<Tuple<int, double>> excaLevel = new List<Tuple<int, double>>(); // number of excavated stages, depth
        public List<Tuple<string, double, int, string, int>> supLevel = new List<Tuple<string, double, int, string, int>>();
        //number of sturs' stage, depth, number of struts, struts' type, spacing

        public List<Tuple<double, double, double, double, string, double>> centralCol = new List<Tuple<double, double, double, double, string, double>>();
        //central collumms' spacing, length, depth, diameter, type, length of H beams

        public List<Tuple<string, double, double>> sidewall = new List<Tuple<string, double, double>>();
        //sidewalls' floor, thickness,  strength of concrete

        public List<Tuple<string, double, double, double>> back = new List<Tuple<string, double, double, double>>();
        //Backfill floor, depth, thickness, strength of concrete

        public void Main()
        {
            string filePath = @"E:\csv\20180604-BIM資料格式(LG05).csv";

            List<string> lines = File.ReadAllLines(filePath, Encoding.GetEncoding("Big5")).ToList();

            for (int i = 0; i < 55; i++)
            {
                List<string> Line = lines[i].Split(',').ToList();
                if (Line[0] == "擋土壁長度：")
                {
                    double.TryParse(Line[1], out Wall_length);
                }
                else if (Line[0] == "連續壁厚度：")
                {
                    double.TryParse(Line[1], out Wall_width);
                }
                else if (Line[0] == "開挖深度：")
                {
                    double.TryParse(Line[1], out depth);
                }
                else if (Line[0] == "開挖範圍")
                {
                    int j = i + 1;
                    do
                    {
                        List<string> xy = lines[j].Split(',').ToList();
                        if (xy[1].Trim() == "") break;
                        double t1 = 0;
                        double t2 = 0;
                        double.TryParse(xy[1], out t1);
                        double.TryParse(xy[2], out t2);
                        var data = Tuple.Create(t1, t2);
                        excaRange.Add(data);

                        j++;
                    } while (true);
                }
                else if (Line[0] == "開挖階數：")
                {
                    int j = i + 1;
                    do
                    {
                        List<string> xy = lines[j].Split(',').ToList();
                        if (xy[1].Trim() == "") break;
                        int t1 = 0;
                        double t2 = 0;
                        int.TryParse(xy[1], out t1);
                        double.TryParse(xy[2], out t2);
                        var data = Tuple.Create(t1, t2);
                        excaLevel.Add(data);

                        j++;
                    } while (true);
                }
                else if (Line[0] == "支撐階數：")
                {
                    int j = i + 1;
                    do
                    {
                        List<string> xy = lines[j].Split(',').ToList();
                        if (xy[1].Trim() == "") break;
                        int t3, t5;
                        double t2 = 0;
                        string t1, t4;
                        double.TryParse(xy[2], out t2);
                        int.TryParse(xy[3], out t3);
                        t1 = xy[1].ToString();
                        t4 = xy[4].ToString();
                        int.TryParse(xy[5], out t5);
                        var data = Tuple.Create(t1, t2, t3, t4, t5);
                        supLevel.Add(data);

                        j++;
                    } while (true);
                }
                else if (Line[0] == "中間柱：")
                {
                    int j = i + 1;
                    do
                    {
                        List<string> xy = lines[j].Split(',').ToList();
                        if (xy[1].Trim() == "") break;
                        double t1, t2, t3, t4, t6;
                        string t5;

                        double.TryParse(xy[1], out t1);
                        double.TryParse(xy[2], out t2);
                        double.TryParse(xy[3], out t3);
                        double.TryParse(xy[4], out t4);
                        double.TryParse(xy[6], out t6);
                        t5 = xy[5];
                        var data = Tuple.Create(t1, t2, t3, t4, t5, t6);

                        centralCol.Add(data);

                        j++;
                    } while (true);
                }
                else if (Line[0] == "側牆：")
                {
                    int j = i + 1;
                    do
                    {
                        List<string> xy = lines[j].Split(',').ToList();
                        if (xy[1].Trim() == "") break;
                        string t1;
                        double t2, t3;

                        t1 = xy[1];
                        double.TryParse(xy[2], out t2);
                        double.TryParse(xy[3], out t3);
                        var data = Tuple.Create(t1, t2, t3);
                        sidewall.Add(data);

                        j++;
                    } while (true);
                }
                else if (Line[0] == "樓板回築&回填：")
                {
                    int j = i + 1;
                    do
                    {
                        List<string> xy = lines[j].Split(',').ToList();
                        if (xy[1].Trim() == "") break;
                        string t1;
                        double t2, t3, t4;

                        t1 = xy[1];
                        double.TryParse(xy[2], out t2);
                        double.TryParse(xy[3], out t3);
                        double.TryParse(xy[4], out t4);
                        var data = Tuple.Create(t1, t2, t3, t4);
                        back.Add(data);

                        j++;
                    } while (true);
                }
            }
        }
    }
}