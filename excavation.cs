using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using csv_reader;

namespace Sino_Excavation_net
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class excavation : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            Transaction trans = new Transaction(doc);
            trans.Start("交易開始");

            //訂定土體範圍
            List<XYZ> topxyz = new List<XYZ>();
            XYZ t1 = new XYZ(-200, -200, 0);
            XYZ t2 = new XYZ(200, -200, 0);
            XYZ t3 = new XYZ(200, 200, 0);
            XYZ t4 = new XYZ(-200, 200, 0);
            topxyz.Add(t1 * 1000 / 304.8);
            topxyz.Add(t2 * 1000 / 304.8);
            topxyz.Add(t3 * 1000 / 304.8);
            topxyz.Add(t4 * 1000 / 304.8);
            TopographySurface.Create(doc, topxyz);

            //開挖深度所需參數
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(BuildingPadType));
            BuildingPadType bdtp = collector.FirstElement() as BuildingPadType;

            csv_read csv = new csv_read();
            csv.Main();

            //開挖各階之深度輸入
            List<double> height = new List<double>();
            foreach (var data in csv.excaLevel)
                height.Add(data.Item2 * -1);

            //建立開挖階數
            Level[] levlist = new Level[height.Count()];
            for (int i = 0; i != height.Count(); i++)
            {
                levlist[i] = Level.Create(doc, height[i] * 1000 / 304.8);
                levlist[i].Name = "開挖階數" + (i + 1).ToString();
            }

            double wallLength = csv.Wall_length;
            Level DW_level = Level.Create(doc, wallLength * -1 * 1000 / 304.8);

            //建立回築樓層
            Level[] re_levlist = new Level[csv.back.Count()];
            for (int i = 0; i != csv.back.Count(); i++)
            {
                re_levlist[i] = Level.Create(doc, (csv.back[i].Item2 * -1 + csv.back[i].Item3 / (2)) * 1000 / 304.8);
                re_levlist[i].Name = (csv.back[i].Item1).ToString();
            }
            Level[] all_level_list = new Level[levlist.Count() + re_levlist.Count()];
            all_level_list = levlist.Concat(re_levlist).ToArray();

            //訂定開挖範圍
            IList<CurveLoop> profileloops = new List<CurveLoop>();
            IList<Curve> wall_profileloops = new List<Curve>();

            //須回到原點

            XYZ[] points = new XYZ[csv.excaRange.Count()];

            for (int i = 0; i != csv.excaRange.Count(); i++)
                points[i] = new XYZ(csv.excaRange[i].Item1, csv.excaRange[i].Item2, 0) * 1000 / 304.8;

            CurveLoop profileloop = new CurveLoop();
            for (int i = 0; i < points.Count() - 1; i++)
            {
                Line line = Line.CreateBound(points[i], points[i + 1]);
                wall_profileloops.Add(line);
                profileloop.Append(line);
            }
            profileloops.Add(profileloop);
            Level levdeep = null;

            //建立開挖深度
            ICollection<Level> level_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
            foreach (Level lev in level_familyinstance)
            {
                if (lev.Name == levlist[levlist.Count() - 1].Name)
                {
                    BuildingPad b = BuildingPad.Create(doc, bdtp.Id, lev.Id, profileloops);
                    levdeep = lev;
                }
            }

            //建立連續壁
            IList<Curve> inner_wall_curves = new List<Curve>();
            double wall_W = csv.Wall_width * 1000; //連續壁厚度
            List<Wall> inner_wall = new List<Wall>();
            ICollection<WallType> walltype_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().ToList();
            foreach (WallType walltype in walltype_familyinstance)
            {
                if (walltype.Name == "連續壁")
                {
                    for (int i = 0; i < points.Count<XYZ>() - 1; i++)
                    {
                        Curve c = wall_profileloops[i].CreateOffset(wall_W / 2 / 304.8, new XYZ(0, 0, -1)); //偏移連續壁厚度1/2的距離，做為建置參考線
                        Wall w = Wall.Create(doc, c, walltype.Id, DW_level.Id, wallLength * 1000 / 304.8, 0, false, false);
                        w.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("連續壁");
                        inner_wall.Add(w);
                    }
                }
            }

            trans.Commit();

            //取得連續壁內座標點
            XYZ[] innerwall_points = new XYZ[points.Count<XYZ>()];
            XYZ[] for_check_innerwall_points = new XYZ[points.Count<XYZ>()]; //計算支撐中間樁之擺放界線
            for (int i = 0; i < (inner_wall.Count<Wall>()); i++)
            {
                //inner
                XYZ wall_curve_point = (inner_wall[i].Location as LocationCurve).Curve.Tessellate()[0];
                wall_curve_point = new XYZ(wall_curve_point.X, wall_curve_point.Y, 0);
                innerwall_points[i] = points[i] - (points[i] - wall_curve_point) * 2;
                for_check_innerwall_points[i] = points[i] - (points[i] - wall_curve_point) * 1.4; //計算支撐中間樁之擺放界線
            }
            innerwall_points[points.Count<XYZ>() - 1] = innerwall_points[0];
            for_check_innerwall_points[points.Count<XYZ>() - 1] = for_check_innerwall_points[0];

            //取得側牆profile
            IList<CurveLoop> bdtpCurve = new List<CurveLoop>();
            IList<Curve> side_wall_profileloops = new List<Curve>();
            CurveLoop bdtploop = new CurveLoop();
            for (int i = 0; i < innerwall_points.Count() - 1; i++)
            {
                Line line = Line.CreateBound(innerwall_points[i], innerwall_points[i + 1]);
                side_wall_profileloops.Add(line);
                bdtploop.Append(line);
            }

            //trans.commit();

            trans.Start("inner_buildingpad");
            // /*
            bdtpCurve.Add(bdtploop);
            //建立開挖深度
            ICollection<Level> level_family = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
            foreach (Level lev in level_family)
            {
                if (lev.Name == levlist[levlist.Count() - 1].Name)
                {
                    BuildingPad b = BuildingPad.Create(doc, bdtp.Id, lev.Id, bdtpCurve);
                    //b.get_Parameter(BuiltInParameter.)
                    //b.LookupParameter("厚度").SetValueString("1");
                    levdeep = lev;
                }
            }
            // */
            trans.Commit();

            trans.Start("side_wall");
            //建立側牆
            List<Wall> side_wall = new List<Wall>();
            foreach (WallType walltype in walltype_familyinstance)
            {
                if (walltype.Name == "側牆")
                {
                    for (int i = 0; i < csv.sidewall.Count; i++)
                    {
                        double floor_width = csv.back[i].Item3 * 1000 / 304.8; //set the width to a new value
                        double floor_deep = -csv.back[i].Item2 * 1000 / 304.8;
                        double floor_bottom = (floor_deep + floor_width / 2); //側牆底部高程點(考慮厚度做offset後-往上半個厚度)
                        double floor_top = (csv.back[i + 1].Item2 + csv.back[i + 1].Item3 / 2) * 1000 / 304.8; //側牆頂部高程點(考慮厚度做offset後-往下半個厚度)
                        for (int j = 0; j < points.Count<XYZ>() - 1; j++)
                        {
                            Curve side_c = side_wall_profileloops[j].CreateOffset((csv.sidewall[i].Item2 * 1000 / 2) / 304.8, new XYZ(0, 0, -1)); //此步驟為偏移擋土牆厚度1/2距離，作為建置參考線
                            Wall side_w = Wall.Create(doc, side_c, walltype.Id, re_levlist[i].Id, (floor_bottom * (-1) - floor_top), 0, false, false);
                            side_w.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(csv.sidewall[i].Item1 + "側牆");
                            side_wall.Add(side_w);
                        }
                    }
                }
            }
            trans.Commit();

            //取得側牆牆點座標
            CurveArray profileloops_array = new CurveArray();
            XYZ[] sidewall_points = new XYZ[points.Count<XYZ>()];
            for (int i = 0; i < (inner_wall.Count<Wall>()); i++)
            {
                //side
                XYZ side_wall_curve_point = (side_wall[i].Location as LocationCurve).Curve.Tessellate()[0];
                side_wall_curve_point = new XYZ(side_wall_curve_point.X, side_wall_curve_point.Y, 0);
                sidewall_points[i] = innerwall_points[i] - (innerwall_points[i] - side_wall_curve_point) * 2;
            }
            sidewall_points[points.Count<XYZ>() - 1] = sidewall_points[0];

            //建立底板範圍
            for (int i = 0; i < points.Count() - 1; i++)
            {
                Line line = Line.CreateBound(innerwall_points[i], innerwall_points[i + 1]);
                profileloops_array.Append(line);
            }

            trans.Start("建立底板");

            //建立底板類型及實作元件
            ICollection<FloorType> family = new FilteredElementCollector(doc).OfClass(typeof(FloorType)).Cast<FloorType>().ToList();
            FloorType floor_type = family.Where(x => x.Name == "通用 300mm").First();
            int base_control = 0;
            foreach (Tuple<string, double, double, double> base_tuple in csv.back)
            {
                FloorType newFamSym = null;
                try { newFamSym = floor_type.Duplicate(base_tuple.Item1) as FloorType; } catch { newFamSym = family.Where(x => x.Name == base_tuple.Item1).First(); }
                double floor_width = base_tuple.Item3 * 1000 / 304.8; //set the width to a new value
                double floor_deep = base_tuple.Item2 * -1 * 1000 / 304.8;
                double floor_offset = (floor_deep + floor_width / 2) * 304.8;
                newFamSym.GetCompoundStructure().GetLayers()[0].Width = floor_width;
                CompoundStructure ly = newFamSym.GetCompoundStructure();
                ly.SetLayerWidth(0, floor_width);
                newFamSym.SetCompoundStructure(ly);
                Floor floor = doc.Create.NewFloor(profileloops_array, newFamSym as FloorType, re_levlist[base_control], false);
                floor.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(base_tuple.Item1);
                floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).SetValueString("0");
                base_control += 1;
            }
            trans.Commit();

            //取得所有XY數值
            List<double> Xs = new List<double>();
            List<double> Ys = new List<double>();
            double[] slope = new double[inner_wall.Count<Wall>()];
            double[] bias = new double[inner_wall.Count<Wall>()];
            for (int i = 0; i < (innerwall_points.Count<XYZ>()); i++)
            {
                Xs.Add(innerwall_points[i].X);
                Ys.Add(innerwall_points[i].Y);
                if (i < slope.Count())
                {
                    if (innerwall_points[i + 1].X - innerwall_points[i].X == 0)
                    {
                        slope[i] = 20172018;
                        bias[i] = innerwall_points[i + 1].X;
                    }
                    else
                    {
                        slope[i] = (innerwall_points[i + 1].Y - innerwall_points[i].Y) / (innerwall_points[i + 1].X - innerwall_points[i].X);
                        if (slope[i] == 0 || Math.Abs(slope[i]) < 0.0000001)
                        {
                            bias[i] = innerwall_points[i + 1].Y;
                        }
                        else
                        {
                            bias[i] = innerwall_points[i + 1].Y - slope[i] * innerwall_points[i + 1].X;
                        }
                    }
                }
            }

            Transaction trans_2 = new Transaction(doc);
            trans_2.Start("交易開始");
            //開始建立中間樁
            double columns_dis = csv.centralCol[0].Item1 * 1000 / 304.8; //中間樁間距
            ICollection<FamilySymbol> columns_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
            foreach (FamilySymbol column_type in columns_familyinstance)
            {
                if (column_type.Name == "中間樁")
                {
                    for (int j = 0; j < (Math.Abs(Ys.Max() - Ys.Min()) / columns_dis - 1); j++)
                    {
                        for (int i = 0; i < (Math.Abs(Xs.Max() - Xs.Min()) / columns_dis - 1); j++)
                        {
                            XYZ column_location = new XYZ(Xs.Min() + (i + 1) * columns_dis, Ys.Min() + (j + 1) * columns_dis, 0);
                            if (IsInPolygon(column_location, innerwall_points) == true)
                            {
                                //TaskDialog.Show("asd", "asd");
                                FamilyInstance column_instance = doc.Create.NewFamilyInstance(column_location, column_type, levdeep, Autodesk.Revit.DB.Structure.StructuralType.Column);
                                column_instance.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).SetValueString("0"); //中間樁長度
                                column_instance.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("中間樁");
                            }
                        }
                    }
                }
            }

            //迴圈起始點
            for (int lev = 0; lev != csv.supLevel.Count(); lev++)
            {
                //建立圍囹
                //開始建立圍囹
                ICollection<FamilySymbol> beam_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
                foreach (FamilySymbol beam_type in beam_familyinstance)
                {
                    if (beam_type.Name == "H100x100")
                    {
                        double beam_H = double.Parse(beam_type.LookupParameter("H").AsValueString());
                        double beam_B = double.Parse(beam_type.LookupParameter("B").AsValueString());
                        for (int i = 0; i < innerwall_points.Count<XYZ>() - 1; i++)
                        {
                            Curve c = null;
                            if (i == points.Count<XYZ>())
                            {
                                try { int.Parse(csv.supLevel[lev].Item1.ToString()); c = Line.CreateBound(innerwall_points[i], innerwall_points[0]).CreateOffset((beam_H) / 304.8, new XYZ(0, 0, -1)); }
                                catch { c = Line.CreateBound(sidewall_points[i], sidewall_points[0]).CreateOffset((beam_H) / 304.8, new XYZ(0, 0, -1)); }
                            }
                            else
                            {
                                try { int.Parse(csv.supLevel[lev].Item1.ToString()); c = Line.CreateBound(innerwall_points[i], innerwall_points[i + 1]).CreateOffset((beam_H) / 304.8, new XYZ(0, 0, -1)); }
                                catch { c = Line.CreateBound(sidewall_points[i], sidewall_points[i + 1]).CreateOffset((beam_H) / 304.8, new XYZ(0, 0, -1)); }
                            }
                            FamilyInstance beam = doc.Create.NewFamilyInstance(c, beam_type, levlist[0], Autodesk.Revit.DB.Structure.StructuralType.Beam);
                            beam.LookupParameter("斷面旋轉").SetValueString("90");
                            StructuralFramingUtils.DisallowJoinAtEnd(beam, 0);
                            StructuralFramingUtils.DisallowJoinAtEnd(beam, 1);
                            beam.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set((csv.supLevel[lev].Item1).ToString() + "-圍囹");

                            //判斷圍囹之垂直深度，斜率零為負，反之為正

                            if ((i % 2 == 0))
                            {
                                beam.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((csv.supLevel[lev].Item2 * 1000 - beam_B / 2).ToString()); //2000為支撐階數深度，表1中
                            }
                            else
                            {
                                beam.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((csv.supLevel[lev].Item2 * 1000 + beam_B / 2).ToString());
                            }
                        }
                    }
                }

                //開始建立支撐
                //建立支撐

                XYZ frame_startpoint = null;
                XYZ frame_endpoint = null;

                ICollection<FamilySymbol> frame_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();

                foreach (FamilySymbol frame_type in frame_familyinstance)
                {
                    if (frame_type.Name == "H100x100")
                    {
                        double frame_H = double.Parse(frame_type.LookupParameter("H").AsValueString());

                        //X向支撐
                        for (int j = 0; j < Math.Abs(Ys.Max() - Ys.Min()) / columns_dis - 1; j++)
                        {
                            for (int i = 0; i < Math.Abs(Xs.Max() - Xs.Min()) / columns_dis - 1; i++)
                            {
                                XYZ frame_location = new XYZ(Xs.Min() + (i + 1) * columns_dis, Ys.Min() + (j + 1) * columns_dis, 0);
                                frame_startpoint = intersection(frame_location, for_check_innerwall_points, slope, bias, true)[0];
                                if (IsInPolygon(frame_location, innerwall_points) == true)
                                {
                                    try
                                    {
                                        frame_endpoint = intersection(frame_location, for_check_innerwall_points, slope, bias, true)[1];
                                        Line line = Line.CreateBound(frame_startpoint, frame_endpoint);
                                        FamilyInstance frame_instance = doc.Create.NewFamilyInstance(line, frame_type, levdeep, Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                        frame_instance.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set((csv.supLevel[lev].Item1).ToString() + "-支撐");
                                        //處理偏移與延伸問題
                                        frame_instance.get_Parameter(BuiltInParameter.Z_OFFSET_VALUE).SetValueString((csv.supLevel[lev].Item2 * -1000).ToString());
                                        frame_instance.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((-frame_H).ToString());
                                        try
                                        {
                                            int.Parse(csv.supLevel[lev].Item1.ToString());
                                            frame_instance.get_Parameter(BuiltInParameter.START_EXTENSION).SetValueString((-frame_H).ToString());
                                            frame_instance.get_Parameter(BuiltInParameter.END_EXTENSION).SetValueString((-frame_H).ToString());
                                        }
                                        catch
                                        {
                                            frame_instance.get_Parameter(BuiltInParameter.START_EXTENSION).SetValueString((-frame_H - csv.sidewall[0].Item2 * 1000).ToString());
                                            frame_instance.get_Parameter(BuiltInParameter.END_EXTENSION).SetValueString((-frame_H - csv.sidewall[0].Item2 * 1000).ToString());
                                        }
                                        //取消接合
                                        StructuralFramingUtils.DisallowJoinAtEnd(frame_instance, 0);
                                        StructuralFramingUtils.DisallowJoinAtEnd(frame_instance, 1);

                                        //若為雙向支撐，則鏡射支撐
                                        if (csv.supLevel[lev].Item3 == 2)
                                        {
                                            ElementTransformUtils.MirrorElement(doc, frame_instance.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, frame_startpoint));
                                        }
                                        break;
                                    }
                                    catch { }
                                }
                            }
                        }

                        //Y向支撐
                        for (int i = 0; i < (Math.Abs(Xs.Max() - Xs.Min()) / columns_dis - 1); i++)
                        {
                            for (int j = 0; j < (Math.Abs(Ys.Max() - Ys.Min()) / columns_dis - 1); j++)
                            {
                                XYZ frame_location = new XYZ(Xs.Min() + (i + 1) * columns_dis, Ys.Min() + (1 + j) * columns_dis, 0);
                                frame_startpoint = intersection(frame_location, for_check_innerwall_points, slope, bias, false)[0];
                                if (IsInPolygon(frame_location, innerwall_points) == true)
                                {
                                    try
                                    {
                                        frame_endpoint = intersection(frame_location, for_check_innerwall_points, slope, bias, false)[1];
                                        Line line = Line.CreateBound(frame_startpoint, frame_endpoint);
                                        FamilyInstance frame_instance = doc.Create.NewFamilyInstance(line, frame_type, levdeep, Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                        frame_instance.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set((csv.supLevel[lev].Item1).ToString() + "-支撐");

                                        //處理偏移與延伸問題

                                        frame_instance.get_Parameter(BuiltInParameter.Z_OFFSET_VALUE).SetValueString((csv.supLevel[lev].Item2 * -1000 + frame_H).ToString());//2000為支撐階數深度，表1中
                                        frame_instance.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((-frame_H).ToString());
                                        try
                                        {
                                            int.Parse(csv.supLevel[lev].Item1.ToString());
                                            frame_instance.get_Parameter(BuiltInParameter.START_EXTENSION).SetValueString((-frame_H).ToString());
                                            frame_instance.get_Parameter(BuiltInParameter.END_EXTENSION).SetValueString((-frame_H).ToString());
                                        }
                                        catch
                                        {
                                            frame_instance.get_Parameter(BuiltInParameter.START_EXTENSION).SetValueString((-frame_H - csv.sidewall[0].Item2 * 1000).ToString());
                                            frame_instance.get_Parameter(BuiltInParameter.END_EXTENSION).SetValueString((-frame_H - csv.sidewall[0].Item2 * 1000).ToString());
                                        }
                                        //取消接合
                                        StructuralFramingUtils.DisallowJoinAtEnd(frame_instance, 0);
                                        StructuralFramingUtils.DisallowJoinAtEnd(frame_instance, 1);

                                        //若為雙向支撐，則鏡射支撐
                                        if (csv.supLevel[lev].Item3 == 2)
                                        {
                                            ElementTransformUtils.MirrorElement(doc, frame_instance.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, frame_startpoint));
                                        }
                                        break;
                                    }
                                    catch { }
                                }
                            }
                        }
                    }
                }

                //建立斜撐
                ICollection<FamilySymbol> slopframe_symbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
                foreach (FamilySymbol slopframe_type in slopframe_symbol)
                {
                    if (slopframe_type.Name == "斜撐")
                    {
                        //X向斜撐
                        for (int j = 0; j < (Math.Abs(Ys.Max() - Ys.Min()) / columns_dis - 1); j++)
                        {
                            for (int i = 0; i < (Math.Abs(Xs.Max() - Xs.Min()) / columns_dis - 1); i++)
                            {
                                XYZ frame_location = new XYZ(Xs.Min() + (i + 1) * columns_dis, Ys.Min() + (j + 1) * columns_dis, 0);
                                frame_startpoint = intersection(frame_location, for_check_innerwall_points, slope, bias, true)[0];
                                if (IsInPolygon(frame_location, innerwall_points) == true)
                                {
                                    frame_endpoint = intersection(frame_location, for_check_innerwall_points, slope, bias, true)[1];
                                    FamilyInstance slopframe_1 = doc.Create.NewFamilyInstance(frame_startpoint, slopframe_type, all_level_list[lev], Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                    FamilyInstance slopframe_2 = doc.Create.NewFamilyInstance(frame_endpoint, slopframe_type, all_level_list[lev], Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                    slopframe_1.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set((csv.supLevel[lev].Item1).ToString() + "-斜撐");
                                    slopframe_2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set((csv.supLevel[lev].Item1).ToString() + "-斜撐");
                                    slopframe_1.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(Math.Abs(all_level_list[lev].Elevation - csv.supLevel[lev].Item2 * -1000 / 304.8 + slopframe_1.LookupParameter("支撐厚度").AsDouble() / 2));

                                    slopframe_2.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(Math.Abs(all_level_list[lev].Elevation - csv.supLevel[lev].Item2 * -1000 / 304.8 + slopframe_1.LookupParameter("支撐厚度").AsDouble() / 2));
                                    int o;
                                    if (int.TryParse(csv.supLevel[lev].Item1, out o) == false)//305為圍囹寬度，應該要查表而非指定。
                                    {
                                        slopframe_1.LookupParameter("圍囹寬度").SetValueString((305 + csv.sidewall[0].Item2 * 1000).ToString());
                                        slopframe_2.LookupParameter("圍囹寬度").SetValueString((305 + csv.sidewall[0].Item2 * 1000).ToString());
                                    }
                                    else
                                    {
                                        slopframe_1.LookupParameter("圍囹寬度").SetValueString((305).ToString());
                                        slopframe_2.LookupParameter("圍囹寬度").SetValueString((305).ToString());
                                    }
                                    //旋轉斜撐元件
                                    Line rotate_line_s = Line.CreateBound(frame_startpoint, frame_startpoint + new XYZ(0, 0, 1));
                                    Line rotate_line_e = Line.CreateBound(frame_endpoint, frame_endpoint + new XYZ(0, 0, 1));
                                    slopframe_1.Location.Rotate(rotate_line_s, 1.5 * Math.PI);
                                    slopframe_2.Location.Rotate(rotate_line_e, 0.5 * Math.PI);

                                    //鏡射斜撐元件
                                    if (csv.supLevel[lev].Item3 == 2)//雙排
                                    {
                                        ElementTransformUtils.MirrorElement(doc, slopframe_1.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, frame_startpoint));
                                        ElementTransformUtils.MirrorElement(doc, slopframe_2.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, frame_endpoint));
                                    }
                                    else//單排
                                    {
                                        ElementTransformUtils.MirrorElement(doc, slopframe_1.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, frame_startpoint.Add(new XYZ(0, -(slopframe_1.LookupParameter("支撐厚度").AsDouble() / 2 + slopframe_1.LookupParameter("中間樁直徑").AsDouble() / 2), 0))));
                                        ElementTransformUtils.MirrorElement(doc, slopframe_2.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, frame_endpoint));
                                        slopframe_2.Location.Move((new XYZ(0, -(slopframe_1.LookupParameter("支撐厚度").AsDouble() + slopframe_1.LookupParameter("中間樁直徑").AsDouble()), 0)));
                                    }

                                    break;
                                }
                            }
                        }

                        //Y向斜撐
                        for (int i = 0; i < (Math.Abs(Xs.Max() - Xs.Min()) / columns_dis - 1); i++)
                        {
                            for (int j = 0; j < (Math.Abs(Ys.Max() - Ys.Min()) / columns_dis - 1); j++)
                            {
                                XYZ frame_location = new XYZ(Xs.Min() + (i + 1) * columns_dis, Ys.Min() + (1 + j) * columns_dis, 0);
                                frame_startpoint = intersection(frame_location, for_check_innerwall_points, slope, bias, false)[0];
                                if (IsInPolygon(frame_location, innerwall_points) == true)
                                {
                                    frame_endpoint = intersection(frame_location, for_check_innerwall_points, slope, bias, false)[1];
                                    FamilyInstance slopframe_1 = doc.Create.NewFamilyInstance(frame_startpoint, slopframe_type, all_level_list[lev], Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                    FamilyInstance slopframe_2 = doc.Create.NewFamilyInstance(frame_endpoint, slopframe_type, all_level_list[lev], Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                    slopframe_1.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set((csv.supLevel[lev].Item1).ToString() + "-斜撐");
                                    slopframe_2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set((csv.supLevel[lev].Item1).ToString() + "-斜撐");
                                    slopframe_1.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(Math.Abs(all_level_list[lev].Elevation - csv.supLevel[lev].Item2 * -1000 / 304.8 - slopframe_1.LookupParameter("支撐厚度").AsDouble() / 2));

                                    slopframe_2.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(Math.Abs(all_level_list[lev].Elevation - csv.supLevel[lev].Item2 * -1000 / 304.8 - slopframe_2.LookupParameter("支撐厚度").AsDouble() / 2));
                                    int o;
                                    if (int.TryParse(csv.supLevel[lev].Item1, out o) == false)//305為圍囹寬度，應該要查表而非指定。
                                    {
                                        slopframe_1.LookupParameter("圍囹寬度").SetValueString((305 + csv.sidewall[0].Item2 * 1000).ToString());
                                        slopframe_2.LookupParameter("圍囹寬度").SetValueString((305 + csv.sidewall[0].Item2 * 1000).ToString());
                                    }
                                    else
                                    {
                                        slopframe_1.LookupParameter("圍囹寬度").SetValueString((305).ToString());
                                        slopframe_2.LookupParameter("圍囹寬度").SetValueString((305).ToString());
                                    }
                                    //旋轉斜撐元件
                                    Line rotate_line = Line.CreateBound(frame_endpoint, frame_endpoint + new XYZ(0, 0, 1));
                                    slopframe_2.Location.Rotate(rotate_line, Math.PI);

                                    //鏡射斜撐元件
                                    if (csv.supLevel[lev].Item3 == 2)//雙排
                                    {
                                        ElementTransformUtils.MirrorElement(doc, slopframe_1.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, frame_startpoint));
                                        ElementTransformUtils.MirrorElement(doc, slopframe_2.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, frame_endpoint));
                                    }
                                    else//單排
                                    {
                                        ElementTransformUtils.MirrorElement(doc, slopframe_1.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, frame_startpoint.Add(new XYZ(slopframe_1.LookupParameter("支撐厚度").AsDouble() / 2 + slopframe_1.LookupParameter("中間樁直徑").AsDouble() / 2, 0, 0))));
                                        ElementTransformUtils.MirrorElement(doc, slopframe_2.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, frame_endpoint));
                                        slopframe_2.Location.Move((new XYZ((slopframe_1.LookupParameter("支撐厚度").AsDouble() + slopframe_1.LookupParameter("中間樁直徑").AsDouble()), 0, 0)));
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            trans_2.Commit();
            return Result.Succeeded;
        }

        public static bool calculator(XYZ checkPoint, List<XYZ> ranoutpoints) //判斷點與點是否為相同點，因為xyz類型無法做equals
        {
            bool a = false;
            foreach (XYZ ranoutpoint in ranoutpoints)
            {
                if (ranoutpoint.X - checkPoint.X < 1 && ranoutpoint.X - checkPoint.X > -1)
                {
                    if (ranoutpoint.Y - checkPoint.Y < 1 && ranoutpoint.Y - checkPoint.Y > -1)
                    {
                        if (ranoutpoint.Z - checkPoint.Z < 1 && ranoutpoint.Z - checkPoint.Z > -1)
                        {
                            a = true;
                        }
                    }
                }
            }
            return a;
        }

        public static bool IsInPolygon(XYZ checkPoint, XYZ[] polygonPoints) //判斷點是否位於開挖範圍內
        {
            bool inside = false;
            int pointCount = polygonPoints.Count<XYZ>();
            XYZ p1, p2;
            for (int i = 0, j = pointCount - 1; i < pointCount; j = i, i++)
            {
                p1 = polygonPoints[i];
                p2 = polygonPoints[j];
                if (checkPoint.Y < p2.Y)
                {
                    if (p1.Y <= checkPoint.Y)
                    {
                        if ((checkPoint.Y - p1.Y) * (p2.X - p1.X) > (checkPoint.X - p1.X) * (p2.Y - p1.Y))
                        {
                            inside = (!inside);
                        }
                    }
                }
                else if (checkPoint.Y < p1.Y)
                {
                    if ((checkPoint.Y - p1.Y) * (p2.X - p1.X) < (checkPoint.X - p1.X) * (p2.Y - p1.Y))
                    {
                        inside = (!inside);
                    }
                }
            }
            return inside;
        }

        public static XYZ[] intersection(XYZ checkPoint, XYZ[] polygonPoints, double[] slope, double[] bias, bool x_or_y) //計算支撐與連續壁交點，true為計算X向，false為計算Y向
        {
            List<XYZ> intersection_points = new List<XYZ>();
            List<double> x = new List<double>();
            List<double> y = new List<double>();
            if (x_or_y == true)
            {
                for (int i = 0; i < slope.Count<double>(); i++)
                {
                    if (slope[i] != 0 || Math.Abs(slope[i]) > 0.00001)
                    {
                        if (slope[i] == 20172018)
                        {
                            intersection_points.Add(new XYZ(bias[i], checkPoint.Y, 0));
                        }
                        else
                        {
                            intersection_points.Add(new XYZ((checkPoint.Y - bias[i]) / slope[i], checkPoint.Y, 0));
                        }
                    }
                }

                foreach (XYZ point in intersection_points)
                {
                    if (IsInPolygon(point, polygonPoints))
                    {
                        x.Add(point.X);
                        y.Add(point.Y);
                    }
                }
            }
            else //若為false則計算Y向
            {
                for (int i = 0; i < slope.Count<double>(); i++)
                {
                    if (slope[i] != 20172018)
                    {
                        if (slope[i] == 0 || Math.Abs(slope[i]) < 0.00001)
                        {
                            intersection_points.Add(new XYZ(checkPoint.X, bias[i], 0));
                        }
                        else
                        {
                            intersection_points.Add(new XYZ(checkPoint.X, slope[i] * checkPoint.X + bias[i], 0));
                        }
                    }
                }
                foreach (XYZ point in intersection_points)
                {
                    if (IsInPolygon(point, polygonPoints))
                    {
                        x.Add(point.X);
                        y.Add(point.Y);
                    }
                }
            }
            XYZ[] intersection = new XYZ[2];
            try
            {
                intersection[0] = new XYZ(x.Min(), y.Min(), 0);
                intersection[1] = new XYZ(x.Max(), y.Max(), 0);
                return intersection;
            }
            catch { return intersection; }
        }
    }
}