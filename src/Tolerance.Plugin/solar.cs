using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using AutoCAD;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using Autodesk.Civil;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using PvGrade;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;
using Clipper2Lib;
using System.Diagnostics;
using System.Text;
using ClosedXML;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace solar
{

    public class solarFunctions
    {


        public static void BlockToCogo()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database docDb = doc.Database;
            var docScale = docDb.Cannoscale.DrawingUnits;
            using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
            {

                CogoPointCollection cogoPoints = CivilApplication.ActiveDocument.CogoPoints;
                /*var boundaryLayer = new PromptStringOptions("\nName of Boundary Layer");

                var boundaryResult = doc.Editor.GetString(boundaryLayer);
                if (boundaryResult.Status != PromptStatus.OK)
                {
                    return;
                }*/



                TypedValue[] tvs = new TypedValue[] {

                    new TypedValue(Convert.ToInt32(DxfCode.Operator), "<or"),
                    new TypedValue(Convert.ToInt32(DxfCode.Start), "POLYLINE"),
                    new TypedValue(Convert.ToInt32(DxfCode.Start), "LWPOLYLINE"),
                    new TypedValue(Convert.ToInt32(DxfCode.Start), "POLYLINE2D"),
                    new TypedValue(Convert.ToInt32(DxfCode.Start), "POLYLINE3d"),
                    new TypedValue(Convert.ToInt32(DxfCode.Operator), "or>"),

                };

                // Assign the filter criteria to a SelectionFilter object
                SelectionFilter acSelFtr = new SelectionFilter(tvs);

                // Request for objects to be selected in the drawing area
                PromptSelectionResult acSSPrompt;
                acSSPrompt = doc.Editor.SelectAll(acSelFtr);

                // If the prompt status is OK, objects were selected
                if (acSSPrompt.Status == PromptStatus.OK)
                {



                    TypedValue[] bf = new TypedValue[] {
                    new TypedValue(Convert.ToInt32(DxfCode.Start), "INSERT"),
                    };

                    doc.Editor.Command("Zoom", "all");
                    SelectionFilter blockFilter = new SelectionFilter(bf);
                    foreach (SelectedObject so in acSSPrompt.Value)
                    {

                        var ent = (Polyline)transaction.GetObject(so.ObjectId, OpenMode.ForRead);
                        PromptSelectionResult blockPrompt;

                        blockPrompt = doc.Editor.SelectWindow(ent.Bounds.Value.MinPoint, ent.Bounds.Value.MaxPoint, blockFilter);
                        foreach (SelectedObject b in blockPrompt.Value)
                        {
                            var bl = (BlockReference)transaction.GetObject(b.ObjectId, OpenMode.ForRead);
                            ObjectId pointId = cogoPoints.Add(bl.Position, ent.Layer, false);
                            CogoPoint cogoPoint = pointId.GetObject(OpenMode.ForWrite) as CogoPoint;
                            cogoPoint.RawDescription = ent.Layer;
                            cogoPoint.Layer = ent.Layer;
                        }



                    }
                }
                else
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Number of objects selected: 0; contact Taylor if this isnt intended");
                }








                doc.TransactionManager.EnableGraphicsFlush(true);
                doc.TransactionManager.FlushGraphics();
                doc.Editor.Regen();
                transaction.Commit();
            }
        }

        public struct MyVector
        {
            private readonly double _x, _y;


            // Constructor
            public MyVector(double x, double y)
            {
                _x = x;
                _y = y;
            }


            // Distance from this point to another point, squared
            private double DistanceSquared(MyVector otherPoint)
            {
                double dx = otherPoint._x - this._x;
                double dy = otherPoint._y - this._y;
                return dx * dx + dy * dy;
            }


            // Find the distance from this point to a line segment (which is not the same as from this 
            //  point to anywhere on an infinite line). Also returns the closest point.
            public double DistanceToLineSegment(MyVector lineSegmentPoint1, MyVector lineSegmentPoint2,
                                                out MyVector closestPoint)
            {
                return Math.Sqrt(DistanceToLineSegmentSquared(lineSegmentPoint1, lineSegmentPoint2,
                                 out closestPoint));
            }


            // Same as above, but avoid using Sqrt(), saves a new nanoseconds in cases where you only want 
            //  to compare several distances to find the smallest or largest, but don't need the distance
            public double DistanceToLineSegmentSquared(MyVector lineSegmentPoint1,
                                                    MyVector lineSegmentPoint2, out MyVector closestPoint)
            {
                // Compute length of line segment (squared) and handle special case of coincident points
                double segmentLengthSquared = lineSegmentPoint1.DistanceSquared(lineSegmentPoint2);
                if (segmentLengthSquared < 1E-7f)  // Arbitrary "close enough for government work" value
                {
                    closestPoint = lineSegmentPoint1;
                    return this.DistanceSquared(closestPoint);
                }

                // Use the magic formula to compute the "projection" of this point on the infinite line
                MyVector lineSegment = lineSegmentPoint2 - lineSegmentPoint1;
                double t = (this - lineSegmentPoint1).DotProduct(lineSegment) / segmentLengthSquared;

                // Handle the two cases where the projection is not on the line segment, and the case where 
                //  the projection is on the segment
                if (t <= 0)
                    closestPoint = lineSegmentPoint1;
                else if (t >= 1)
                    closestPoint = lineSegmentPoint2;
                else
                    closestPoint = lineSegmentPoint1 + (lineSegment * t);
                return this.DistanceSquared(closestPoint);
            }


            public double DotProduct(MyVector otherVector)
            {
                return this._x * otherVector._x + this._y * otherVector._y;
            }

            public static MyVector operator +(MyVector leftVector, MyVector rightVector)
            {
                return new MyVector(leftVector._x + rightVector._x, leftVector._y + rightVector._y);
            }

            public static MyVector operator -(MyVector leftVector, MyVector rightVector)
            {
                return new MyVector(leftVector._x - rightVector._x, leftVector._y - rightVector._y);
            }

            public static MyVector operator *(MyVector aVector, double aScalar)
            {
                return new MyVector(aVector._x * aScalar, aVector._y * aScalar);
            }

            // Added using ReSharper due to CodeAnalysis nagging

            public bool Equals(MyVector other)
            {
                return _x.Equals(other._x) && _y.Equals(other._y);
            }

            public override bool Equals(object obj)
            {
                if (ReferenceEquals(null, obj)) return false;
                return obj is MyVector && Equals((MyVector)obj);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    return (_x.GetHashCode() * 397) ^ _y.GetHashCode();
                }
            }

            public static bool operator ==(MyVector left, MyVector right)
            {
                return left.Equals(right);
            }

            public static bool operator !=(MyVector left, MyVector right)
            {
                return !left.Equals(right);
            }
        }

        public static class JustTesting
        {
            public static void Main()
            {
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();

                for (int i = 0; i < 10000000; i++)
                {
                    TestIt(1, 0, 0, 0, 1, 1, 0.70710678118654757);
                    TestIt(5, 4, 0, 0, 20, 10, 1.3416407864998738);
                    TestIt(30, 15, 0, 0, 20, 10, 11.180339887498949);
                    TestIt(-30, 15, 0, 0, 20, 10, 33.541019662496844);
                    TestIt(5, 1, 0, 0, 10, 0, 1.0);
                    TestIt(1, 5, 0, 0, 0, 10, 1.0);
                }

                stopwatch.Stop();
                TimeSpan timeSpan = stopwatch.Elapsed;
            }


            private static void TestIt(float aPointX, float aPointY,
                                       float lineSegmentPoint1X, float lineSegmentPoint1Y,
                                       float lineSegmentPoint2X, float lineSegmentPoint2Y,
                                       double expectedAnswer)
            {
                // Katz
                double d1 = DistanceFromPointToLineSegment(new MyVector(aPointX, aPointY),
                                                     new MyVector(lineSegmentPoint1X, lineSegmentPoint1Y),
                                                     new MyVector(lineSegmentPoint2X, lineSegmentPoint2Y));
                Debug.Assert(d1 == expectedAnswer);

            }

            private static double DistanceFromPointToLineSegment(MyVector aPoint,
                                                   MyVector lineSegmentPoint1, MyVector lineSegmentPoint2)
            {
                MyVector closestPoint;  // Not used
                return aPoint.DistanceToLineSegment(lineSegmentPoint1, lineSegmentPoint2,
                                                    out closestPoint);
            }

            private static double DistanceFromPointToLineSegmentSquared(MyVector aPoint,
                                                   MyVector lineSegmentPoint1, MyVector lineSegmentPoint2)
            {
                MyVector closestPoint;  // Not used
                return aPoint.DistanceToLineSegmentSquared(lineSegmentPoint1, lineSegmentPoint2,
                                                           out closestPoint);
            }
        }
        struct Pile
        {
            public double Y;
            public ObjectId pointId;

            public Pile(double y, ObjectId pointId)
            {
                this.Y = y;
                this.pointId = pointId;
            }

        }

        struct NSPile
        {
            public double X;
            public double Y;
            public int pb;
            public string layer;
            public int pointNum;
            public ObjectId pointId;

            public NSPile(double x, double y, int pb, string layer, ObjectId pointId)
            {
                this.X = x;
                this.Y = y;
                this.pb = pb;
                this.layer = layer;
                this.pointId = pointId;
            }

        }

        class PileComparer : IComparer<Pile>
        {
            public int Compare(Pile a, Pile b)
            {
                if (a.Y < b.Y) return 0;
                return -1;
            }
        }


        private static int determinePBrow(double val, double[] breakpoints, int ind)
        {
            if (ind == breakpoints.Length) return ind;
            if (val > breakpoints[ind]) return ind;
            return determinePBrow(val, breakpoints, ind + 1);
        }

        private static void writeSummaryCSV(Dictionary<string, int> d, string pb, string filePath)
        {

            using (var workbook = new XLWorkbook("F:\\CAD Support\\Lisp Routines\\solarRoutines\\pileNumbering\\script\\Block1.xlsx"))
            {
                var ws = workbook.Worksheet(1);

                //ws.Column("A").Width = 14;
                //ws.Column("B").Width = 21;
                //ws.Row(1).Height = 30;

                int Total = 0;
                int rowInd = 2;

                while (rowInd < 100)
                {
                    var sl = ws.Cell("A" + rowInd.ToString()).Value;
                    if (sl.ToString() == "TOTAL")
                    {
                        ws.Cell("B" + rowInd.ToString()).Value = Total;
                        break;
                    }
                    if (d.Keys.Contains(sl))
                    {
                        ws.Cell("B" + rowInd.ToString()).Value = d[sl.ToString()];
                        Total += d[sl.ToString()];
                    }
                    rowInd++;
                }
                //foreach (KeyValuePair<string, int> row in d.OrderByDescending(x => x.Key))
                //{
                //    ws.Cell("A"+rowInd.ToString()).Value = row.Key;
                //    ws.Cell("B" + rowInd.ToString()).Value = row.Value;
                //    Total += row.Value;
                //    rowInd++;
                //}

                //ws.Cell("A1").Value = "Size and length";
                //ws.Cell("B1").Value = "Quantity";
                //ws.Cell("A" + rowInd.ToString()).Value = "TOTAL";
                //ws.Cell("B" + rowInd.ToString()).Value = Total;
                //var table = ws.Range("A1:B" + rowInd.ToString());
                //table.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                //table.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                //ws.Range("A1:B1").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                //ws.Range("A" + rowInd.ToString()+ ":B" + rowInd.ToString()).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                workbook.SaveAs($"{filePath}/PB{pb}_Summary.xlsx");

            }
        }

        private static List<double> getRowBreakpoints(Document doc)
        {
            List<double> breakpointList = new List<double>();
            PromptPointOptions ppo = new PromptPointOptions("\nclick on row breakpoints, escape to continue");
            ppo.SetMessageAndKeywords($"\n {breakpointList.Count} selected; click on row breakpoints or [Done/Restart]", "Done Restart");

            var pointResult = doc.Editor.GetPoint(ppo);
            while (pointResult.Status == PromptStatus.OK)
            {
                breakpointList.Add(pointResult.Value.Y);
                ppo.SetMessageAndKeywords($"\n {breakpointList.Count} selected; click on row breakpoints or [Done/Restart]", "Done Restart");
                pointResult = doc.Editor.GetPoint(ppo);

            }
            if (pointResult.StringResult == "Done") return breakpointList;
            return getRowBreakpoints(doc);

        }
        private static (Dictionary<string, string[]>, List<string>, Dictionary<string, int>, List<string>, string) processRowCSV()
        {
            Dictionary<string, string[]> rowTypes = new Dictionary<string, string[]>();
            List<string> dimCounts = new List<string>();
            Dictionary<string, int> rowNames = new Dictionary<string, int>();
            List<string> globalLayers = new List<string>();
            PromptOpenFileOptions pofo = new PromptOpenFileOptions("\nEnter Row Type CSV File");
            pofo.PreferCommandLine = false;
            pofo.DialogName = "Select File";
            pofo.DialogCaption = "Select File";

            PromptFileNameResult pfnr = Application.DocumentManager.MdiActiveDocument.Editor.GetFileNameForOpen(pofo);

            string FileName = pfnr.StringResult;
            string[] filepathsplit = pfnr.StringResult.Split('\\');
            Array.Resize(ref filepathsplit, filepathsplit.Length - 1);
            string filePath = String.Join("\\", filepathsplit) + "\\csv\\";
            System.IO.Directory.CreateDirectory(filePath);
            FileStream Fstream = new FileStream(FileName, FileMode.Open);
            StreamReader read = new StreamReader(Fstream);
            string st = read.ReadLine();
            st = read.ReadLine();

            while (st != null)
            {
                var line = st.Split(',');
                rowTypes[line[0]] = new string[] { line[2], line[3] };
                if (line[9] != "")
                {
                    rowNames[line[9]] = Int32.Parse(line[10]);
                    globalLayers.Add(line[0]);
                }
                dimCounts.Add(line[3]);
                st = read.ReadLine();

            }
            Fstream.Dispose();

            return (rowTypes, dimCounts, rowNames, globalLayers, filePath);
        }

        static TypedValue[] bf = new TypedValue[] { new TypedValue(Convert.ToInt32(DxfCode.Start), "INSERT"), };
        //static TypedValue[] tvs = new TypedValue[] {
        //            new TypedValue(Convert.ToInt32(DxfCode.Operator), "<or"),
        //            new TypedValue(Convert.ToInt32(DxfCode.Start), "POLYLINE"),
        //            new TypedValue(Convert.ToInt32(DxfCode.Start), "LWPOLYLINE"),
        //            new TypedValue(Convert.ToInt32(DxfCode.Start), "POLYLINE2D"),
        //            new TypedValue(Convert.ToInt32(DxfCode.Start), "POLYLINE3d"),
        //            new TypedValue(Convert.ToInt32(DxfCode.LayerName), "CAB1"),
        //            new TypedValue(Convert.ToInt32(DxfCode.LayerName), "CAB2"),
        //            new TypedValue(Convert.ToInt32(DxfCode.LayerName), "CAB3"),
        //            new TypedValue(Convert.ToInt32(DxfCode.LayerName), "LBD1"),
        //            new TypedValue(Convert.ToInt32(DxfCode.LayerName), "LBD2"),
        //            new TypedValue(Convert.ToInt32(DxfCode.LayerName), "INV1"),
        //            new TypedValue(Convert.ToInt32(DxfCode.LayerName), "INV2"),
        //            new TypedValue(Convert.ToInt32(DxfCode.Operator), "or>"),
        //        };

        //public static SelectionFilter polylineFilter = new SelectionFilter(tvs);
        public static SelectionFilter blockFilter = new SelectionFilter(bf);

        private static SelectionFilter createSelectionFilter(List<string> layers, bool isGlobal)
        {
            List<TypedValue> gtvs;

            if (isGlobal)
            {
                gtvs = new List<TypedValue> {
                    new TypedValue(Convert.ToInt32(DxfCode.Start), "AECC_COGO_POINT"),
                    new TypedValue(Convert.ToInt32(DxfCode.Operator), "<or")
                 };
            }
            else
            {
                gtvs = new List<TypedValue> {
                    new TypedValue(Convert.ToInt32(DxfCode.Operator), "<or"),
                    new TypedValue(Convert.ToInt32(DxfCode.Start), "POLYLINE"),
                    new TypedValue(Convert.ToInt32(DxfCode.Start), "LWPOLYLINE"),
                    new TypedValue(Convert.ToInt32(DxfCode.Start), "POLYLINE2D"),
                    new TypedValue(Convert.ToInt32(DxfCode.Start), "POLYLINE3d"),
                };
            }
            for (int i = 0; i < layers.Count; i++)
            {
                gtvs.Add(new TypedValue(Convert.ToInt32(DxfCode.LayerName), layers[i]));
            };

            gtvs.Add(new TypedValue(Convert.ToInt32(DxfCode.Operator), "or>"));
            return new SelectionFilter(gtvs.ToArray());
        }
        public static void processGlobalPiles()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var (rowTypes, _, rowNames, globalLayers, filePath) = processRowCSV();

            using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
            {

                doc.Editor.Command("Zoom", "extents");


                SelectionFilter globalFilter = createSelectionFilter(globalLayers, true);
                List<NSPile> globalPiles = new List<NSPile>();
                Dictionary<string, List<NSPile>> globalByPB = new Dictionary<string, List<NSPile>>();

                var globalCogoSelection = doc.Editor.SelectAll(globalFilter);

                for (int bb = 0; bb < globalCogoSelection.Value.Count; bb++)
                {
                    CogoPoint obj = (CogoPoint)transaction.GetObject(globalCogoSelection.Value[bb].ObjectId, OpenMode.ForRead);
                    globalPiles.Add(new NSPile(Math.Round(obj.Location.X, 4), Math.Round(obj.Location.Y, 4), Int32.Parse(obj.RawDescription), obj.Layer, globalCogoSelection.Value[bb].ObjectId));

                }

                foreach (KeyValuePair<string, int> row in rowNames)
                {
                    var pileGroup = globalPiles.Where(o => o.layer.Contains(row.Key));
                    var sorted = pileGroup.OrderBy(p => p.X).ThenBy(p => p.Y);
                    int offset = 1;
                    foreach (NSPile p in sorted)
                    {
                        //calc point num
                        CogoPoint cogoPoint = p.pointId.GetObject(OpenMode.ForWrite) as CogoPoint;
                        cogoPoint.PointNumber = (uint)(row.Value + offset);
                        NSPile p2 = p;
                        p2.pointNum = row.Value + offset;
                        if (globalByPB.ContainsKey(p2.pb.ToString()))
                        {
                            globalByPB[p2.pb.ToString()].Add(p2);
                        }
                        else
                        {
                            globalByPB[p2.pb.ToString()] = new List<NSPile> { p2 };
                        }
                        offset += 1;
                    }

                }
                //get elevation value at surface point
                { // Prompt for the point
                  //https://help.autodesk.com/view/CIV3D/2024/ENU/?guid=d280ff4f-4c4b-f910-7a27-0e787a43448f
                  //PromptPointResult pointResult = ed.GetPoint("\nEnter a point to get elevation: ");
                  //if (pointResult.Status != PromptStatus.OK) return;
                  //Point3d point = pointResult.Value;

                    //// Access the surface
                    //using (Transaction tr = doc.TransactionManager.StartTransaction())
                    //{
                    //    CivilDocument doca = CivilApplication.ActiveDocument;
                    //    Editor ed = doc.Editor;
                    //    ObjectId surfaceId = doca.GetSurfaceIds()[0]; // Assuming the first surface in the drawing
                    //    TinSurface surface = tr.GetObject(surfaceId, OpenMode.ForRead) as TinSurface;

                    //    // Get elevation at the specified XY location
                    //    double elevation = surface.FindElevationAtXY(point.X, point.Y);
                    //    ed.WriteMessage($"\nThe elevation at X: {point.X}, Y: {point.Y} is: {elevation}");
                }

                foreach (KeyValuePair<string, List<NSPile>> row in globalByPB)
                {
                    using (FileStream fs = new FileStream($"{filePath}/PB{row.Key}.csv", FileMode.Append, FileAccess.Write))
                    {
                        using (StreamWriter sw = new StreamWriter(fs))
                        {
                            foreach (NSPile p in row.Value) { sw.WriteLine($"{p.pointNum},{p.X.ToString()},{p.Y.ToString()},0,{rowTypes[p.layer][0]}"); }

                        }
                    }
                }
                doc.TransactionManager.EnableGraphicsFlush(true);
                doc.TransactionManager.FlushGraphics();
                doc.Editor.Regen();
                transaction.Commit();
            }


        }

        /* TODO:
            *rowNames unique values only
            *add restart key for breakpoint selection
            *figure out how to handle inverter piles
            *split global piles into different command
         
         
         */
        public static void processBlockPiles()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            Database docDb = doc.Database;
            var docScale = docDb.Cannoscale.DrawingUnits;
            using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
            {
                //TinSurface surface;
                //surface.SampleElevations

                var (rowTypes, dimCounts, rowNames, globalLayers, filePath) = processRowCSV();

                CogoPointCollection cogoPoints = CivilApplication.ActiveDocument.CogoPoints;
                PromptIntegerOptions pio = new PromptIntegerOptions("power block number");
                pio.SetMessageAndKeywords("power block number or [Done]", "Done");
                var pbNum = doc.Editor.GetInteger(pio);
                if (pbNum.Status != PromptStatus.OK) { return; }

                while (pbNum.Status == PromptStatus.OK)
                {
                    //per iteration variables
                    double[] breakpoints;
                    System.Collections.Generic.Dictionary<double, List<Pile>>[] cogoDict;
                    var csv = new StringBuilder();
                    csv.AppendLine("Pile Number,Northern,Eastern,Pile Reveal,Description");
                    Dictionary<string, int> dimCountsPB = new Dictionary<string, int>();
                    foreach (string s in dimCounts) { dimCountsPB[s] = 0; }

                    SelectionFilter polylineFilter = createSelectionFilter(globalLayers, false);

                    breakpoints = getRowBreakpoints(doc).OrderByDescending(d => d).ToList().ToArray();

                    var PolyAndBlockSelectionOptions = new PromptSelectionOptions();
                    PolyAndBlockSelectionOptions.MessageForAdding = "\nSelect all objects belonging to power block, enter to continue";
                    var PolyAndBlockSelection = doc.Editor.GetSelection(PolyAndBlockSelectionOptions, polylineFilter);
                    if (PolyAndBlockSelection.Status != PromptStatus.OK)
                    {
                        return;
                    }
                    // select power block boundary
                    {
                        //var options = new PromptEntityOptions("\nSelect Power Block");
                        //var pb = doc.Editor.GetEntity(options);
                        //if (pb.Status != PromptStatus.OK)
                        //{
                        //    return;
                        //}

                        //var ent = (Polyline)transaction.GetObject(pb.ObjectId, OpenMode.ForRead);
                        //Point3dCollection powerBlockPoly = new Point3dCollection();
                        //Point3dCollection powerBlockPoly2 = new Point3dCollection();
                        //int vn = ent.NumberOfVertices;
                        //double[] polyToBuffer = new double[(vn*2)+2];
                        //PathsD polygon = new PathsD();
                        //PathsD buffered = new PathsD();


                        //for (int i = 0; i < vn; i++)

                        //{
                        //    Point3d pt = ent.GetPoint3dAt(i);
                        //    polyToBuffer[i*2] = pt.X;
                        //    polyToBuffer[(i*2)+1] = pt.Y;
                        //    powerBlockPoly.Add(pt);

                        //}
                        //polyToBuffer[vn * 2] = polyToBuffer[0];
                        //polyToBuffer[(vn*2)+1] = polyToBuffer[1];

                        //polygon.Add(Clipper.MakePath(polyToBuffer));
                        //buffered = Clipper.InflatePaths(polygon, 1,JoinType.Square,EndType.Square);

                        //for (int i = 0; i < buffered[0].Count; i++)
                        //{
                        //    powerBlockPoly2.Add(new Point3d(buffered[0][i].x, buffered[0][i].y, 0));
                        //}


                        //PromptSelectionResult blockBoundaryPrompt = doc.Editor.SelectWindowPolygon(powerBlockPoly2, polylineFilter);
                    }
                    //zoom to extent of dwg, needed for select window to work
                    doc.Editor.Command("Zoom", "extents");
                    //array of dictionary of piles, where the index of the dictionary in the array is the row #
                    //dictionary keys are the X values of the row, and values are Piles that reference cogo points
                    cogoDict = new Dictionary<double, List<Pile>>[breakpoints.Length + 1];
                    for (int i = 0; i < breakpoints.Length + 1; i++)
                        cogoDict[i] = new Dictionary<double, List<Pile>>();


                    //iterate through the selection of objects
                    for (int bb = 0; bb < PolyAndBlockSelection.Value.Count; bb++)
                    {
                        Autodesk.AutoCAD.DatabaseServices.DBObject obj = transaction.GetObject(PolyAndBlockSelection.Value[bb].ObjectId, OpenMode.ForRead);
                        //check for type, if polyline get standard pile blocks within, if block then process as global pile
                        Type objType = obj.GetType();
                        if (objType.Name == "Polyline")
                        {
                            Polyline blockBndy = (Polyline)obj;
                            //skip if the polyline is the power block boundary
                            if (blockBndy.Area > 5000) { continue; }
                            if (!rowTypes.Keys.Contains(blockBndy.Layer)) { continue; }
                            //get all blocks within rectangle blockBndy
                            PromptSelectionResult blockPrompta = doc.Editor.SelectWindow(blockBndy.Bounds.Value.MinPoint, blockBndy.Bounds.Value.MaxPoint, blockFilter);
                            //iterate through blocks in blockBndy and add to cogoDict
                            foreach (SelectedObject b in blockPrompta.Value)
                            {
                                var bl = (BlockReference)transaction.GetObject(b.ObjectId, OpenMode.ForRead);
                                ObjectId pointId = cogoPoints.Add(bl.Position, blockBndy.Layer, false);
                                //doing this because the double can differ in very small decimal places in piles of same row
                                double roundedX = Math.Round(bl.Position.X, 4);
                                int row = determinePBrow(bl.Position.Y, breakpoints, 0);
                                if (cogoDict[row].ContainsKey(roundedX))
                                {
                                    cogoDict[row][roundedX].Add(new Pile { Y = bl.Position.Y, pointId = pointId });
                                }
                                else
                                {

                                    cogoDict[row].Add(roundedX, new List<Pile> { new Pile { Y = bl.Position.Y, pointId = pointId } });
                                }


                            }
                        }
                        else if (objType.Name == "BlockReference")
                        {
                            BlockReference block = (BlockReference)obj;
                            ObjectId res = cogoPoints.Add(block.Position, block.Layer, false);
                            CogoPoint cogoPoint = res.GetObject(OpenMode.ForWrite) as CogoPoint;
                            cogoPoint.Layer = block.Layer;
                            cogoPoint.RawDescription = pbNum.Value.ToString();

                            dimCountsPB[rowTypes[block.Layer][1]]++;
                        }

                    }
                    //auto determine row breaks
                    {
                        //List<double> rowBreaks = new List<double>();
                        //foreach(KeyValuePair<double, List<Pile>> row in cogoDict)
                        //{
                        //    var sorted = row.Value.OrderBy(p => p.Y);
                        //    Pile prev = null;
                        //    foreach (Pile p in sorted)
                        //    {
                        //        if (prev == null)
                        //        {
                        //            prev = p;
                        //            continue;
                        //        }
                        //        if ((p.Y-prev.Y) > PWRBLOCKROWDIST)
                        //        {
                        //            rowBreaks.Add(p.Y);
                        //        }
                        //        prev = p;
                        //    }
                        //}

                        //var h = new HashSet<double>(rowBreaks);
                        //double[] arr2 = h.ToArray();
                    }
                    //iterate through the cogodict and reNumber the points correctly
                    for (int i = 0; i < breakpoints.Length + 1; i++)
                    {
                        var trackerRow = 1;
                        foreach (KeyValuePair<double, List<Pile>> row in cogoDict[i].OrderBy(key => key.Key))
                        {
                            var sorted = row.Value.OrderBy(p => p.Y);
                            var trackerRowIndex = 1;
                            foreach (var p in sorted)
                            {
                                CogoPoint cogoPoint = p.pointId.GetObject(OpenMode.ForWrite) as CogoPoint;
                                //cogoPoint.RawDescription = ent.Layer;
                                cogoPoint.Layer = cogoPoint.RawDescription;
                                cogoPoint.PointNumber = (uint)(trackerRowIndex + (trackerRow * 100) + (10000 * (i + 1)) + (100000 * pbNum.Value));
                                csv.AppendLine($"{cogoPoint.PointNumber},{row.Key},{p.Y},0,{rowTypes[cogoPoint.RawDescription][0]}");
                                dimCountsPB[rowTypes[cogoPoint.RawDescription][1]]++;
                                trackerRowIndex++;
                            }
                            trackerRow++;
                        }
                    }
                    //zoom back to previous extent when selection was made
                    doc.Editor.Command("Zoom", "previous");

                    File.WriteAllText($"{filePath}/PB{pbNum.Value}.csv", csv.ToString());

                    writeSummaryCSV(dimCountsPB, pbNum.Value.ToString(), filePath);

                    docDb.TransactionManager.QueueForGraphicsFlush();
                    pbNum = doc.Editor.GetInteger(pio);
                }


                doc.TransactionManager.EnableGraphicsFlush(true);
                doc.TransactionManager.FlushGraphics();
                doc.Editor.Regen();
                transaction.Commit();
            }
        }


    }
}
