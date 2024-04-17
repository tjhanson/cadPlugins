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
using SixLabors.Fonts.Unicode;
using Autodesk.Aec.Modeler;
using System.Runtime.InteropServices;

namespace solar
{

    public class solarFunctionsAuto
    {

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

            public NSPile(double x, double y, string layer, ObjectId pointId)
            {
                this.X = x;
                this.Y = y;
                //this.pb = pb;
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


        private static int determinePBrow(double val, double[] breakpoints,int ind)
        {
            if (ind == breakpoints.Length) return ind;
            if (val >= breakpoints[ind]) return ind;
            return determinePBrow(val,breakpoints, ind + 1);
        }

        private static void writeSummaryCSV(Dictionary<string, int> d,string pb,string filePath)
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
                    if(sl.ToString() == "TOTAL")
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
                workbook.SaveAs($"{filePath}/PB{pb}_Summary.xlsx");

            }
        } 

        private static ObjectId promptSurfaceSelection(Transaction transaction)
        {
            ObjectIdCollection surfaceId = CivilApplication.ActiveDocument.GetSurfaceIds();
            Dictionary<string, ObjectId> surfaces = new Dictionary<string, ObjectId>();
            Document acDoc = Application.DocumentManager.MdiActiveDocument;

            PromptKeywordOptions pKeyOpts = new PromptKeywordOptions("");
            pKeyOpts.Message = "\npick a surface ";


            pKeyOpts.AllowNone = false;


            for (int i = 0; i < surfaceId.Count; i++)
            {
                Autodesk.Civil.DatabaseServices.TinSurface surface = transaction.GetObject(surfaceId[i], OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.TinSurface;
                surfaces[surface.Name] = surfaceId[i];
                pKeyOpts.Keywords.Add(surface.Name);
            }
            PromptResult pKeyRes = acDoc.Editor.GetKeywords(pKeyOpts);

            return surfaces[pKeyRes.StringResult];
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
        
        private static (Dictionary<string, string[]>,List<string>, Dictionary<string, int>, List<string>,string) processRowCSV()
        {
            Dictionary<string, string[]> rowTypes = new Dictionary<string, string[]>();
            List<string> dimCounts = new List<string>();
            Dictionary<string,int> rowNames = new Dictionary<string, int>();
            List<string> globalLayers = new List<string>();
            PromptOpenFileOptions pofo = new PromptOpenFileOptions("\nEnter Row Type CSV File");
            pofo.PreferCommandLine = false;
            pofo.DialogName = "Select RowType CSV File";
            pofo.DialogCaption = "Select RowType CSV File";
            pofo.Filter = "csv |*.csv";

            PromptFileNameResult pfnr = Application.DocumentManager.MdiActiveDocument.Editor.GetFileNameForOpen(pofo);

            string FileName = pfnr.StringResult;
            string[] filepathsplit = pfnr.StringResult.Split('\\');
            Array.Resize(ref filepathsplit, filepathsplit.Length - 1);
            string filePath = String.Join("\\", filepathsplit)+"\\csv\\";
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
                    rowNames[line[9]] = Int32.Parse(line[10]) ;
                    globalLayers.Add(line[0]);
                }
                dimCounts.Add(line[3]);
                st = read.ReadLine();

            }
            Fstream.Dispose();
            
            return (rowTypes,dimCounts,rowNames,globalLayers,filePath);
        }
        
        private static SelectionFilter createSelectionFilter(List<string> layers, bool isGlobal)
        {
            List<TypedValue> gtvs;

            if (isGlobal)
            {
                 gtvs = new List<TypedValue> {
                    new TypedValue(Convert.ToInt32(DxfCode.Start), "INSERT"),//"AECC_COGO_POINT"),
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
        
        public static void createGlobalCogoPoints(Transaction transaction, CogoPointCollection cogoPoints, List<string> globalLayers, Dictionary<string,int> rowNames)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            //create point style that is just a dot
            ObjectId pointStyleId = CivilApplication.ActiveDocument.Styles.PointStyles.Add("tempPointStyle");
            PointStyle pointStyle = pointStyleId.GetObject(OpenMode.ForWrite) as PointStyle;
            pointStyle.MarkerType = PointMarkerDisplayType.UseCustomMarker;
            pointStyle.CustomMarkerStyle = CustomMarkerType.CustomMarkerDot;        
            //select all non standard piles
            SelectionFilter globalFilter = createSelectionFilter(globalLayers,true);
            List<NSPile> globalPiles = new List<NSPile>();
            var globalSelection = doc.Editor.SelectAll(globalFilter);
            //create cogo points 
            for (int bb = 0; bb < globalSelection.Value.Count; bb++)
            {
                BlockReference block = (BlockReference)transaction.GetObject(globalSelection.Value[bb].ObjectId, OpenMode.ForRead);
                ObjectId res = cogoPoints.Add(block.Position, block.Layer, true);
                globalPiles.Add(new NSPile(Math.Round(block.Position.X, 4), Math.Round(block.Position.Y, 4), block.Layer, res));
            }
            //number each cogo point
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
                    cogoPoint.Layer = cogoPoint.RawDescription;
                    cogoPoint.StyleId = pointStyleId;
                    cogoPoint.IsLabelVisible = false;
                    offset += 1;
                }
            }
            //flush the dwg so that that rest of function can use newly created cogo points
            transaction.TransactionManager.QueueForGraphicsFlush();
            doc.TransactionManager.FlushGraphics();
            doc.Editor.Regen();
        }

        
        private static SelectionFilter blockFilter = new SelectionFilter(new TypedValue[] { new TypedValue(Convert.ToInt32(DxfCode.Start), "INSERT")  });
        private static SelectionFilter pbBoundaries = new SelectionFilter(new TypedValue[] { new TypedValue(Convert.ToInt32(DxfCode.LayerName), "S-PB Outline") });
        private static SelectionFilter pbTextSF = new SelectionFilter(new TypedValue[] { new TypedValue(Convert.ToInt32(DxfCode.LayerName), "_PB-TEXT") });

        /* TODO:
            *
         
         70 hrs as of 4/17
         */
        public static void processAllPilesAuto()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            Database docDb = doc.Database;
            CogoPointCollection cogoPoints = CivilApplication.ActiveDocument.CogoPoints;

            //get row break detection distance
            //this value is used to determine the threshold for determining row breaks
            //for each column of blocks, if the difference between their Y values is greater, a row break is added
            var textHeightOptions = new PromptDistanceOptions("\nEnter row break detection distance (differences in Y values greater than this will be considered power block row breaks)");
            textHeightOptions.DefaultValue = 40;
            textHeightOptions.AllowNegative = false;
            var textHeight = doc.Editor.GetDistance(textHeightOptions);
            if (textHeight.Status != PromptStatus.OK)
            {
                return;
            }
            double PWRBLOCKROWDIST = textHeight.Value;

            var (rowTypes, dimCounts, rowNames, globalLayers, filePath) = processRowCSV();
            Dictionary<string, int> dimCountsPBtotal = new Dictionary<string, int>();
            foreach (string s in dimCounts) { dimCountsPBtotal[s] = 0; }

            SelectionFilter polylineFilter = createSelectionFilter(globalLayers, false);

            using (Transaction transaction = docDb.TransactionManager.StartTransaction())
            {
                //get surface for point z values
                Autodesk.Civil.DatabaseServices.TinSurface surface = transaction.GetObject(promptSurfaceSelection(transaction), OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.TinSurface;

                //needed for selection window to work correctly
                //if erroring out, check the dwg entents for objects super far away
                doc.Editor.Command("Zoom", "extents");
                
                createGlobalCogoPoints(transaction, cogoPoints, globalLayers, rowNames);

                
                //get all power block boundaries
                var pbSelection = doc.Editor.SelectAll(pbBoundaries);

                ProgressMeter pm = new ProgressMeter();
                pm.Start("processing");
                pm.SetLimit(pbSelection.Value.Count);

                for (int bb = 0; bb < pbSelection.Value.Count; bb++)
                {
                    pm.MeterProgress();
                    pm.
                    System.Windows.Forms.Application.DoEvents();
                    //per iteration variables
                    Dictionary<double, List<Pile>> cogoDict = new Dictionary<double, List<Pile>>();
                    Dictionary<string, int> dimCountsPB = new Dictionary<string, int>();
                    foreach (string s in dimCounts) { dimCountsPB[s] = 0; }
                    var csv = new StringBuilder();
                    csv.AppendLine("Pile Number,Northern,Eastern,Elevation,Description");

                    //generate a polygon boundary for selection
                    Polyline ent = (Polyline)transaction.GetObject(pbSelection.Value[bb].ObjectId, OpenMode.ForRead);
                    Point3dCollection powerBlockPoly = new Point3dCollection();

                    for (int i = 0; i < ent.NumberOfVertices; i++)
                    {
                        Point3d pt = ent.GetPoint3dAt(i);
                        powerBlockPoly.Add(pt);
                    }
                    //get the number of the power block
                    PromptSelectionResult pbText = doc.Editor.SelectWindowPolygon(powerBlockPoly, pbTextSF);
                    MText pbNumM = (MText)transaction.GetObject(pbText.Value[0].ObjectId, OpenMode.ForRead);
                    int pbNumb = Int32.Parse(pbNumM.Text.Split()[1]);

                    //gets all the polylines and global piles cogo points withing pb boundary
                    PromptSelectionResult pbContents = doc.Editor.SelectWindowPolygon(powerBlockPoly, polylineFilter);
                    //iterate through the selection of objects
                    for (int i = 0; i < pbContents.Value.Count; i++)
                    {
                        Autodesk.AutoCAD.DatabaseServices.DBObject obj = transaction.GetObject(pbContents.Value[i].ObjectId, OpenMode.ForRead);
                        //check for type, if polyline get standard pile blocks within, if block then process as global pile
                        Type objType = obj.GetType();
                        if (objType.Name == "Polyline")
                        {
                            Polyline blockBndy = (Polyline)obj;
                            //is this still needed? skip if the polyline is the power block boundary
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
                                double roundedX = Math.Round(bl.Position.X, 3);
                                double roundedY = Math.Round(bl.Position.Y, 3);
                                if (cogoDict.ContainsKey(roundedX))
                                { cogoDict[roundedX].Add(new Pile { Y = roundedY, pointId = pointId }); }

                                else
                                {cogoDict.Add(roundedX, new List<Pile> { new Pile { Y = roundedY, pointId = pointId } }); }
                            }
                        }
                        else if (objType.Name == "CogoPoint")
                        {
                            CogoPoint cp = (CogoPoint)obj;
                            //turn back on the point number labeling
                            cp.IsLabelVisible = true;
                            dimCountsPB[rowTypes[cp.Layer][1]]++;
                            double elevation = surface.FindElevationAtXY(cp.Location.X, cp.Location.Y);
                            csv.AppendLine($"{cp.PointNumber},{cp.Location.X},{cp.Location.Y},{Math.Round(elevation, 3)},{rowTypes[cp.RawDescription][0]}");
                        }
                    }
                    //find power block row breaks
                    List<double> rowBreaks = new List<double>();
                    foreach (KeyValuePair<double, List<Pile>> row in cogoDict)
                    {
                        var sorted = row.Value.OrderBy(p => p.Y).ToArray();
                        Pile prev = sorted[0];
                        foreach (Pile p in sorted)
                        {
                            if ((p.Y - prev.Y) > PWRBLOCKROWDIST)
                            {
                                rowBreaks.Add(p.Y);
                            }
                            prev = p;
                        }
                    }
                    var h = new HashSet<double>(rowBreaks);
                    double[] arr2 = h.OrderByDescending(d => d).ToList().ToArray();

                    //array of dictionaries of X values as keys, with the corresponding piles in a list. array index is row value
                    Dictionary<double, List<Pile>> [] cogoDictRows = new Dictionary<double, List<Pile>>[arr2.Length + 1];
                    for (int i = 0; i < arr2.Length + 1; i++)
                        cogoDictRows[i] = new Dictionary<double, List<Pile>>();

                    foreach (KeyValuePair<double, List<Pile>> r in cogoDict)
                    {
                        foreach (Pile p in r.Value)
                        {
                            int row = determinePBrow(p.Y, arr2, 0);
                            if (cogoDictRows[row].ContainsKey(r.Key))
                            { cogoDictRows[row][r.Key].Add(p); }
                            else
                            {cogoDictRows[row].Add(r.Key, new List<Pile> { p });}
                        }
                    }
                    //number each cogo point and compute csv values
                    for (int i = 0; i < arr2.Length + 1; i++)
                    {
                        var trackerRow = 1;
                        foreach (KeyValuePair<double, List<Pile>> row in cogoDictRows[i].OrderBy(key => key.Key))
                        {
                            var sorted = row.Value.OrderBy(p => p.Y);
                            var trackerRowIndex = 1;
                            foreach (var p in sorted)
                            {
                                CogoPoint cogoPoint = p.pointId.GetObject(OpenMode.ForWrite) as CogoPoint;
                                //cogoPoint.RawDescription = ent.Layer;
                                cogoPoint.Layer = cogoPoint.RawDescription;
                                cogoPoint.PointNumber = (uint)(trackerRowIndex + (trackerRow * 100) + (10000 * (i + 1)) + (100000 * pbNumb));
                                double elevation = surface.FindElevationAtXY(cogoPoint.Location.X, cogoPoint.Location.Y);
                                csv.AppendLine($"{cogoPoint.PointNumber},{row.Key},{p.Y},{Math.Round(elevation,3)},{rowTypes[cogoPoint.RawDescription][0]}");
                                dimCountsPB[rowTypes[cogoPoint.RawDescription][1]]++;
                                trackerRowIndex++;
                            }
                            trackerRow++;
                        }
                    }
                    //write csvs
                    File.WriteAllText($"{filePath}/PB{pbNumb}.csv", csv.ToString());
                    writeSummaryCSV(dimCountsPB, pbNumb.ToString(), filePath);
                    //update site summary count
                    foreach (KeyValuePair<string,int> s in dimCountsPB) { dimCountsPBtotal[s.Key] = s.Value+dimCountsPBtotal[s.Key]; }
                    //idk if this is needed
                    doc.Editor.Regen();
                }
                pm.Stop();
                //write site summary count csv
                writeSummaryCSV(dimCountsPBtotal, "total", filePath);
                doc.TransactionManager.EnableGraphicsFlush(true);
                doc.TransactionManager.FlushGraphics();
                doc.Editor.Regen();
                transaction.Commit();
            }
        }
    }
}
