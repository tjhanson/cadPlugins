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
using Autodesk.AutoCAD.Windows;
using Autodesk.Civil;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using PvGrade;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;
using Clipper2Lib;

namespace solar
{

    public class solarFunctions
    {
        

        public static void BlockToCogo()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
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

                    doc.Editor.Command("Zoom","all");
                    SelectionFilter blockFilter = new SelectionFilter(bf);
                    foreach (SelectedObject so in acSSPrompt.Value)
                    {

                        var ent = (Polyline)transaction.GetObject(so.ObjectId, OpenMode.ForRead);
                        PromptSelectionResult blockPrompt;
                        
                        blockPrompt = doc.Editor.SelectWindow( ent.Bounds.Value.MinPoint, ent.Bounds.Value.MaxPoint, blockFilter);
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
                    Application.ShowAlertDialog("Number of objects selected: 0; contact Taylor if this isnt intended");
                }

                


                



                doc.TransactionManager.EnableGraphicsFlush(true);
                doc.TransactionManager.FlushGraphics();
                doc.Editor.Regen();
                transaction.Commit();
            }
        }
        public class Pile
        {
            public double Y { get; set; }

            public ObjectId pointId { get; set; }
            
        }

        class PileComparer : IComparer<Pile>
        {
            public int Compare(Pile a, Pile b)
            {
                if (a.Y < b.Y) return 0;
                return -1;
            }
        }
        /* TODO:
            *buffer polygon
                *in c# or preprocess in arcgis?
            *truncate northing coord to 4? decimal places 
            *create algo for determining row within power block
            *get power block number from anno or ask user for input
            *double check cogo point # when changing
            *
            *figure out how to handle global piles
         
         
         */
        public static void processBlockPiles()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            Database docDb = doc.Database;
            var docScale = docDb.Cannoscale.DrawingUnits;
            using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
            {

                CogoPointCollection cogoPoints = CivilApplication.ActiveDocument.CogoPoints;
                var options = new PromptEntityOptions("\nSelect Power Block");
                var pb = doc.Editor.GetEntity(options);
                if (pb.Status != PromptStatus.OK)
                {
                    return;
                }

                TypedValue[] bf = new TypedValue[] { new TypedValue(Convert.ToInt32(DxfCode.Start), "INSERT"), };
                TypedValue[] tvs = new TypedValue[] {
                    new TypedValue(Convert.ToInt32(DxfCode.Operator), "<or"),
                    new TypedValue(Convert.ToInt32(DxfCode.Start), "POLYLINE"),
                    new TypedValue(Convert.ToInt32(DxfCode.Start), "LWPOLYLINE"),
                    new TypedValue(Convert.ToInt32(DxfCode.Start), "POLYLINE2D"),
                    new TypedValue(Convert.ToInt32(DxfCode.Start), "POLYLINE3d"),
                    new TypedValue(Convert.ToInt32(DxfCode.Operator), "or>"),
                };
                SelectionFilter polylineFilter = new SelectionFilter(tvs);

                doc.Editor.Command("Zoom", "all");
                SelectionFilter blockFilter = new SelectionFilter(bf);

                var ent = (Polyline)transaction.GetObject(pb.ObjectId, OpenMode.ForRead);
                Point3dCollection powerBlockPoly = new Point3dCollection();
                int vn = ent.NumberOfVertices;
                //double[] polyToBuffer = [];
                for (int i = 0; i < vn; i++)

                {

                    // Could also get the 3D point here

                    Point3d pt = ent.GetPoint3dAt(i);
                    powerBlockPoly.Add(pt);

                }
                //powerBlockPoly.Add(ent.GetPoint3dAt(0));
                
                PromptSelectionResult blockBoundaryPrompt = doc.Editor.SelectWindowPolygon(powerBlockPoly, polylineFilter);

                System.Collections.Generic.Dictionary<double, List<Pile>> cogoDict = new Dictionary<double, List<Pile>>();

                foreach (SelectedObject so in blockBoundaryPrompt.Value)
                {
                    var blockBndy = (Polyline)transaction.GetObject(so.ObjectId, OpenMode.ForRead);
                    PromptSelectionResult blockPrompta = doc.Editor.SelectWindow(blockBndy.Bounds.Value.MinPoint, blockBndy.Bounds.Value.MaxPoint, blockFilter);

                    foreach (SelectedObject b in blockPrompta.Value)
                    {
                        var bl = (BlockReference)transaction.GetObject(b.ObjectId, OpenMode.ForRead);
                        ObjectId pointId = cogoPoints.Add(bl.Position, blockBndy.Layer, false); ;
                        if (cogoDict.ContainsKey(bl.Position.X))
                        {
                            cogoDict[bl.Position.X].Add(new Pile { Y = bl.Position.Y, pointId = pointId });
                        }
                        else
                        {
                            
                            cogoDict.Add(bl.Position.X, new List<Pile > { new Pile { Y = bl.Position.Y, pointId = pointId } });
                        }


                    }
                    
                }
                var trackerRow = 0;
                foreach (KeyValuePair<double, List<Pile>> row in cogoDict.OrderBy(key => key.Key))
                {
                    var sorted = row.Value.OrderBy(p => p.Y);
                    var trackerRowIndex = 0;
                    foreach (var p in sorted)
                    {
                        CogoPoint cogoPoint = p.pointId.GetObject(OpenMode.ForWrite) as CogoPoint;
                        //cogoPoint.RawDescription = ent.Layer;
                        cogoPoint.Layer = cogoPoint.RawDescription;
                        cogoPoint.PointNumber = (uint)(trackerRowIndex + (trackerRow * 100)+10000+100000);
                        trackerRowIndex++;
                    }
                    trackerRow++;
                }
                

                doc.TransactionManager.EnableGraphicsFlush(true);
                doc.TransactionManager.FlushGraphics();
                doc.Editor.Regen();
                transaction.Commit();
            }
        }
    }
}
