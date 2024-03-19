using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using AutoCAD;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil.DatabaseServices;
using static System.Net.Mime.MediaTypeNames;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace LabelTest
{

    public static class ListExtensions
    {
        public static List<List<T>> ChunkBy<T>(this List<T> source, int chunkSize)
        {
            return source
                .Select((x, i) => new { Index = i, Value = x })
                .GroupBy(x => x.Index / chunkSize)
                .Select(x => x.Select(v => v.Value).ToList())
                .ToList();
        }
    }
    class BlockJig : EntityJig
    {
        BlockReference br;
        Polyline pline;
        Point3d dragPt;
        Plane plane;
        Dictionary<AttributeReference, TextInfo> attInfos;

        public BlockJig(BlockReference br, Polyline pline, Dictionary<AttributeReference, TextInfo> attInfos) : base(br)
        {
            this.br = br;
            this.pline = pline;
            this.attInfos = attInfos;
            plane = new Plane(Point3d.Origin, pline.Normal);
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            var options = new JigPromptPointOptions("\nSpecicfy insertion point: ");
            options.UserInputControls =
                UserInputControls.Accept3dCoordinates |
                UserInputControls.UseBasePointElevation;
            var result = prompts.AcquirePoint(options);
            if (result.Value.IsEqualTo(dragPt))
                return SamplerStatus.NoChange;
            dragPt = result.Value;
            return SamplerStatus.OK;
        }

        protected override bool Update()
        {
            var point = pline.GetClosestPointTo(dragPt, false);
            var angle = pline.GetFirstDerivative(point).AngleOnPlane(plane);
            br.Position = point;
            br.Rotation = angle;
            foreach (var entry in attInfos)
            {
                var att = entry.Key;
                var info = entry.Value;
                att.Position = info.Position.TransformBy(br.BlockTransform);
                att.Rotation = info.Rotation + angle;
                if (info.IsAligned)
                {
                    att.AlignmentPoint = info.Alignment.TransformBy(br.BlockTransform);
                    att.AdjustAlignment(br.Database);
                }
                if (att.IsMTextAttribute)
                {
                    att.UpdateMTextAttribute();
                }
            }
            return true;
        }
    }


    class TextInfo
    {
        public Point3d Position { get; }

        public Point3d Alignment { get; }

        public bool IsAligned { get; }

        public double Rotation { get; }

        public TextInfo(DBText text)
        {
            Position = text.Position;
            IsAligned = text.Justify != AttachmentPoint.BaseLeft;
            Alignment = text.AlignmentPoint;
            Rotation = text.Rotation;
        }
    }

    class TextJig : EntityJig
    {
        DBText text;
        Polyline pline;
        Point3d dragPt;
        Plane plane;
        Database db;

        public TextJig(DBText text, Polyline pline) : base(text)
        {
            this.text = text;
            this.pline = pline;
            plane = new Plane(Point3d.Origin, pline.Normal);
            db = HostApplicationServices.WorkingDatabase;
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            var options = new JigPromptPointOptions("\nSpecicfy insertion point: ");
            options.UserInputControls =
                UserInputControls.Accept3dCoordinates |
                UserInputControls.UseBasePointElevation;
            var result = prompts.AcquirePoint(options);
            if (result.Value.IsEqualTo(dragPt))
                return SamplerStatus.NoChange;
            dragPt = result.Value;
            return SamplerStatus.OK;
        }

        protected override bool Update()
        {
            var point = pline.GetClosestPointTo(dragPt, false);
            var angle = pline.GetFirstDerivative(point).AngleOnPlane(plane);
            //if angle greater than 90 or less than 270, add 180
            if (angle > 1.5708 && angle < 4.71239)
            {
                angle = angle + 3.14159;
            }

            text.AlignmentPoint = point;
            text.Rotation = angle;
            text.AdjustAlignment(db);
            return true;
        }
    }

    class TextPlacementJig : EntityJig

    {

        // Declare some internal state



        Database _db;

        Transaction _tr;

        Point3d _position;

        double _angle, _txtSize;



        // Constructor



        public TextPlacementJig(

          Transaction tr, Database db, MText ent

        ) : base(ent)

        {

            _db = db;

            _tr = tr;


        }



        protected override SamplerStatus Sampler(

          JigPrompts jp

        )

        {

            JigPromptPointOptions po =

              new JigPromptPointOptions(

                "\nPosition of text"

              );

            po.UserInputControls =
              (UserInputControls.Accept3dCoordinates |
                UserInputControls.NullResponseAccepted |
                UserInputControls.NoNegativeResponseAccepted |
                UserInputControls.GovernedByOrthoMode);

            PromptPointResult ppr = jp.AcquirePoint(po);

            if (ppr.Status == PromptStatus.OK)
            {
                // Check if it has changed or not (reduces flicker)

                if (

                  _position.DistanceTo(ppr.Value) <

                    Tolerance.Global.EqualPoint

                )

                    return SamplerStatus.NoChange;



                _position = ppr.Value;

                return SamplerStatus.OK;

            }



            return SamplerStatus.Cancel;

        }



        protected override bool Update()

        {

            // Set properties on our text object



            MText txt = (MText)Entity;



            txt.Location = _position;
            //txt.TextHeight = 2;
            return true;

        }

    }



public static class Extension
    {
        /// <summary>
        /// Gets the transformation matrix of the display coordinate system (DCS)
        /// of the specified window to the world coordinate system (WCS).
        /// </summary>
        /// <param name="vp">The instance to which this method applies.</param>
        /// <returns>The transformation matrix from DCS to WCS.</returns>
        public static Matrix3d DCS2WCS(this Viewport vp)
        {
            return
                Matrix3d.Rotation(-vp.TwistAngle, vp.ViewDirection, vp.ViewTarget) *
                Matrix3d.Displacement(vp.ViewTarget.GetAsVector()) *
                Matrix3d.PlaneToWorld(vp.ViewDirection);
        }

        /// <summary>
        /// Gets the transformation matrix of the world coordinate system (WCS)
        /// to the display coordinate system (DCS) of the specified window.
        /// </summary>
        /// <param name="vp">The instance to which this method applies.</param>
        /// <returns>The transformation matrix from WCS to DCS.</returns>
        public static Matrix3d WCS2DCS(this Viewport vp)
        {
            return
                Matrix3d.WorldToPlane(vp.ViewDirection) *
                Matrix3d.Displacement(vp.ViewTarget.GetAsVector().Negate()) *
                Matrix3d.Rotation(vp.TwistAngle, vp.ViewDirection, vp.ViewTarget);
        }

        /// <summary>
        /// Gets the transformation matrix of the display coordinate system of the specified paper space window (DCS)
        /// to the paper space display coordinate system (PSDCS).
        /// </summary>
        /// <param name="vp">The instance to which this method applies.</param>
        /// <returns>The transformation matrix from DCS to PSDCS.</returns>
        public static Matrix3d DCS2PSDCS(this Viewport vp)
        {
            return
                Matrix3d.Scaling(vp.CustomScale, vp.CenterPoint) *
                Matrix3d.Displacement(vp.ViewCenter.Convert3d().GetVectorTo(vp.CenterPoint));
        }

        /// <summary>
        /// Gets the transformation matrix of the paper space display coordinate system (PSDCS)
        /// to the display coordinate system of the specified paper space window (DCS). 
        /// </summary>
        /// <param name="vp">The instance to which this method applies.</param>
        /// <returns>The transformation matrix from PSDCS to DCS.</returns>
        public static Matrix3d PSDCS2DCS(this Viewport vp)
        {
            return
                Matrix3d.Displacement(vp.CenterPoint.GetVectorTo(vp.ViewCenter.Convert3d())) *
                Matrix3d.Scaling(1.0 / vp.CustomScale, vp.CenterPoint);
        }

        /// <summary>
        /// Converts a Point2d into a Point3d with a Z coordinate equal to 0.
        /// </summary>
        /// <param name="pt">The instance to which this method applies.</param>
        /// <returns>The corresponding Point3d.</returns>
        public static Point3d Convert3d(this Point2d pt) =>
            new Point3d(pt.X, pt.Y, 0.0);
    }
    public class LabelText
    {
        public static void addTextLabel(double x, double y, double elevation, double scaler, BlockTableRecord btr, Transaction transaction)
        {
            using (DBText acText = new DBText())
            {
                acText.Height = scaler * .1;
                acText.Position = new Point3d(x,y,elevation);
                acText.TextString = Math.Round(Convert.ToDecimal(elevation), 2, MidpointRounding.AwayFromZero).ToString();
                //acText.Layer = "E_SPOT";

                btr.AppendEntity(acText);
                transaction.AddNewlyCreatedDBObject(acText, true);
            }           
        }

        public static Point3d GetOffset(Autodesk.AutoCAD.DatabaseServices.DBObject point, double offsetX, double offsetY)
        {
            var pointX = (point.Bounds.Value.MaxPoint.X + point.Bounds.Value.MinPoint.X) / 2;
            var pointY = (point.Bounds.Value.MaxPoint.Y + point.Bounds.Value.MinPoint.Y) / 2;
            return new Point3d(pointX+offsetX,pointY+offsetY, point.Bounds.Value.MinPoint.Z);
        }
        public static void LabelContour()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database docDb = doc.Database;
            var ed = doc.Editor;
            var docScale = docDb.Cannoscale.DrawingUnits;
            using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
            {


                

                BlockTable acBlkTbl;
                acBlkTbl = transaction.GetObject(docDb.BlockTableId,
                                                OpenMode.ForRead) as BlockTable;
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = transaction.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;


                var textHeightOptions = new PromptDistanceOptions("\nEnter Text Height (Default is .1 x Document Scale)");
                textHeightOptions.DefaultValue = docScale*.1;
                textHeightOptions.AllowNegative = false;
                var textHeight = doc.Editor.GetDistance(textHeightOptions);
                if (textHeight.Status != PromptStatus.OK)
                {
                    return;
                }

                var options = new PromptEntityOptions("\nSelect contour to label");
                options.SetRejectMessage("\nSelected object is no a Polyline.");
                options.AddAllowedClass(typeof(Polyline), true);
                var polylineResult = doc.Editor.GetEntity(options);
                if (polylineResult.Status != PromptStatus.OK)
                {
                    return;
                }
                while (polylineResult.Status == PromptStatus.OK)
                {
                    var pline = (Polyline)transaction.GetObject(polylineResult.ObjectId, OpenMode.ForRead);
                    
                    using (var text = new DBText())
                    {
                        text.SetDatabaseDefaults();
                        text.Normal = pline.Normal;
                        text.Justify = AttachmentPoint.MiddleCenter;
                        text.AlignmentPoint = Point3d.Origin;
                        text.TextString = pline.Elevation.ToString();
                        text.Height = textHeight.Value;
                        //text.Layer = "VA-SURF-MAJR-LABL";


                        var jig = new TextJig(text, pline);
                        var result = ed.Drag(jig);
                        if (result.Status == PromptStatus.OK)
                        {
                            var currentSpace = (BlockTableRecord)transaction.GetObject(docDb.CurrentSpaceId, OpenMode.ForWrite);
                            currentSpace.AppendEntity(text);
                            transaction.AddNewlyCreatedDBObject(text, true);
                        }
                    }
                    docDb.TransactionManager.QueueForGraphicsFlush();
                    //doc.Editor.Regen();
                    
                    polylineResult = doc.Editor.GetEntity(options);

                }
                transaction.Commit();

                //addTextLabel(x, y, contour.Elevation, docScale, acBlkTblRec, transaction);


            }
        }
        public static void LabelLineSlope()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layouts = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                var layout = (Layout)tr.GetObject(layouts.GetAt(LayoutManager.Current.CurrentLayout), OpenMode.ForRead);
                short cvport = (short)Application.GetSystemVariable("CVPORT");
                Viewport vp;
                if (cvport == 1)
                {
                    var peo = new PromptEntityOptions("\nSelect viewport: ");
                    peo.SetRejectMessage("\nSelected object is not a viewport.");
                    peo.AddAllowedClass(typeof(Viewport), true);
                    var per = ed.GetEntity(peo);
                    if (per.Status != PromptStatus.OK)
                        return;
                    vp = (Viewport)tr.GetObject(per.ObjectId, OpenMode.ForWrite);
                    ed.SwitchToModelSpace();
                    
                    Application.SetSystemVariable("CVPORT", vp.Number);
                }
                else
                {
                    vp = (Viewport)tr.GetObject(ed.ActiveViewportId, OpenMode.ForWrite);
                }

                var xform = vp.WCS2DCS() * vp.PSDCS2DCS();

                

                BlockTable acBlkTbl;
                acBlkTbl = tr.GetObject(db.BlockTableId,
                                                OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = tr.GetObject(acBlkTbl[BlockTableRecord.PaperSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;


                var pointsPromptOptions = new PromptSelectionOptions();
                pointsPromptOptions.MessageForAdding = "\nSelect points to label";
                var psr = doc.Editor.GetSelection();
                if (psr.Status != PromptStatus.OK)
                {
                    return;
                }
                while (psr.Status == PromptStatus.OK)
                {


                    foreach (SelectedObject so in psr.Value)
                    {
                        var ent = (Autodesk.AutoCAD.DatabaseServices.Entity)tr.GetObject(so.ObjectId, OpenMode.ForRead);
                        var pt1 = ent.Bounds.Value.MinPoint.Z.ToString();
                        var min = ent.Bounds.Value.MinPoint;
                        var max = ent.Bounds.Value.MaxPoint;
                        double dist = Math.Sqrt((Math.Pow((min.Y - max.Y), 2) + Math.Pow((min.X - max.X), 2)));
                        double slope = Math.Abs((min.Z - max.Z) / dist);

                        ed.SwitchToPaperSpace();
                        ed.Command("id");
                        List<string> cmdLineList = Autodesk.AutoCAD.Internal.Utils.GetLastCommandLines(20, true);
                        string[] idCoords = cmdLineList[0].Split();
                        //"id Specify point: Specify point:  X = 23.91     Y = 7.75     Z = 0.00"

                        using (DBText acText = new DBText())
                        {
                            acText.Position = new Point3d(Convert.ToDouble(idCoords[8]), Convert.ToDouble(idCoords[15]), 0);
                            acText.Height = .1;
                            //acText.Normal = new Vector3d(1, 0, 0);
                            acText.Rotation = 0;
                            
                            acText.TextString = Math.Round(slope,2, MidpointRounding.AwayFromZero).ToString() + ":1";


                            acBlkTblRec.AppendEntity(acText);
                            tr.AddNewlyCreatedDBObject(acText, true);
                        }

                    }
                    db.TransactionManager.QueueForGraphicsFlush();
                    ed.SwitchToModelSpace();
                    psr = doc.Editor.GetSelection();
                }

                tr.Commit();


            }
        }

        public static void PointElevationLabels()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            Database docDb = doc.Database;
            var docScale = docDb.Cannoscale.DrawingUnits;
            using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
            {
                BlockTable acBlkTbl;
                acBlkTbl = transaction.GetObject(docDb.BlockTableId,
                                                OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = transaction.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;
                var dict = docDb.MLeaderStyleDictionaryId.GetObject(OpenMode.ForRead) as DBDictionary;
                var textdict = (TextStyleTable)transaction.GetObject(docDb.TextStyleTableId, OpenMode.ForRead);

                var mlstyle = dict.GetAt("_FE Text Elevation");
                //var tStyle = textdict.getAt("romanS");

                var textLabels = new List<string>();

                var options = new PromptEntityOptions("\nSelect points to add to mleader in order of appearance, enter to finish");
                options.SetRejectMessage("\nSelected object is no a cogo point.");
                options.AddAllowedClass(typeof(CogoPoint), true);
                var pointResult = doc.Editor.GetEntity(options);
                if (pointResult.Status != PromptStatus.OK)
                {
                    return;
                }
                while (pointResult.Status == PromptStatus.OK)
                {
                    var point = (CogoPoint)transaction.GetObject(pointResult.ObjectId, OpenMode.ForRead);
                    textLabels.Add(point.Elevation.ToString()+" "+point.FullDescription);

                    pointResult = doc.Editor.GetEntity(options);
                }

                var mlText = new MText();
                mlText.Contents = String.Join("\n", textLabels);
                mlText.TextHeight = docScale * .1;
                mlText.TextStyleId = textdict["romanS"];

                TextPlacementJig pj = new TextPlacementJig(transaction, docDb, mlText);
                PromptStatus stat = PromptStatus.Keyword;

                while (stat == PromptStatus.Keyword)

                {

                    PromptResult res = doc.Editor.Drag(pj);

                    stat = res.Status;

                    if (

                      stat != PromptStatus.OK &&

                      stat != PromptStatus.Keyword

                    )

                        return;

                }


                //var ptoptions = new PromptPointOptions("\nSelect text insertion location: ");
                //var mlTIns = doc.Editor.GetPoint(ptoptions).Value;
                var pmloptions = new PromptPointOptions("\nSelect the Mleader insertion point: ");
                var mlIns = doc.Editor.GetPoint(pmloptions).Value;

                
                

                    using (var mldr = new MLeader())
                {
                    mldr.MLeaderStyle = mlstyle;
                    //mldr.TextStyleId = tStyle;
                    var index = mldr.AddLeaderLine(new Point3d(0, 0, 0));
                    mldr.AddFirstVertex(index, mlIns);
                    mldr.ContentType = ContentType.MTextContent;
                    mldr.MText = mlText;
                    mldr.TextLocation = mlText.Location;


                    acBlkTblRec.AppendEntity(mldr);
                    transaction.AddNewlyCreatedDBObject(mldr, true);
                    
                }

                doc.TransactionManager.EnableGraphicsFlush(true);
                doc.TransactionManager.FlushGraphics();
                doc.Editor.Regen();
                transaction.Commit();
            }
        }

        public static void LabelTextZ()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            Database docDb = doc.Database;
            var docScale = docDb.Cannoscale.DrawingUnits;
            using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
            {
                var offsetXPromptOptions = new PromptDistanceOptions("\nOffset X");
                offsetXPromptOptions.DefaultValue = 5.0;
                offsetXPromptOptions.AllowNegative = true;
                var offsetXPromptResult = doc.Editor.GetDistance(offsetXPromptOptions);
                if (offsetXPromptResult.Status != PromptStatus.OK)
                {
                    return;
                }

                var offsetYPromptOptions = new PromptDistanceOptions("\nOffset Y");
                offsetYPromptOptions.DefaultValue = -5.0;
                offsetYPromptOptions.AllowNegative = true;
                var offsetYPromptResult = doc.Editor.GetDistance(offsetYPromptOptions);
                if (offsetYPromptResult.Status != PromptStatus.OK)
                {
                    return;
                }
                var textHeightPromptOptions = new PromptDistanceOptions("\nText Height");
                textHeightPromptOptions.DefaultValue = .1*docScale;
                textHeightPromptOptions.AllowNegative = true;
                var textHeightPromptResult = doc.Editor.GetDistance(textHeightPromptOptions);
                if (textHeightPromptResult.Status != PromptStatus.OK)
                {
                    return;
                }


                var pointsPromptOptions = new PromptSelectionOptions();
                pointsPromptOptions.MessageForAdding = "\nSelect points to label";
                var psr = doc.Editor.GetSelection();
                if (psr.Status != PromptStatus.OK)
                {
                    return;
                }



                var layerName = "blocksToDelete";

                // check if the layer already exists
                var lt = (LayerTable)transaction.GetObject(docDb.LayerTableId, OpenMode.ForRead);
                if (!lt.Has(layerName))
                {
                    // if not create it
                    var layer = new LayerTableRecord()
                    {
                        Name = layerName,
                        Color = Color.FromRgb(200, 30, 80)
                    };
                    lt.UpgradeOpen();
                    lt.Add(layer);
                    transaction.AddNewlyCreatedDBObject(layer, true);
                }

                BlockTable acBlkTbl;
                acBlkTbl = transaction.GetObject(docDb.BlockTableId,
                                                OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = transaction.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;

                //var pointsIds = pointsSelectionResult.Value.GetObjectIds().ToList();
                List<ObjectId> pointsToMove = new List<ObjectId>();
                foreach (SelectedObject so in psr.Value)
                {
                    var ent = (Autodesk.AutoCAD.DatabaseServices.Entity)transaction.GetObject(so.ObjectId, OpenMode.ForWrite);

                    if (ent is BlockReference)
                    {
                        String textElevation = Math.Round(ent.Bounds.Value.MinPoint.Z, 1, MidpointRounding.AwayFromZero).ToString();
                        if (!textElevation.Contains('.'))
                        {
                            //pointsToMove.Add(pointsId);
                            //Entity pointToMove = transaction.GetObject(pointsId, OpenMode.ForWrite) as Entity;
                            ent.Layer = layerName;
                        }

                        using (DBText acText = new DBText())
                        {
                            acText.Position = GetOffset(ent, offsetXPromptResult.Value, offsetYPromptResult.Value);
                            acText.Height = textHeightPromptResult.Value;

                            acText.TextString = textElevation;
                            if (!textElevation.Contains('.'))
                            {
                                acText.Layer = layerName;
                            }

                            acBlkTblRec.AppendEntity(acText);
                            transaction.AddNewlyCreatedDBObject(acText, true);
                        }

                    }

                }

                doc.TransactionManager.EnableGraphicsFlush(true);
                doc.TransactionManager.FlushGraphics();
                doc.Editor.Regen();
                transaction.Commit();
            }
        }
    }
}
