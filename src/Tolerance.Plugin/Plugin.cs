using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Linq;

namespace LabelTest
{
    public class Plugin : IExtensionApplication
    {
        public void Initialize()
        {
        }

        public void Terminate()
        {
        }
        [CommandMethod("LabelLineSlope")]
        public void LabelSlopeCommand()
        {
            LabelText.LabelLineSlope();
        }
        
        [CommandMethod("LabelTextZ")]
        public void LabelCommand()
        {
            LabelText.LabelTextZ();
        }
        [CommandMethod("LabelContour")]
        public void LabelCont()
        {
            LabelText.LabelContour();
        }
        [CommandMethod("CHANGELAYERtest", CommandFlags.UsePickSet)]
        public static void ChangeLayerSample()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            // get the selected enities (return if none)
            var psr = ed.GetSelection();
            if (psr.Status != PromptStatus.OK)
                return;

            // start a transaction
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layerName = "foo";

                // check if the layer already exists
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
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
                }

                // set this layer to selected entites
                foreach (SelectedObject so in psr.Value)
                {
                    var ent = (Entity)tr.GetObject(so.ObjectId, OpenMode.ForWrite);
                    ent.Layer = layerName;
                }
                tr.Commit();
            }
        }

        [CommandMethod("TIV", CommandFlags.NoTileMode)]
        public void Test()
        {

        }
        [CommandMethod("plabel")]
        public void LabelPointElevations()
        {
            LabelText.PointElevationLabels();
        }
        
    }
}

namespace csvControl
{
    public class Plugin : IExtensionApplication
    {
        public void Initialize()
        {
        }

        public void Terminate()
        {
        }

        [CommandMethod("csvControl")]
        public void LabelCommand()
        {
            csvControl.LabelControl();
        }
    }
}

namespace GisImport
{
    public class Plugin : IExtensionApplication
    {
        public void Initialize()
        {
        }

        public void Terminate()
        {
        }

        [CommandMethod("importGIScsv")]
        public void LabelCommand()
        {
            GisDataImport.gisCSVtoBlocks();
        }
    }
}


namespace generalMethods
{
    public class Plugin : IExtensionApplication
    {
        public void Initialize()
        {
        }

        public void Terminate()
        {
        }

        [CommandMethod("lfind")]
        public void LabelCommand()
        {
            generalMethodsClass.FindLayer();
        }
        [CommandMethod("rt")]
        public void mTextCommand()
        {
            generalMethodsClass.MTextUp();
        }

        [CommandMethod("rm")]
        public void mLeaderCommand()
        {
            generalMethodsClass.MLeaderUp();
        }


    }
}

namespace solar
{
    public class Plugin : IExtensionApplication
    {
        public void Initialize()
        {
        }

        public void Terminate()
        {
        }

        [CommandMethod("block2Cogo")]
        public void LabelCommand()
        {
            solarFunctions.BlockToCogo();
        }
        [CommandMethod("processBlockPiles")]
        public void BlockPilesCommand()
        {
            solarFunctions.processBlockPiles();
        }
        [CommandMethod("processGlobalPiles")]
        public void GlobalPilesCommand()
        {
            solarFunctions.processGlobalPiles();
        }
        [CommandMethod("processPilesAuto")]
        public void PilesAutoCommand()
        {
            solarFunctionsAuto.processAllPilesAuto();
        }

        [CommandMethodAttribute("typeTest")]
        public void typeTest()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
            {
                //var options = new PromptEntityOptions("\nSelect contour to label");

                //var polylineResult = doc.Editor.GetEntity(options);
                //Autodesk.AutoCAD.DatabaseServices.DBObject obj = transaction.GetObject(polylineResult.ObjectId, OpenMode.ForRead);
                ////check for type, if polyline get standard pile blocks within, if block then process as global pile
                //Type objType = obj.GetType();
                CivilDocument doca = CivilApplication.ActiveDocument;
                Editor ed = doc.Editor;

                //https://help.autodesk.com/view/CIV3D/2024/ENU/?guid=d280ff4f-4c4b-f910-7a27-0e787a43448f
                PromptPointResult pointResult = ed.GetPoint("\nEnter a point to get elevation: ");
                if (pointResult.Status != PromptStatus.OK) return;
                Point3d point = pointResult.Value;

                var surfaceId = doca.GetSurfaceIds(); // Assuming the first surface in the drawing
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

                Application.ShowAlertDialog("Entered keyword: " +
                                            pKeyRes.StringResult);
                // Get elevation at the specified XY location
                //double elevation = surface.FindElevationAtXY(point.X, point.Y);
                //ed.WriteMessage($"\nThe elevation at X: {point.X}, Y: {point.Y} is: {elevation}");
                transaction.Commit();
            }
        }


    }
}

namespace template
{
    public class Plugin : IExtensionApplication
    {
        public void Initialize()
        {
        }

        public void Terminate()
        {
        }

        [CommandMethod("nameOfCommandToTypeIntoCad")]
        public void LabelCommand()
        {
            templateClass.mainFunctionName();
        }
    }
}
