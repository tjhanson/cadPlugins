using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
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
