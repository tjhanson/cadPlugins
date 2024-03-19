using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ConstrainedExecution;
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


namespace generalMethods
{

    public class generalMethodsClass
    {
        

        public static void FindLayer()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            Database docDb = doc.Database;
            var docScale = docDb.Cannoscale.DrawingUnits;
            using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
            {

                var lyrTbl = transaction.GetObject(docDb.LayerTableId, OpenMode.ForWrite) as LayerTable;

                doc.Editor.Command("_.LAYON");
                doc.Editor.Command("_.LAYTHW");

                var options = new PromptNestedEntityOptions("\nSelect entity on layer to turn on");
                var pointResult = doc.Editor.GetNestedEntity(options);
                if (pointResult.Status != PromptStatus.OK)
                {
                    return;
                }
                var ent = transaction.GetObject(pointResult.ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                doc.Editor.Command("_.LAYERP");
                doc.TransactionManager.EnableGraphicsFlush(true);
                doc.TransactionManager.FlushGraphics();
                doc.Editor.Regen();
                var layer = lyrTbl[ent.Layer];
                var acLyrTblRec = transaction.GetObject(lyrTbl[ent.Layer], OpenMode.ForWrite) as LayerTableRecord;
                acLyrTblRec.IsFrozen = false;
                acLyrTblRec.IsOff = false;
                //doc.SendStringToExecute("REGEN", true, false, false);
                //doc.Editor.Command("-LAYER T "+ "Topo with some layers frozen|INTER_CONTOUR ");
                //doc.SendStringToExecute("-LAYER T "+ent.Layer+"  ", true, false, false);



                transaction.Commit();
            }
            doc.Editor.Regen();
        }
        public static void MTextUp()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database docDb = doc.Database;
            var snapang = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("SNAPANG");
            var docScale = docDb.Cannoscale.DrawingUnits;

            using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
            {

                BlockTable acBlkTbl = transaction.GetObject(docDb.BlockTableId,
                                                OpenMode.ForRead) as BlockTable;
                BlockTableRecord acBlkTblRec = transaction.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;

                var p1 = doc.Editor.GetPoint("first corner");
                using (MText acMText = new MText())
                {
                    acMText.Location = p1.Value;
                    acMText.Contents = "";
                    acMText.TextHeight = docScale *.1;
                    acMText.Rotation = (double)snapang;
                    acBlkTblRec.AppendEntity(acMText);
                    transaction.AddNewlyCreatedDBObject(acMText, true);
                }
                doc.Editor.Regen();
                var lastSelected = doc.Editor.SelectLast();
                if (lastSelected.Status == PromptStatus.OK)
                {
                    var lastID = lastSelected.Value.GetObjectIds()[0];
                    MText mt = transaction.GetObject(lastID, OpenMode.ForWrite) as MText;

                    if (mt == null)
                        return;
                    InplaceTextEditor.Invoke(mt, new InplaceTextEditorSettings());


                    transaction.Commit();
                };
                 
                
            }

            doc.Editor.Regen();
        }

        public static void MLeaderUp()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
            {
                doc.Editor.Command("UCS", "View");
                doc.Editor.Command("MLEADER");
                doc.Editor.Command("UCS", "World");

                transaction.Commit();
            }

            doc.Editor.Regen();
        }
    }
}
