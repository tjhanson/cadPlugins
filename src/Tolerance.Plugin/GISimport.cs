using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

namespace GisImport
{

    public class GisDataImport
    {
        

        public static void gisCSVtoBlocks()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            Database docDb = doc.Database;
            var docScale = docDb.Cannoscale.DrawingUnits;
            using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
            {
                var blockNameOptions = new PromptStringOptions("\nEnter Block Name");

                var bName = doc.Editor.GetString(blockNameOptions);
                if (bName.Status != PromptStatus.OK)
                {
                    return;
                }

                var layerOptions = new PromptStringOptions("\nEnter Layer Name");

                var insertLayer = doc.Editor.GetString(layerOptions);
                if (insertLayer.Status != PromptStatus.OK)
                {
                    return;
                }

                PromptOpenFileOptions pofo = new PromptOpenFileOptions("\nEnter File");
                pofo.PreferCommandLine = false;
                pofo.DialogName = "Select File";
                pofo.DialogCaption = "Select File";

                PromptFileNameResult pfnr = Application.DocumentManager.MdiActiveDocument.Editor.GetFileNameForOpen(pofo);
                if (pfnr.Status != PromptStatus.OK) return;

                string FileName = pfnr.StringResult;
                FileStream Fstream = new FileStream(FileName, FileMode.Open);
                StreamReader read = new StreamReader(Fstream);
                string st = read.ReadLine();
                st = read.ReadLine();


                BlockTable acBlkTbl;
                acBlkTbl = transaction.GetObject(docDb.BlockTableId,
                                                OpenMode.ForRead) as BlockTable;

                BlockTableRecord blockDef = acBlkTbl[bName.StringResult].GetObject(OpenMode.ForRead) as BlockTableRecord;
                if (blockDef == null)
                {
                    doc.Editor.WriteMessage("Block does not exist");
                    return;
                }
                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = transaction.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;
                var stLen = st.Split(',').Length;
                while (st != null)
                {
                    var line = st.Split(',');
                    var b = new BlockReference(new Point3d(double.Parse(line[stLen-2]), double.Parse(line[stLen - 1]), 0), blockDef.ObjectId);
                    b.ScaleFactors = new Scale3d(docScale);
                    b.BlockUnit = UnitsValue.Undefined;
                    b.Layer = insertLayer.StringResult;
                    acBlkTblRec.AppendEntity(b);
                    transaction.AddNewlyCreatedDBObject(b, true);

                    st = read.ReadLine();
                }
                Fstream.Dispose();

                doc.TransactionManager.EnableGraphicsFlush(true);
                doc.TransactionManager.FlushGraphics();
                doc.Editor.Regen();
                transaction.Commit();
            }
        }
    }
}
