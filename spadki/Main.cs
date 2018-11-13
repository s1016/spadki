using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using ZwSoft.ZwCAD;
using ZwSoft.ZwCAD.Runtime;
using ZwSoft.ZwCAD.ApplicationServices;
using ZwSoft.ZwCAD.DatabaseServices;
using ZwSoft.ZwCAD.EditorInput;
using ZwSoft.ZwCAD.Geometry;
using System.IO;

namespace spadki
{
    public class KlasaSpadki
    {
        CultureInfo ci = new CultureInfo("pl-PL");
        Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
        Database db = HostApplicationServices.WorkingDatabase;
        Document doc = Application.DocumentManager.MdiActiveDocument;

        [CommandMethod("kkk")]
        public void KotaKotaKota()
       
        {

            ObjectId blockRefID = SelectAnyBlockWithAttributes("\nWybierz kotę 1: ");
            if (blockRefID == ObjectId.Null)
            {
                return;
            }
            string attr1 = GetAttributeOId(blockRefID, "poziom");
            Point3d rz1P = GetReferenceInsertPoint(blockRefID);

            ObjectId blockRefID2 = SelectAnyBlockWithAttributes("\nWybierz kotę 2: ");
            if (blockRefID == ObjectId.Null)
            {
                return;
            }
            string attr2 = GetAttributeOId(blockRefID2, "poziom");
            Point3d rz2P = GetReferenceInsertPoint(blockRefID2);

            double dist = rz2P.DistanceTo(rz1P);
            Vector3d ang = rz2P.GetVectorTo(rz1P);
            double spadek;
            double rz1 = double.Parse(attr1);
            double rz2 = double.Parse(attr2);
            Point3d punktwstawiania = myMidPoint(rz1P, rz2P);
            double rzN;
            if (rz1 >= rz2)
            {
                spadek = (rz1 - rz2) / dist;
            }
            else
            {
                spadek = (rz2 - rz1) / dist;
            }

            Transaction tr = db.TransactionManager.StartTransaction();
            try
            {
                BlockTable acBlkTbl;
                acBlkTbl = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                string warstwa = ((BlockReference)tr.GetObject(blockRefID2, OpenMode.ForRead)).Layer;
                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = tr.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                Polyline acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(rz1P.X, rz1P.Y), 0, 0, 0);
                acPoly.AddVertexAt(1, new Point2d(rz2P.X, rz2P.Y), 0, 0, 0);
                acPoly.Color = ZwSoft.ZwCAD.Colors.Color.FromRgb(221, 33, 215);

                acBlkTblRec.AppendEntity(acPoly);
                tr.AddNewlyCreatedDBObject(acPoly, true);
                // Update the display and display an alert message
                doc.Editor.Regen();

               
                PromptPointResult pPtRes;
                PromptPointOptions pPtOpts = new PromptPointOptions("");
                pPtOpts.Message = "\nWskaż miejsce wstawienia: ";
                pPtRes = ed.GetPoint(pPtOpts);
                Point3d rzNowaP = pPtRes.Value;
                punktwstawiania = rzNowaP;
                double distNew = rzNowaP.DistanceTo(rz1P);

                if (rz1 > rz2)
                    rzN = rz1 - (distNew * spadek);
                else if (rz1 < rz2)
                    rzN = rz1 + (distNew * spadek);
                else
                    rzN = rz1;
                ed.WriteMessage("\nrzędna= " + rzN.ToString("N2"));
                string attText = rzN.ToString("N2");
                wstawblok(db, ed, punktwstawiania, 0, attText, 0, "Kota2", "Poziom", warstwa);
                acPoly.Erase(true);
                tr.Commit();
            }

            catch (ZwSoft.ZwCAD.Runtime.Exception ex)
            {
                ed.WriteMessage(("Exception: " + ex.Message));
            }
            finally
            {
                tr.Dispose();
            }
        }

        [CommandMethod("kks")]
        public void KotaKotaSpadek()
        {
            string warstwa = "!DR_spadki";
            string fileName = "SPL.dwg";

            ObjectId blockRefID = SelectAnyBlockWithAttributes("\nWybierz kotę 1: ");
            if (blockRefID == ObjectId.Null)
            {
                return;
            }
            string attr1 = GetAttributeOId(blockRefID, "poziom");
            Point3d rz1P = GetReferenceInsertPoint(blockRefID);


            ObjectId blockRefID2 = SelectAnyBlockWithAttributes("\nWybierz kotę 2: ");
            if (blockRefID == ObjectId.Null)
            {
                return;
            }
            string attr2 = GetAttributeOId(blockRefID2, "poziom");
            Point3d rz2P = GetReferenceInsertPoint(blockRefID2);

            double dist = rz2P.DistanceTo(rz1P);
            Vector3d ang = rz2P.GetVectorTo(rz1P);
            double spadek;
            double rz1 = double.Parse(attr1);
            double rz2 = double.Parse(attr2);
            Point3d punktwstawiania = myMidPoint(rz1P, rz2P);
            double angle, angle2, atangle;
            if (rz1 >= rz2)
            {
                spadek = (rz1 - rz2) / dist;
                angle = Angle(rz2P, rz1P);
            }
            else
            {
                spadek = (rz2 - rz1) / dist;
                angle = Angle(rz1P, rz2P);
            }

            angle2 = (angle * (180.0 / Math.PI));

            if (angle2 >= 90 && angle2 <= 270)
                atangle = RadiansToDegrees(angle2 + 180);
            else
                atangle = RadiansToDegrees(angle2);

            Transaction tr = db.TransactionManager.StartTransaction();
            try
            {
                BlockTable acBlkTbl;
                acBlkTbl = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = tr.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                string blockPath = Path.Combine(Path.GetDirectoryName(this.GetType().Assembly.Location), @"bloki\", fileName);
                ImportBlocks(blockPath);
                CreateLayer(db, "!DR_spadki", 0);
                ed.WriteMessage("\nSpadek " + spadek.ToString("P2", ci));
                ed.WriteMessage("\nOdległość " + dist.ToString("N2", ci));
                
                string attText = spadek.ToString("P2");
                // ed.WriteMessage("pnkt" + punktwstawiania + " angle2 " + angle2 + " atttext " + attText + " atangle " + atangle);
                
                //ed.WriteMessage(Path.Combine(Path.GetDirectoryName(this.GetType().Assembly.Location)));
                wstawblok(db, ed, punktwstawiania, angle2, attText, atangle, "SPL", "SP1", warstwa);

                tr.Commit();
            }

            catch (ZwSoft.ZwCAD.Runtime.Exception ex)
            {
                ed.WriteMessage(("Exception: " + ex.Message));
            }
            finally
            {
                tr.Dispose();
            }
        }

        [CommandMethod("ksk")]
        public void KotaSpadekKota()
        {

            ObjectId blockRefID1 = SelectAnyBlockWithAttributes("\nWybierz kotę 1: ");
            if (blockRefID1 == ObjectId.Null)
            {
                return;
            }
            string attr1 = GetAttributeOId(blockRefID1, "poziom");
            Point3d rz1P = GetReferenceInsertPoint(blockRefID1);

            PromptPointResult pPtRes;
            PromptPointOptions pPtOpts = new PromptPointOptions("");
            pPtOpts.Message = "\nWskaż miejsce wstawienia: ";
            pPtRes = ed.GetPoint(pPtOpts);
            Point3d punktwstawiania = pPtRes.Value;
            double dist = punktwstawiania.DistanceTo(rz1P);


            Vector3d ang = punktwstawiania.GetVectorTo(rz1P);


            double spadek = 0.01;
            PromptDoubleOptions pDblOpts = new PromptDoubleOptions("\nPodaj spadek: ");
            pDblOpts.AllowNone = false;
            
            PromptDoubleResult pDblRes = ed.GetDouble(pDblOpts);
            spadek = (pDblRes.Value/100);

            double rz1 = double.Parse(attr1);
            double rz2 = rz1 + (dist * spadek);

                     
            Transaction tr = db.TransactionManager.StartTransaction();
            try
            {
                BlockTable acBlkTbl;
                acBlkTbl = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                string warstwa = ((BlockReference)tr.GetObject(blockRefID1, OpenMode.ForRead)).Layer;
                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = tr.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
               // Polyline acPoly = new Polyline();
                //acPoly.AddVertexAt(0, new Point2d(rz1P.X, rz1P.Y), 0, 0, 0);
                //acPoly.AddVertexAt(1, new Point2d(punktwstawiania.X, punktwstawiania.Y), 0, 0, 0);
                //acPoly.Color = ZwSoft.ZwCAD.Colors.Color.FromRgb(221, 33, 215);

                //acBlkTblRec.AppendEntity(acPoly);
                //tr.AddNewlyCreatedDBObject(acPoly, true);
                // Update the display and display an alert message
                //doc.Editor.Regen();

                ed.WriteMessage("\nrzędna= " + rz2.ToString("N2"));
                string attText = rz2.ToString("N2");
                wstawblok(db, ed, punktwstawiania, 0, attText, 0, "Kota2", "Poziom", warstwa);
                //acPoly.Erase(true);
                tr.Commit();
            }

            catch (ZwSoft.ZwCAD.Runtime.Exception ex)
            {
                ed.WriteMessage(("Exception: " + ex.Message));
            }
            finally
            {
                tr.Dispose();
            }
        }

                
        private ObjectId SelectAnyBlockWithAttributes(string prompt)
        {
            
            using (ed.Document.Database.TransactionManager.StartTransaction())
            {
                var peo = new PromptEntityOptions(prompt);
                peo.SetRejectMessage("\nTylko bloki.");
                peo.AddAllowedClass(typeof(BlockReference), true);
                while (true)
                {
                    PromptEntityResult per = ed.GetEntity(peo);
                    if (per.Status != PromptStatus.OK)
                    {
                        return ObjectId.Null;
                    }

                    var attCol = (per.ObjectId.GetObject(OpenMode.ForRead) as BlockReference).AttributeCollection;
                    if (attCol != null && attCol.Count > 0)
                    {

                        return per.ObjectId;
                    }

                    ed.WriteMessage("\nZły wybór, \nTylko bloki z atrybutem.");
                }
            }
        }

        //******************* wstawianie warstwy jesli brak
        public static ObjectId CreateLayer(Database db, string layerName, short color)
        {
            ObjectId layerId = ObjectId.Null;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForWrite);
                //Check if EmployeeLayer exists...
                if (lt.Has(layerName))
                {
                    layerId = lt[layerName];
                }
                else
                {
                    //if not, create the layer here.
                    LayerTableRecord ltr = new LayerTableRecord();
                    ltr.Name = layerName;
                    ltr.Color = ZwSoft.ZwCAD.Colors.Color.FromColorIndex(ZwSoft.ZwCAD.Colors.ColorMethod.ByAci, color);
                    layerId = lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                }
                tr.Commit();
            }
            return layerId;
        }
        //******************* importowanie def bloku do rysunku
        public void ImportBlocks(
                string sourceFileName
                )
        {
            DocumentCollection dm = Application.DocumentManager;
            Editor ed = dm.MdiActiveDocument.Editor;
            Database destDb = dm.MdiActiveDocument.Database;
            Database sourceDb = new Database(false, true);
            //PromptResult sourceFileName;
            try
            {

                sourceDb.ReadDwgFile(sourceFileName, System.IO.FileShare.Read, true, "");

                // Create a variable to store the list of block identifiers
                ObjectIdCollection blockIds = new ObjectIdCollection();

                ZwSoft.ZwCAD.DatabaseServices.TransactionManager tm = sourceDb.TransactionManager;

                using (Transaction myT = tm.StartTransaction())
                {
                    // Open the block table
                    BlockTable bt = (BlockTable)tm.GetObject(sourceDb.BlockTableId, OpenMode.ForRead, false);

                    // Check each block in the block table
                    foreach (ObjectId btrId in bt)
                    {
                        BlockTableRecord btr = (BlockTableRecord)tm.GetObject(btrId, OpenMode.ForRead, false);
                        // Only add named & non-layout blocks to the copy list
                        if (!btr.IsAnonymous && !btr.IsLayout)
                            blockIds.Add(btrId);
                        btr.Dispose();
                    }
                }
                // Copy blocks from source to destination database
                IdMapping mapping = new IdMapping();
                sourceDb.WblockCloneObjects(blockIds, destDb.BlockTableId, mapping, DuplicateRecordCloning.Replace, false);
                ed.WriteMessage("\nCopied "
                                + blockIds.Count.ToString()
                                + " block definitions from "
                                + sourceFileName
                                + " to the current drawing.");
            }
            catch (ZwSoft.ZwCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("\nError during copy: " + ex.Message);
            }
            sourceDb.Dispose();
        }


        //******************* wstawianie bloku 
        public void wstawblok(
            Database db,
            Editor ed,
            Point3d pnkwstaw,
            double kat,
            string attText,
            double atangle,
            string blockName,
            string attTag,
            string warstwa
            )
        {

            using (Transaction tr1 = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr1.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                using (BlockReference br = new BlockReference(Point3d.Origin, bt[blockName]))
                {
                    br.Rotation = RadiansToDegrees(kat);
                    br.Layer = warstwa;
                    br.TransformBy(Matrix3d
                        .Displacement(pnkwstaw - Point3d.Origin)
                        .PreMultiplyBy(ed.CurrentUserCoordinateSystem));

                    btr.AppendEntity(br);
                    tr1.AddNewlyCreatedDBObject(br, true);
                    InsertAttibuteInBlockRef(br, attTag, attText, atangle, tr1);
                }
                tr1.Commit();
            }
        }

        //******************* zamiana radianów na stopnie
        public static double RadiansToDegrees(
           double radians)
        {
            return (radians * (Math.PI / 180.0));
        }
        //******************* wstaw atrybut do bloku

        public static void InsertAttibuteInBlockRef(
           BlockReference blkRef,
           string attributeTag,
           string attributeText,
           double atangle,
           Transaction tr)
        {
            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead);
            foreach (ObjectId id in btr)
            {
                AttributeDefinition aDef = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
                if (aDef != null)
                {
                    AttributeReference aRef = new AttributeReference();
                    aRef.SetAttributeFromBlock(aDef, blkRef.BlockTransform);
                    if (aRef.Tag == attributeTag)
                        aRef.TextString = attributeText;
                    aRef.Rotation = atangle;
                    blkRef.AttributeCollection.AppendAttribute(aRef);
                    tr.AddNewlyCreatedDBObject(aRef, true);
                }
            }
        }


        // znalezienie punktu środkowego pomiedzy dwoa punktami
        public static Point3d myMidPoint(Point3d firPnt, Point3d secPnt)
        {
            LineSegment3d line3d = new LineSegment3d(firPnt, secPnt);
            Point3d MIDpoint = line3d.MidPoint;
            return MIDpoint;
        }


        // wyciaganie kąta pomiędzy dwoma punktami
        public static double Angle(Point3d firPnt, Point3d secPnt)
        {
            double lineAngle;
            if (firPnt.X == secPnt.X)
            {
                if (firPnt.Y < secPnt.Y)
                    lineAngle = Math.PI / 2.0;
                else
                    lineAngle = (Math.PI / 2.0) * 3.0;
            }
            else
            {
                lineAngle = Math.Atan((secPnt.Y - firPnt.Y) / (secPnt.X - firPnt.X));
                if (firPnt.X > secPnt.X)
                    lineAngle = Math.PI + lineAngle;
                else
                    if (lineAngle < 0.0)
                    lineAngle = (Math.PI * 2.0) + lineAngle;
            }
            return (lineAngle);
        }

        // wyciaganie konkretnego atrybutu z bloku
        public static string GetAttribute(Transaction tr, BlockReference blkRef, string attribute)
        {
            AttributeCollection attCol = blkRef.AttributeCollection;

            foreach (ObjectId attId in attCol)
            {
                AttributeReference attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForWrite);

                if (attRef.Tag.ToUpper() == attribute.ToUpper())
                {
                    return attRef.TextString;
                }
            }

            return string.Empty;
        }

        // wyciaganie konkretnego atrybutu z bloku
        public static string GetAttributeOId(ObjectId blockRefID, string attribute)
        {
            string poziom = null;
            using (Transaction tr = blockRefID.Database.TransactionManager.StartTransaction())
            {
                var blkRef = blockRefID.GetObject(OpenMode.ForRead) as BlockReference;
                AttributeCollection attCol = blkRef.AttributeCollection;

                foreach (ObjectId attId in attCol)
                {
                    AttributeReference attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);

                    if (attRef.Tag.ToUpper() == attribute.ToUpper())
                    {
                        poziom = attRef.TextString;
                        poziom = poziom.Replace('.', ',');
                    }
                }
                tr.Commit();
            }
            return poziom;
        }

        public Point3d GetReferenceInsertPoint(ObjectId blockRefID)
        {
            Point3d inserPoint = new Point3d();

            Transaction tr = db.TransactionManager.StartTransaction();
            try
            {
                inserPoint = ((BlockReference)tr.GetObject(blockRefID, OpenMode.ForRead)).Position;
            }

            catch (ZwSoft.ZwCAD.Runtime.Exception ex)
            {
                ed.WriteMessage(("Exception: " + ex.Message));
            }
            finally
            {
                tr.Dispose();
            }

            return inserPoint;
        }
    }

}