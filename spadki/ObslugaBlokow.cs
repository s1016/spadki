using System.IO;
using ZwSoft.ZwCAD.ApplicationServices;
using ZwSoft.ZwCAD.DatabaseServices;
using ZwSoft.ZwCAD.Geometry;
using ZwSoft.ZwCAD.EditorInput;
using Zwa = ZwSoft.ZwCAD.ApplicationServices.Application;

namespace tabela.dodatkowe
{

    public class ObslugaBlokow

    {
        Document doc;
        Database db;
        Editor ed;

        public ObslugaBlokow()
        {
            ActiveDoc = Zwa.DocumentManager.MdiActiveDocument;
        }
        //ActiveDoc = Zwa.DocumentManager.MdiActiveDocument;    
        Document ActiveDoc
        {
            get { return doc; }
            set
            {
                doc = value;
                if (doc == null)
                {
                    db = null;
                    ed = null;
                }
                else
                {
                    db = doc.Database;
                    ed = doc.Editor;
                }
            }
        }
        //=======================================================================
        public const double kPi = 3.14159265358979323846;
        //=======================================================================
        #region metoda RadiansToDegrees
        public static double RadiansToDegrees(
            double radians)
        {
            return (radians * (kPi / 180.0));
        }
        #endregion

        /// <summary>
        ///  CreateLayer layer if not existing 
        /// </summary>
        /// <param name="db"></param>
        /// <param name="layerName"></param>
        /// <param name="color"></param>
        /// <returns> ObjectId of existing or new layer </returns>
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
                    ltr.Color = ZwSoft.ZwCAD.Colors.Color.FromColorIndex(
                        ZwSoft.ZwCAD.Colors.ColorMethod.ByAci, color);
                    layerId = lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                }
                tr.Commit();
            }
            return layerId;
        }
        //=======================================================================

            /// <summary>
            /// Wyciąganie atrubutu z bloku
            /// </summary>
            /// <param name="tr"></param>
            /// <param name="blkRef"></param>
            /// <param name="attribute"></param>
            /// <returns></returns>
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


        /// <summary>
        ///  InsertBlockReference
        /// </summary>
        /// <param name="db"></param>
        /// <param name="blockName"></param>
        /// <param name="insertPoint"> In Current UCS</param>
        /// <param name="scale"> </param>
        /// <param name="angle"> </param>
        /// <param name="layer"> </param>
        /// <param name="attValues"></param>
        /// <returns>ObjectId of BlockReference </returns>
        public ObjectId InsertBlockReference(
            Database db,
            string blockName,
            string blockPath,
            Point3d insertPoint,
            Scale3d scale,
            double angle,
            string layer,
            System.Collections.Hashtable attValues)
        {
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                if (!bt.Has(blockName))
                {
                    //string blockPath = HostApplicationServices.Current.FindFile(blockName + ".dwg", db, FindFileHint.Default);
                    //blockPath = "D:\\OneDrive\\Visual Studio 2017\\repo\\tabela\\tabela\\bin\\Debug\\bloki\\netloadBPBW.dwg";
                    //blockPath = "d:\\BPBW.dwg";
                    if (string.IsNullOrEmpty(blockPath))
                    {
                        return ObjectId.Null;
                    }
                    bt.UpgradeOpen();
                    using (Database tmpDb = new Database(false, true))
                    {
                        tmpDb.ReadDwgFile(blockPath, FileShare.Read, true, null);
                        db.Insert(blockName, tmpDb, true);
                    }
                }
                BlockTableRecord btr = tr.GetObject(bt[blockName], OpenMode.ForRead) as BlockTableRecord;
                BlockReference br = new BlockReference(insertPoint, bt[blockName])
                {
                    ScaleFactors = scale,
                    Rotation = angle,
                    Layer = layer
                };
                br.TransformBy(ed.CurrentUserCoordinateSystem);

                br.RecordGraphicsModified(true);
                (tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord).AppendEntity(br);
                foreach (ObjectId id in btr)
                {
                    AttributeDefinition aDef = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
                    if (aDef != null)
                    {
                        AttributeReference aRef = new AttributeReference();
                        aRef.SetAttributeFromBlock(aDef, br.BlockTransform);
                        aRef.Position = aDef.Position + br.Position.GetAsVector();
                        if (attValues.ContainsKey(aDef.Tag.ToUpper()))
                        {
                            aRef.TextString = attValues[aDef.Tag.ToUpper()].ToString();
                        }
                        br.AttributeCollection.AppendAttribute(aRef);
                        tr.AddNewlyCreatedDBObject(aRef, true);
                    }
                }
                tr.AddNewlyCreatedDBObject(br, true);
                tr.Commit();
                return br.ObjectId;
            }
        }
        //=======================================================================

        /// <summary>
        /// BlockInsert
        /// </summary>
        /// <param name="blokPath">ścieżka do bloku</param>
        /// <param name="blockName">nazwa bloku</param>
        /// <param name="ht">kolekcja z atrybutami</param>
        /// <param name="layerName">warstwa na która wstawiamy blok</param>
        /// <param name="layerColor">kolor warstwy</param>
        public void BlockInsert
            (
            string blokPath,
            string blockName,
            double scale,
            System.Collections.Hashtable ht,
            string layerName,
            short layerColor
            )
        {
            //blockName = "tab1";

            //scale = 2.5;
            double rotation = RadiansToDegrees(0.0);
            CreateLayer(db, layerName, layerColor);

            PromptPointResult ppr = ed.GetPoint("\nWskaż punkt wstawiania tabeli: ");
            if (ppr.Status != PromptStatus.OK)
                return;

            ObjectId blockRefID = InsertBlockReference(
                                    db,
                                    blockName,
                                    blokPath,
                                    ppr.Value,   // point in current UCS
                                    new Scale3d(scale, scale, 1.0),
                                    rotation,
                                    layerName,
                                    ht
            );
        }
        //=======================================================================

        /// <summary>
        /// FreezeLayer
        /// </summary>
        public void FreezeLayer()
        {
            // Get the current document and database
            Document zwDoc = Application.DocumentManager.MdiActiveDocument;
            Database zwCurDb = zwDoc.Database;

            // Start a transaction
            using (Transaction zwTrans = zwCurDb.TransactionManager.StartTransaction())
            {
                // Open the Layer table for read
                LayerTable zwLyrTbl;
                zwLyrTbl = zwTrans.GetObject(zwCurDb.LayerTableId,
                                                OpenMode.ForRead) as LayerTable;

                string sLayerName = "!DR_Tabelka_faksymilki";

                if (zwLyrTbl.Has(sLayerName) == false)
                {
                    using (LayerTableRecord zwLyrTblRec = new LayerTableRecord())
                    {
                        // Assign the layer a name
                        zwLyrTblRec.Name = sLayerName;

                        // Upgrade the Layer table for write
                        zwLyrTbl.UpgradeOpen();

                        // Append the new layer to the Layer table and the transaction
                        zwLyrTbl.Add(zwLyrTblRec);
                        zwTrans.AddNewlyCreatedDBObject(zwLyrTblRec, true);

                        // Freeze the layer
                        zwLyrTblRec.IsFrozen = true;
                    }
                }
                else
                {
                    LayerTableRecord acLyrTblRec = zwTrans.GetObject(zwLyrTbl[sLayerName],
                                                    OpenMode.ForWrite) as LayerTableRecord;

                    // Freeze the layer
                    acLyrTblRec.IsFrozen = true;
                }

                // Save the changes and dispose of the transaction
                zwTrans.Commit();
            }
        }
    }
}
