using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

namespace PointLoadCalc
{
    public class Class2
    {
        class AttInfo //No Idea what this does yet.
        {
            private Point3d _pos;
            private Point3d _aln;
            private bool _aligned;

            public AttInfo(Point3d pos, Point3d aln, bool aligned)
            {
                _pos = pos;
                _aln = aln;
                _aligned = aligned;
            }

            public Point3d Position
            {
                set { _pos = value; }
                get { return _pos; }
            }

            public Point3d Alignment
            {
                set { _aln = value; }
                get { return _aln; }
            }

            public bool IsAligned
            {
                set { _aligned = value; }
                get { return _aligned; }
            }
        }

        class BlockJig : EntityJig //Jig the block
        {
            private Point3d _pos;
            private Dictionary<ObjectId, AttInfo> _attInfo;
            private Transaction _tr;

            public BlockJig(
              Transaction tr,
              BlockReference br,
              Dictionary<ObjectId, AttInfo> attInfo
            )
                : base(br)
            {
                _pos = br.Position;
                _attInfo = attInfo;
                _tr = tr;
            }

            protected override bool Update()
            {
                BlockReference br = Entity as BlockReference;

                br.Position = _pos;

                if (br.AttributeCollection.Count != 0)
                {
                    foreach (ObjectId id in br.AttributeCollection)
                    {
                        DBObject obj =
                          _tr.GetObject(id, OpenMode.ForRead);
                        AttributeReference ar =
                          obj as AttributeReference;

                        // Apply block transform to att def position

                        if (ar != null)
                        {
                            ar.UpgradeOpen();
                            AttInfo ai = _attInfo[ar.ObjectId];
                            ar.Position =
                              ai.Position.TransformBy(br.BlockTransform);
                            if (ai.IsAligned)
                            {
                                ar.AlignmentPoint =
                                  ai.Alignment.TransformBy(br.BlockTransform);
                            }
                            if (ar.IsMTextAttribute)
                            {
                                ar.UpdateMTextAttribute();
                            }
                        }
                    }
                }
                return true;
            }

            protected override SamplerStatus Sampler(JigPrompts prompts)
            {
                JigPromptPointOptions opts =
                  new JigPromptPointOptions("\nSelect insertion point:");
                opts.BasePoint = new Point3d(0, 0, 0);
                opts.UserInputControls =
                  UserInputControls.NoZeroResponseAccepted;

                PromptPointResult ppr = prompts.AcquirePoint(opts);

                if (_pos == ppr.Value)
                {
                    return SamplerStatus.NoChange;
                }

                _pos = ppr.Value;

                return SamplerStatus.OK;
            }

            public PromptStatus Run()
            {
                Document doc =
                  Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;

                PromptResult promptResult = ed.Drag(this);
                return promptResult.Status;
            }
        }

        public class Commands
        {
            //[CommandMethod("IB")]
            public static void ImportBlocks()
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                Database destDb = doc.Database;
                Database sourceDb = new Database(false, true);

                String sourceFileName;

                try
                {
                    //Get name of DWG from which to copy blocks
                    sourceFileName = "C:\\Program Files\\AutoCAD MEP 2010\\DynaPrograms\\TagHgr.dwg";

                    //Read the DWG file into the database
                    sourceDb.ReadDwgFile(sourceFileName.ToString(), System.IO.FileShare.Read, true, "");

                    //Create a variable to store the list of block Identifiers
                    ObjectIdCollection blockIds = new ObjectIdCollection();

                    Autodesk.AutoCAD.DatabaseServices.TransactionManager tm = sourceDb.TransactionManager;

                    using (Transaction myTm = tm.StartTransaction())
                    {
                        //Open the block table
                        BlockTable bt = (BlockTable)tm.GetObject(sourceDb.BlockTableId, OpenMode.ForRead, false);

                        //Check each block in the table
                        foreach (ObjectId btrId in bt)
                        {
                            BlockTableRecord btr = (BlockTableRecord)tm.GetObject(btrId, OpenMode.ForRead, false);

                            //Only add named and non-layout blocks to the copy file if the don't already exist.
                            if (!btr.IsAnonymous && !btr.IsLayout)
                                blockIds.Add(btrId);
                            btr.Dispose();
                        }
                    }
                    //Copy blocks from source to dest database
                    IdMapping mapping = new IdMapping();
                    sourceDb.WblockCloneObjects(blockIds, destDb.BlockTableId, mapping, DuplicateRecordCloning.Replace, false);
                    //Writes the number of blocks copied to cmd line.
                    //ed.WriteMessage("\nCopied: " + blockIds.Count.ToString() + " block definitions from " + sourceFileName + " to the current drawing.");

                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    ed.WriteMessage(ex.Message);
                }
                sourceDb.Dispose();

            } //end ImportBlocks()

            //[CommandMethod("BJ")]
            static public void BlockJigCmd(Double finalCalcText)
            {
                Document doc =
                  Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

                //Set the block Name to use.
                string pr = "TagHgr";

                Transaction tr =
                  doc.TransactionManager.StartTransaction();
                using (tr)
                {
                    BlockTable bt =
                      (BlockTable)tr.GetObject(
                        db.BlockTableId,
                        OpenMode.ForRead
                      );

                    if (!bt.Has(pr))
                    {
                        ed.WriteMessage(
                          "\nBlock \"" + pr + "\" not found.");
                        return;
                    }

                    BlockTableRecord space =
                      (BlockTableRecord)tr.GetObject(
                        db.CurrentSpaceId,
                        OpenMode.ForWrite
                      );

                    BlockTableRecord btr =
                      (BlockTableRecord)tr.GetObject(
                        bt[pr],
                        OpenMode.ForRead);

                    // Block needs to be inserted to current space before
                    // being able to append attribute to it

                    BlockReference br =
                      new BlockReference(new Point3d(), btr.ObjectId);
                    space.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);

                    Dictionary<ObjectId, AttInfo> attInfo =
                      new Dictionary<ObjectId, AttInfo>();

                    if (btr.HasAttributeDefinitions)
                    {
                        foreach (ObjectId id in btr)
                        {
                            DBObject obj =
                              tr.GetObject(id, OpenMode.ForRead);
                            AttributeDefinition ad =
                              obj as AttributeDefinition;

                            if (ad != null && !ad.Constant)
                            {
                                AttributeReference ar =
                                  new AttributeReference();

                                ar.SetAttributeFromBlock(ad, br.BlockTransform);
                                ar.Position =
                                  ad.Position.TransformBy(br.BlockTransform);

                                if (ad.Justify != AttachmentPoint.BaseLeft)
                                {
                                    ar.AlignmentPoint =
                                      ad.AlignmentPoint.TransformBy(br.BlockTransform);
                                }
                                if (ar.IsMTextAttribute)
                                {
                                    ar.UpdateMTextAttribute();
                                }

                                //ar.TextString = ad.TextString;
                                ar.TextString = PopulateAttributes(ad.Tag, finalCalcText.ToString());

                                ObjectId arId =
                                  br.AttributeCollection.AppendAttribute(ar);
                                tr.AddNewlyCreatedDBObject(ar, true);

                                // Initialize our dictionary with the ObjectId of
                                // the attribute reference + attribute definition info

                                attInfo.Add(
                                  arId,
                                  new AttInfo(
                                    ad.Position,
                                    ad.AlignmentPoint,
                                    ad.Justify != AttachmentPoint.BaseLeft
                                  )
                                );
                            }
                        }
                    }
                    // Run the jig

                    BlockJig myJig = new BlockJig(tr, br, attInfo);

                    if (myJig.Run() != PromptStatus.OK)
                        return;

                    // Commit changes if user accepted, otherwise discard

                    tr.Commit();
                }
            }

            public static string PopulateAttributes(string attbName, string finalCalcText)
            {
                string attbValue = "";

                if (attbName == "TRADE-SYMBOL")
                {
                    attbValue = "EC";
                }

                if (attbName == "SEISMIC-LOAD")
                {
                    attbValue = finalCalcText; //Pouplates the seismic load with the calculated value.
                }

                return attbValue;
            }
        }
    }
}
