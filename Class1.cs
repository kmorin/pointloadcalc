using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;            //for transactions, etc.
using Autodesk.AutoCAD.EditorInput;                 //for prompting, etc.
using Autodesk.AutoCAD.Runtime;                     //for CommandMethod, etc.
using Autodesk.AutoCAD.Geometry;
using Autodesk.Aec.Building.DatabaseServices;       //for Member, etc.
using Autodesk.Aec.Building.ApplicationServices;    //for PartManager, etc.
using Autodesk.Aec.Building.Elec.DatabaseServices;  //for Devices, etc.
using ObjectId = Autodesk.AutoCAD.DatabaseServices.ObjectId;
using ObjectIdCollection = Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection;

namespace PointLoadCalc
{
    public class Class1
    {
        [CommandMethod("PointLoadCalc")] //Main command to use.
        public void PointLoadCalc()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            Double dblDistance = GetDistance(); //Get distance asking for Leftmost, Center and Rightmost racks.
            Double distResult = dblDistance / 2; //Take the distance and divide by 2 to get correct length.

            //write the message to the command line (Can Delete if necessary)
            ed.WriteMessage("\n1/2 the distance of " + dblDistance.ToString() + " is: " + distResult.ToString());

            //Get entities selection
            GetAECObjects(distResult);
        } //end PointLoadCalc()

        // Get distance between two user selected points.
        public Double GetDistance()
        {
            Double dist = 0;

            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            //Prompt for user selection of points to calculate distance.
            PromptPointResult ppr;
            Point2dCollection colPt = new Point2dCollection();
            PromptPointOptions ppo = new PromptPointOptions("");

            //Prompt for first point
            ppo.Message = "\nSpecify mid of Leftmost Rack: ";
            ppr = ed.GetPoint(ppo);
            colPt.Add(new Point2d(ppr.Value.X, ppr.Value.Y));

            //Exit if the user presses ESC or cancels cmd
            if (ppr.Status == PromptStatus.Cancel) return 0;

            int count = 1;

            while (count <= 2)
            {
                //Prompt for next points
                switch (count)
                {
                    case 1:
                        ppo.Message = "\nSpecify mid of Center Rack: ";
                        break;
                    case 2:
                        ppo.Message = "\nSpecify mid of Rightmost Rack: ";
                        break;
                }
                //use the previous point as the base point
                ppo.UseBasePoint = true;
                ppo.BasePoint = ppr.Value;

                ppr = ed.GetPoint(ppo);
                colPt.Add(new Point2d(ppr.Value.X, ppr.Value.Y));

                if (ppr.Status == PromptStatus.Cancel) return 0;

                //Increment
                count = count + 1;
            }

            //Create the polyline
            using (Polyline acPoly = new Polyline())
            {
                acPoly.AddVertexAt(0, colPt[0], 0, 0, 0);
                acPoly.AddVertexAt(1, colPt[1], 0, 0, 0);
                acPoly.AddVertexAt(2, colPt[2], 0, 0, 0);

                //Don't close polyline
                acPoly.Closed = false;

                //Query the length
                dist = acPoly.Length;

            }//Dispose of polyline.

            return dist; //returns the value of the distance.
        } //end GetDistance()

        #region Get Entities (not used)
        // Get Entity info
        [CommandMethod("GetEntitiesObjects", CommandFlags.Modal | CommandFlags.Redraw | CommandFlags.UsePickSet)]
        public void GetEntities()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                PromptSelectionResult psr = ed.GetSelection();

                if (psr.Status == PromptStatus.OK)
                {

                    //System.Type objType;
                    //string propName = "";
                    //object propValue = null;
                    //bool recursive = false;

                    ObjectId[] objIds = psr.Value.GetObjectIds();

                    foreach (ObjectId objId in objIds)
                    {

                        System.Type objType = objId.GetType();

                        if (objType == null)
                        {
                            ed.WriteMessage("\nType not found");
                        }
                        else
                        {

                            //ListProperties(objType);
                        }
                    }
                }
            } //end try
            catch (SystemException ex)
            {
                ed.WriteMessage("Error " + ex);
            }
        } //end void GetEntities()

        private void ListProperties(System.Type objType)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            ed.WriteMessage("\nReadable Properties for: " + objType.Name + " : ");

            PropertyInfo[] propInfos = objType.GetProperties();
            foreach (PropertyInfo propInfo in propInfos)
            {
                if (propInfo.CanRead)
                {
                    ed.WriteMessage("\n " + propInfo.Name + " : " + propInfo.PropertyType);
                }
            }
            ed.WriteMessage(System.Environment.NewLine); // "\n";
        } //end void ListProperties()
        #endregion

        //Get selected AEC Objects (Only works for Conduit as of V1.0)
        public void GetAECObjects(Double distRes)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                PromptSelectionResult psr = ed.SelectImplied(); //Get user selection.

                if (psr.Status == PromptStatus.OK)
                {
                    ed.SetImpliedSelection(new ObjectId[0]);
                }
                else
                {
                    PromptSelectionOptions selOpts = new PromptSelectionOptions();
                    selOpts.MessageForAdding = "\nSelect objects: ";
                    psr = ed.GetSelection(selOpts);
                }

                // Get the selected objects one by one.
                Transaction tr = doc.TransactionManager.StartTransaction();
                using (tr)
                {
                    try
                    {
                        double[] partSizes = new double[50]; // accept 50 selected objects
                        string[] partCndtType = new string[50];
                        double[] pSizeMult = new double[50]; // accept 50 selected objects
                        int total = 0;
                        ObjectId[] objIds = psr.Value.GetObjectIds();

                        //Ask for weight Table type
                        String weightTable = "";
                        PromptKeywordOptions pko = new PromptKeywordOptions("");
                        pko.Message = "\nEnter table to use calculations from ";
                        pko.Keywords.Add("Water");
                        pko.Keywords.Add("Conductor");
                        pko.Keywords.Default = "Water";
                        pko.AllowNone = true;

                        PromptResult pr = ed.GetKeywords(pko);

                        weightTable = pr.StringResult;

                        //Loop through Selected Conduits.
                        foreach (ObjectId objId in objIds)
                        {

                            Entity ent = (Entity)tr.GetObject(objId, OpenMode.ForRead); //use entity as object type to find.

                            Member mem = tr.GetObject(objId, OpenMode.ForRead) as Member; //set as member (needed to process AEC)
                            partSizes[total] = GetPartSize(mem);
                            partCndtType[total] = GetPartCndtType(mem);
                            pSizeMult[total] = CalcWeightPerFoot(partSizes[total].ToString(), partCndtType[total], weightTable); //now call the calculation loop for each size.
                            ed.WriteMessage("\n[" + total.ToString() + "] " + partSizes[total].ToString() + " x " + pSizeMult[total].ToString()); //write out to command line for verification.
                            total += 1; //increase increment for loop.
                        }

                        Double multSum = pSizeMult.Sum(); //sum up the multipliers for use later.
                        ed.WriteMessage("\n\nThe Sum of the multipliers= " + multSum.ToString());

                        /*Finally Calculate the Point Load weight.*/
                        // Convert the distance to feet.
                        distRes = distRes / 12;

                        ed.WriteMessage("\n\nFormula: (" + multSum.ToString() + " * " + distRes.ToString() + ") / 2");

                        // Multiply the sum of the multipliers by the distance, then divide by 2 to get the individual point load.
                        Double finalCalc = (multSum * distRes) / 2;

                        finalCalc = Math.Round(finalCalc, 1); //Round the calculation to 1 decimal point ex. 3.4

                        ed.WriteMessage("\n\nThe point load= " + finalCalc.ToString() + " x2");

                        // Place point load tag on drawing.
                        Class2.Commands.ImportBlocks(); //Imports the dynamic block from external dwg.
                        Class2.Commands.BlockJigCmd(finalCalc); //Jigs the block with the correct values and places in drawing at user point.
                    }
                    catch (SystemException ex)
                    {
                        ed.WriteMessage("Error" + ex.ToString());
                    }
                    tr.Commit();
                }

            }
            catch (SystemException ex)
            {
                ed.WriteMessage("Error " + ex.ToString());
            }
        } //end GetAECObjects()

        private double GetPartSize(Member mem)
        {
            DataRecord dr = PartManager.GetPartData(mem); //Pulls the partdata from the object as a datarecord.
            DataField df = dr.FindByContextAndIndex(Context.CatalogPartSizeName, 0); //Searches through the MEP catalog for size name.
            string partSizeName = df.ValueString; //Sets the name to a string value.

            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            //write out partsizename to command line for debugging.
            ed.WriteMessage(partSizeName.ToString());

            //Removes the trailing text, just leaving the number. (FOR CONDUIT ONLY)
            partSizeName = partSizeName.Remove(5);

            //Converts the string to a number for later math.
            Double psn = Convert.ToDouble(partSizeName);

            /*Now do something with the partsize
             * 
             * 1.00 in EMT SS -> 1.0
             * 0.50 in EMT SS -> 0.5
             */
            return psn;
        } //end GetPartSize()

        private string GetPartCndtType(Member mem)
        {
            DataRecord dr = PartManager.GetPartData(mem);
            DataField df = dr.FindByContextAndIndex(Context.CatalogPartSizeName, 0);
            string partCndtType = df.ValueString;

            if (partCndtType.Contains("GRC"))
            {
                partCndtType = "GRC";
            }
            if (partCndtType.Contains("EMT"))
            {
                partCndtType = "EMT";
            }
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("\nCONDUIT TYPE: " + partCndtType);

            return partCndtType;
        }//end GetPartCndtType()

        private double CalcWeightPerFoot(String pSize, String pCntType, String weightTable)
        {
            double multiplier = 0;

            if (weightTable == "Water")
            {
                /*Water Weight Tables*/
                if (pCntType == "EMT")
                {
                    //EMT Multipliers
                    switch (pSize)
                    {
                        case "0.5":
                            multiplier = 0.45;
                            break;
                        case "0.75":
                            multiplier = 0.72;
                            break;
                        case "1":
                            multiplier = 1.09;
                            break;
                        case "1.25":
                            multiplier = 1.73;
                            break;
                        case "1.5":
                            multiplier = 2.14;
                            break;
                        case "2":
                            multiplier = 3.09;
                            break;
                        case "2.5":
                            multiplier = 4.97;
                            break;
                        case "3":
                            multiplier = 6.88;
                            break;
                        case "3.5":
                            multiplier = 9.03;
                            break;
                        case "4":
                            multiplier = 11.01;
                            break;
                    }
                }
                else if (pCntType == "GRC")
                {
                    //GRC Multipliers
                    switch (pSize)
                    {
                        case "0.5":
                            multiplier = 0.95;
                            break;
                        case "0.75":
                            multiplier = 1.32;
                            break;
                        case "1":
                            multiplier = 2.02;
                            break;
                        case "1.25":
                            multiplier = 2.96;
                            break;
                        case "1.5":
                            multiplier = 3.62;
                            break;
                        case "2":
                            multiplier = 5.18;
                            break;
                        case "2.5":
                            multiplier = 7.97;
                            break;
                        case "3":
                            multiplier = 10.87;
                            break;
                        case "3.5":
                            multiplier = 13.60;
                            break;
                        case "4":
                            multiplier = 16.48;
                            break;
                        case "5":
                            multiplier = 23.84;
                            break;
                        case "6":
                            multiplier = 32.28;
                            break;
                    }
                }
                else
                {
                    multiplier = 0;
                }
            }
            else if (weightTable == "Conductor")
            {
                /*Conductor Weight Tables*/
                if (pCntType == "EMT")
                {
                    //EMT Multipliers
                    switch (pSize)
                    {
                        case "0.5":
                            multiplier = 0.51;
                            break;
                        case "0.75":
                            multiplier = 0.84;
                            break;
                        case "1":
                            multiplier = 1.30;
                            break;
                        case "1.25":
                            multiplier = 2.12;
                            break;
                        case "1.5":
                            multiplier = 2.70;
                            break;
                        case "2":
                            multiplier = 4.02;
                            break;
                        case "2.5":
                            multiplier = 5.79;
                            break;
                        case "3":
                            multiplier = 8.26;
                            break;
                        case "3.5":
                            multiplier = 10.98;
                            break;
                        case "4":
                            multiplier = 13.64;
                            break;
                    }
                }
                else if (pCntType == "GRC")
                {
                    //GRC Multipliers
                    switch (pSize)
                    {
                        case "0.5":
                            multiplier = 1.01;
                            break;
                        case "0.75":
                            multiplier = 1.46;
                            break;
                        case "1":
                            multiplier = 2.19;
                            break;
                        case "1.25":
                            multiplier = 3.18;
                            break;
                        case "1.5":
                            multiplier = 4.09;
                            break;
                        case "2":
                            multiplier = 5.94;
                            break;
                        case "2.5":
                            multiplier = 9.01;
                            break;
                        case "3":
                            multiplier = 12.59;
                            break;
                        case "3.5":
                            multiplier = 16.04;
                            break;
                        case "4":
                            multiplier = 19.67;
                            break;
                        case "5":
                            multiplier = 28.76;
                            break;
                        case "6":
                            multiplier = 40.03;
                            break;
                    }
                }
                else
                {
                    multiplier = 0;
                }
            }
            return multiplier;
        } //end CalcWeightPerFoot()
    } //end class
}
