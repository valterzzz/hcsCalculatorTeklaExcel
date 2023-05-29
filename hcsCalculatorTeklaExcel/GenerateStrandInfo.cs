using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Drawing;
using Tekla.Structures.Model;

namespace hcsCalculatorTeklaExcel
{
    public static class GenerateStrandInfo
    {
        //IDictionary<string, string>
        public static ArrayList GetStrandInfo(Tekla.Structures.Model.ModelObject hallowCoreSlab, string strandSettings, Model model)
        {

            //Add the strand component based on settings in Hallow core unit

            Tekla.Structures.Model.Assembly objAsAssembly = (Assembly)hallowCoreSlab;

            Tekla.Structures.Model.Component strandComponent = new Tekla.Structures.Model.Component();

            strandComponent.Name = "Hallowcore reinforcement strands";
            strandComponent.Number = 140000060;

            Tekla.Structures.Model.ComponentInput CI = new ComponentInput();

            CI.AddInputObject(objAsAssembly.GetMainPart());

            strandComponent.SetComponentInput(CI);

            strandComponent.LoadAttributesFromFile(strandSettings);
            strandComponent.SetAttribute("aStrandsClassAttr", "999999");

            if (!strandComponent.Insert())
            {
                MessageBox.Show("Something did not work!");
            }

            model.CommitChanges();

            //Get Slab bottom level for calculation
            string bottomLevelGlobal = "";

            objAsAssembly.GetMainPart().GetReportProperty("BOTTOM_LEVEL_GLOBAL", ref bottomLevelGlobal);

            double hallowCoreBottomLevelDouble = Math.Round(Double.Parse(bottomLevelGlobal) * 1000, 0);

            //Get Strand settings from component dialog

            string strandHeightsComponent = "";

            strandComponent.GetAttribute("strandY", ref strandHeightsComponent);

            //Parse the componenet input string and get actual possible heights

            List<double> strandHeights = new List<double>();

            string[] strandHeightCompList =  strandHeightsComponent.Split(' ');

            double lastHeight = 0.00;

            foreach(string height in strandHeightCompList)
            {
                if (height.Contains("*") == true)
                {
                    string[] heightsSplitAtMulti = height.Split('*');

                    int countOfStrandHeights = int.Parse(heightsSplitAtMulti[0]);

                    double strandHeight = Math.Round(double.Parse(heightsSplitAtMulti[1]), 0);

                    for (int i = 1; i <= countOfStrandHeights; i++)
                    {
                        double strandLineHeight = strandHeight + lastHeight;

                        strandHeights.Add(strandLineHeight);

                        lastHeight = strandLineHeight;
                    }

                }
                else
                {
                    double strandLineHeight = Double.Parse(height);

                    strandHeights.Add(strandLineHeight + lastHeight);

                    lastHeight = strandLineHeight + lastHeight;
                }

            }

            List<int> strandCount = new List<int>();

            List<double> strandList = new List<double>();

            //Get the length of hallow core slab

            double hallowCoreSlabLength = 0.00;

            objAsAssembly.GetReportProperty("CAST_UNIT_LENGTH_ONLY_CONCRETE_PARTS", ref hallowCoreSlabLength);



            foreach (var child in objAsAssembly.GetMainPart().GetChildren())
            {
                if (child.GetType().ToString() == "Tekla.Structures.Model.RebarStrand")
                {
                    RebarStrand strands = (RebarStrand)child;

                    //Only take in account fresly added strands, not what was there before.
                    if (strands.Class == 999999)
                    {
                        foreach (RebarGeometry strand in strands.GetRebarGeometries(Reinforcement.RebarGeometryOptionEnum.NONE))
                        {

                            //Only add the strand to strand count if it is not cut (If lenght matches element lenght) with Treshold of 5 mm.
                            if (Math.Abs(Math.Round(strand.Shape.Length(), 0) - hallowCoreSlabLength) < 5.00)
                            {

                                Tekla.Structures.Geometry3d.Point point = (Tekla.Structures.Geometry3d.Point)strand.Shape.Points[0];

                                double strandZGlobal = Math.Round(point.Z, 0);

                                double strandZRelativeToHCS = strandZGlobal - hallowCoreBottomLevelDouble;

                                strandList.Add(strandZRelativeToHCS);

                            }

                        }
                    }

                    

                }

            }

            //Initialize this list where you can get how many strands will be there with each specific height
            foreach (double strandHeight in strandHeights)
            {
                strandCount.Add(0);
            }

            foreach (double strand in strandList)
            {
                int indexOfStrand = strandHeights.IndexOf(strand);

                strandCount[indexOfStrand] = strandCount[indexOfStrand] + 1;

            }

            //Get Pull value and size of strand
            List<double> strandPull = new List<double>();
            List<double> strandSize = new List<double>();

            int indexForIter = 0;

            foreach (int count in strandCount)
            {
                double strandPullComponent = 0.00;
                strandComponent.GetAttribute(indexForIter.ToString() + "_SP_Pull", ref strandPullComponent);

                strandPull.Add(strandPullComponent);

                string strandSizeComponent = "";
                strandComponent.GetAttribute(indexForIter.ToString() + "_SP_Size", ref strandSizeComponent);

                strandSize.Add(double.Parse(strandSizeComponent));

                indexForIter++;

            }

            ArrayList result = new ArrayList();

            int indexForLoop = 0;

            //Add all to one list you will pass back to call funciton
            foreach (double height in strandHeights)
            {
                ArrayList dataset = new ArrayList();

                dataset.Add(strandHeights[indexForLoop]);
                dataset.Add(strandCount[indexForLoop]);
                dataset.Add(strandPull[indexForLoop]);
                dataset.Add(strandSize[indexForLoop]);

                indexForLoop++;

                result.Add(dataset);
            }

            strandComponent.Delete();

            model.CommitChanges();

            return result;


        }



    }
}
