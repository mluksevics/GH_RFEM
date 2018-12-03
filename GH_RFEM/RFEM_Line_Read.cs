using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;

using Rhino.Geometry;
using Dlubal.RFEM5;
using System.Windows.Forms;

namespace GH_RFEM
{
    public class RFEM_Line_Read : GH_Component
    {
        //definition of RFEM and input variables used in further methods
        IApplication app = null;
        IModel model = null;
        bool run = false;
        double segmentLength = 1;


        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public RFEM_Line_Read()
          : base("Read Lines from RFEM", "RFEM Ln Read",
              "Create Rhino lines from RFEM data",
              "RFEM", "Read")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            // Use the pManager object to register your input parameters.
            // You can often supply default values when creating parameters.
            // All parameters must have the correct access type. If you want 
            // to import lists or trees of values, modify the ParamAccess flag.
            pManager.AddBooleanParameter("Run", "Toggle", "Toggles whether the lines are written to RFEM", GH_ParamAccess.item, false);

            // If you want to change properties of certain parameters, 
            // you can use the pManager instance to access them by index:
            //pManager[0].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            // Use the pManager object to register your output parameters.
            // Output parameters do not have default values, but they too must have the correct access type.

            pManager.AddCurveParameter("Rhino curves", "Curves", "Rhino curves read from RFEM", GH_ParamAccess.list);

            // Sometimes you want to hide a specific parameter from the Rhino preview.
            // You can use the HideParameter() method as a quick way:
            //pManager.HideParameter(0);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object can be used to retrieve data from input parameters and 
        /// to store data in output parameters.</param>

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            // First, we need to retrieve all data from the input parameters.
            // We'll start by declaring variables and assigning them starting values.
            List<Rhino.Geometry.Curve> RhinoCurves = new List<Curve>();
            string lineList = "all";


            // Then we need to access the input parameters individually. 
            // When data cannot be extracted from a parameter, we should abort this method.
            DA.GetData(0, ref run);

            // The actual functionality will be in a method defined below. This is where we run it
            if (run == true)
            {
                RhinoCurves = CreateRhinoCurves(lineList);
            }

            // Finally assign the processed data to the output parameter.
            DA.SetDataList(0, RhinoCurves);
        }

        private List<Rhino.Geometry.Curve> CreateRhinoCurves(string selectedLineList)
        {
            //defining the list with lines that will have to be returned later on
            List<Rhino.Geometry.Curve> rhOutputCurves = new List<Rhino.Geometry.Curve>();


            // Gets interface to running RFEM application.
            app = Marshal.GetActiveObject("RFEM5.Application") as IApplication;
            // Locks RFEM licence
            app.LockLicense();

            // Gets interface to active RFEM model.
            model = app.GetActiveModel();

            // Gets interface to model data.
            IModelData data = model.GetModelData();

            //Create new array for Rhino Curve objects
            List<Rhino.Geometry.Curve> rhinoLineList = new List<Rhino.Geometry.Curve>();
            
            //Create new array for Rhino point objects
            List<Rhino.Geometry.Point3d> rhinoPointArray = new List<Rhino.Geometry.Point3d>();


            try
            {
                for (int index = 0; index < data.GetLineCount(); index++)
                {

                    Dlubal.RFEM5.Line currentLine = data.GetLine(index, ItemAt.AtIndex).GetData();

                    // the code below converts string describing nodes used in line definition into
                    // list fo all used node numbers. e.g. converts "1,2,4-7,9" into "1,2,4,5,6,7,9"
                    string lineNodes = currentLine.NodeList;
                    List<int> nodesList = new List<int>();

                    foreach (string tempLineNode in lineNodes.Split(','))
                    {
                        if (tempLineNode.Contains('-'))
                        {
                            string[] tempLineNodeDashes = new string[2];
                            tempLineNodeDashes = tempLineNode.Split('-');
                            int startNumber = Int32.Parse(tempLineNodeDashes[0]);
                            int endNumber = Int32.Parse(tempLineNodeDashes[1]);

                            for (int i = startNumber; i <= endNumber; i++)
                            {
                                nodesList.Add(i);
                            }
                        }
                        else
                        {
                            nodesList.Add(Int32.Parse(tempLineNode));
                        }
                    }

                    //

                    if (currentLine.Type == LineType.PolylineType)
                    {

                        for (int i = 0; i < nodesList.Count-1; i++)
                        {
                            Dlubal.RFEM5.Node rfemStartPoint = data.GetNode(nodesList[i], ItemAt.AtNo).GetData();
                            Dlubal.RFEM5.Node rfemEndPoint = data.GetNode(nodesList[i+1], ItemAt.AtNo).GetData();

                            Point3d rhinoStartPoint = new Point3d(rfemStartPoint.X, rfemStartPoint.Y, rfemStartPoint.Z);
                            Point3d rhinoEndPoint = new Point3d(rfemEndPoint.X, rfemEndPoint.Y, rfemEndPoint.Z);

                            Rhino.Geometry.LineCurve currentRhinoCurve = new Rhino.Geometry.LineCurve(rhinoStartPoint,rhinoEndPoint);

                            rhOutputCurves.Add(currentRhinoCurve);

                        }
                    }
                    

                }


            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // Releases interface to RFEM model.
            model = null;

            // Unlocks licence and releases interface to RFEM application.
            if (app != null)
            {
                app.UnlockLicense();
                app = null;
            }

            // Cleans Garbage Collector and releases all cached COM interfaces.
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();

            //list with all Rhino Curves is prepared for output
            return rhOutputCurves;


        }



        /// <summary>
        /// The Exposure property controls where in the panel a component icon 
        /// will appear. There are seven possible locations (primary to septenary), 
        /// each of which can be combined with the GH_Exposure.obscure flag, which 
        /// ensures the component will only be visible on panel dropdowns.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.primary; }
        }

        /// <summary>
        /// Provides an Icon for every component that will be visible in the User Interface.
        /// Icons need to be 24x24 pixels.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                // You can add image files to your project resources and access them like this:
                //return Resources.IconForThisComponent;
                return GH_RFEM.Properties.Resources.icon_line;
            }
        }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("db6ea69a-2394-41b0-995e-4b985cbdeb70"); }
        }
    }
}
