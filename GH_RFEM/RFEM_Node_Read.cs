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
    public class RFEM_Node_Read : GH_Component
    {
        //definition of RFEM variables used in further methods
        IApplication app = null;
        IModel model = null;


        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public RFEM_Node_Read()
          : base("Node RFEM", "RFEM Nd Read",
              "Create Rhino points from RFEM nodes",
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
            pManager.AddTextParameter("List of Nodes", "Nodes list", "Input string with numbers of nodes you want to import (use commas and dashes to separate numbers, example: 1,3,4-8", GH_ParamAccess.item,"all");
            pManager.AddBooleanParameter("Run", "Toggle", "Toggles whether the nodes read from RFEM", GH_ParamAccess.item, false);

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

            pManager.AddPointParameter("Rhino Points", "Points", "Rhino points", GH_ParamAccess.item);

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
            string pointsList = "all";
            bool run = false;
            List<Rhino.Geometry.Point3d> RhinoPoints = new List<Rhino.Geometry.Point3d>();

            // Then we need to access the input parameters individually. 
            // When data cannot be extracted from a parameter, we should abort this method.
            DA.GetData(0, ref pointsList);
            DA.GetData(1, ref run);

            // The actual functionality will be in a method defined below. This is where we run it
            if (run == true)
            {
                RhinoPoints = ReadRfemNodes(pointsList);
                // Finally assign the processed data to the output parameter.
                DA.SetDataList(0, RhinoPoints);

            }

        }

        private List<Rhino.Geometry.Point3d> ReadRfemNodes(string pointsListInput)
        {

            // Gets interface to running RFEM application.
            app = Marshal.GetActiveObject("RFEM5.Application") as IApplication;
            // Locks RFEM licence
            app.LockLicense();

            // Gets interface to active RFEM model.
            model = app.GetActiveModel();

            // Gets interface to model data.
            IModelData data = model.GetModelData();
            
            //Create new array for Rhino point objects
            Rhino.Geometry.Point3d[] rhinoPointArray = new Rhino.Geometry.Point3d[data.GetNodeCount()];
            GH_Point[] ghPointArray = new GH_Point[data.GetNodeCount()];

            try
            {
                for (int index = 0; index < data.GetNodeCount(); index++)
                {
                    Dlubal.RFEM5.Node currentNode = data.GetNode(index, ItemAt.AtIndex).GetData();

                    rhinoPointArray[index].X = currentNode.X;
                    rhinoPointArray[index].Y = currentNode.Y;
                    rhinoPointArray[index].Z = currentNode.Z;
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

            ///the lines below outputs created RFEM nodes in output parameter
            ///current funcionality does not use this
            ///it uses a custom class (written within this project) RfemNodeType to wrap the Dlubal.RFEM5.Node objects.
            return rhinoPointArray.ToList();


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
                return GH_RFEM.Properties.Resources.icon_node;
            }
        }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("baaf9b24-7f94-4b72-8352-25dc3adf30ff"); }
        }
    }
}
