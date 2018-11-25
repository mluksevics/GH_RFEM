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
    public class RFEM_Node : GH_Component
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
        public RFEM_Node()
          : base("Node RFEM", "RFEM Nd",
              "Create RFEM nodes from Rhino points",
              "RFEM", "Elements")
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
            pManager.AddPointParameter("Point", "Pt", "Input Rhino points you want to create as RFEM notes", GH_ParamAccess.list);
            pManager.AddBooleanParameter("Run", "Run", "Toggles whether the nodes are written to RFEM", GH_ParamAccess.item, false);

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

            //pManager.AddGenericParameter("RFEM nodes", "RfNd", "RFEM nodes for export in RFEM", GH_ParamAccess.list);

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
            List<Rhino.Geometry.Point3d> rhino_points3d = new List<Point3d>();
            bool run = false;
            List<RfemNodeType> RfemNodes = new List<RfemNodeType>();

            // Then we need to access the input parameters individually. 
            // When data cannot be extracted from a parameter, we should abort this method.
            if (!DA.GetDataList<Rhino.Geometry.Point3d>(0, rhino_points3d)) return;
            DA.GetData(1,ref run);

            // The actual functionality will be in a method defined below. This is where we run it
            if (run == true)
            {
                RfemNodes = CreateRfemNodes(rhino_points3d);
            }

            // Finally assign the processed data to the output parameter.
            //DA.SetData(0, RfemNodes);
        }

        private List<RfemNodeType> CreateRfemNodes(List<Point3d> Rh_pt3d)
        {

            // Gets interface to running RFEM application.
            app = Marshal.GetActiveObject("RFEM5.Application") as IApplication;
            // Locks RFEM licence
            app.LockLicense();

            // Gets interface to active RFEM model.
            model = app.GetActiveModel();

            // Gets interface to model data.
            IModelData data = model.GetModelData();

            //Create new array for RFEM point objects
            Dlubal.RFEM5.Node[] RfemNodeArray = new Dlubal.RFEM5.Node[Rh_pt3d.Count];

            ///This version writes nodes one-by-one because the data.SetNodes() for
            ///array appears not to be working
            try
            {
                // modification
                // Sets all objects to model data.
                 data.PrepareModification();
                
                for (int index = 0; index < Rh_pt3d.Count; index++)
                {
                    RfemNodeArray[index].No = index+1;
                    RfemNodeArray[index].X = Rh_pt3d[index].X;
                    RfemNodeArray[index].Y = Rh_pt3d[index].Y;
                    RfemNodeArray[index].Z = Rh_pt3d[index].Z;
                    data.SetNode(RfemNodeArray[index]);
                }

                // finish modification - RFEM regenerates the data
                data.FinishModification();
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
            List<Dlubal.RFEM5.Node> RfemNodeList = RfemNodeArray.OfType<Dlubal.RFEM5.Node>().ToList(); // this isn't going to be fast.
            List<RfemNodeType> RfemNodeGHParamList = new List<RfemNodeType>();

            foreach (Dlubal.RFEM5.Node rfemNode in RfemNodeList)
            {
                RfemNodeType rfemNodeWrapper = new RfemNodeType(rfemNode);
                RfemNodeGHParamList.Add(rfemNodeWrapper);
            }

            return RfemNodeGHParamList; 

                            
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
            get { return new Guid("4f8440de-7562-45a3-a869-d9a4a01ab513"); }
        }
    }
}
