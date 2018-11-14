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
        //definition of variables used in further functions
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
          : base("Node RFEM", "RfNd",
              "Create RFEM notes from Rhino points",
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
            
            //pManager.AddPlaneParameter("Plane", "P", "Base plane for spiral", GH_ParamAccess.item, Plane.WorldXY);
            //pManager.AddNumberParameter("Inner Radius", "R0", "Inner radius for spiral", GH_ParamAccess.item, 1.0);
            //pManager.AddNumberParameter("Outer Radius", "R1", "Outer radius for spiral", GH_ParamAccess.item, 10.0);
            //pManager.AddIntegerParameter("Turns", "T", "Number of turns between radii", GH_ParamAccess.item, 10);

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
            //pManager.AddCurveParameter
            pManager.AddGenericParameter("RFEM nodes", "RfNd", "RFEM nodes for export in RFEM", GH_ParamAccess.list);
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

            //DA.GetDataList<Rhino.Geometry.Points>(0, myPointsList);

            //List<Point3d> rhino_points3d = new List<Point3d>();


            // Then we need to access the input parameters individually. 
            // When data cannot be extracted from a parameter, we should abort this method.
            if (!DA.GetDataList<Rhino.Geometry.Point3d>(0, rhino_points3d)) return;


            // We should now validate the data and warn the user if invalid data is supplied.
            //if (radius0 < 0.0)
            //{
            //    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Inner radius must be bigger than or equal to zero");
            //    return;
            //}


            // We're set to create the spiral now. To keep the size of the SolveInstance() method small, 
            // The actual functionality will be in a different method:

            List<Dlubal.RFEM5.Node> RfemNodes = CreateRfemNodes(rhino_points3d);

            //Curve spiral = CreateSpiral(plane, radius0, radius1, turns);

            // Finally assign the spiral to the output parameter.
            DA.SetData(0, RfemNodes);
        }

        private List<Dlubal.RFEM5.Node> CreateRfemNodes(List<Point3d> Rh_pt3d)
        {

                // Gets interface to running RFEM application.
                app = Marshal.GetActiveObject("RFEM5.Application") as IApplication;
            // Locks RFEM licence
            //app.LockLicense();

            // Gets interface to active RFEM model.
            //model = app.GetActiveModel();

            // Gets interface to model data.
            //IModelData data = model.GetModelData();

            //List<Dlubal.RFEM5.Node> RfemNodeList = new List<Dlubal.RFEM5.Node>();
            Dlubal.RFEM5.Node[] RfemNodeArray = new Dlubal.RFEM5.Node[Rh_pt3d.Count];


            try
            {
                // modification
                // Sets all objects to model data.
                // data.PrepareModification();

            for (int index = 0; index < Rh_pt3d.Count; index++)

            {

                //Dlubal.RFEM5.Node RfemNode = new Dlubal.RFEM5.Node();
                // Node RfemNode = new Node();
                RfemNodeArray[index].No = index+1;
                RfemNodeArray[index].X = Rh_pt3d[index].X;
                RfemNodeArray[index].Y = Rh_pt3d[index].Y;
                RfemNodeArray[index].Z = Rh_pt3d[index].Z;


                //data.SetNode(RfemNodeArray[index]);
            }

            //modification    
            //data.FinishModification();
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
                    //app.UnlockLicense();
                    app = null;
                }

                // Cleans Garbage Collector and releases all cached COM interfaces.
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();

            List<Dlubal.RFEM5.Node> RfemNodeList = RfemNodeArray.OfType<Dlubal.RFEM5.Node>().ToList(); // this isn't going to be fast.
            List<RfemNodeType> RfemNodeGHParamList = new List<RfemNodeType>();

           // List<RfemNodeType> aaaa = new List<RfemNodeType>();

            foreach (Dlubal.RFEM5.Node rfemNode in RfemNodeList)
            {
                RfemNodeType rfemNodeWrapper = new RfemNodeType(rfemNode);
                RfemNodeGHParamList.Add(rfemNodeWrapper);
            }

            return RfemNodeList;

                            
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
                return null;
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
