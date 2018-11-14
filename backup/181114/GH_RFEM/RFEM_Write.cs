using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;

using Grasshopper.Kernel;
using Rhino.Geometry;
using Dlubal.RFEM5;
using System.Windows.Forms;

namespace GH_RFEM
{
    public class RFEM_Write : GH_Component
    {
        //definition of variables used in further functions
        IApplication app = null;
        IModel model = null;

        /// <summary>
        /// Initializes a new instance of the MyComponent1 class.
        /// </summary>
        public RFEM_Write()
          : base("Write to RFEM", "RfWrite",
              "Write Grasshopper Data to RFEM",
              "RFEM", "Elements")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            //pManager.AddPointParameter("Point", "Pt", "Input Rhino points you want to create as RFEM notes", GH_ParamAccess.list);
            pManager.AddGenericParameter("RFEM Nodes", "RfNd", "Input RFEM Node elements for writing to RFEM program", GH_ParamAccess.list);

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Message", "Msg", "This is status message for writing RFEM data", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            //definition of input data
            // List<Dlubal.RFEM5.Node> RfemNodesInput = new List<Dlubal.RFEM5.Node>();
            // if (!DA.GetDataList<Dlubal.RFEM5.Node>(0, RfemNodesInput)) return;

            //List<Dlubal.RFEM5.Node> RfemNodesInput = new List<Dlubal.RFEM5.Node>();
             List<Grasshopper.Kernel.Types.GH_ObjectWrapper> obj = new List<Grasshopper.Kernel.Types.GH_ObjectWrapper>();
            DA.GetDataList(0, obj);

            //   if (!DA.GetDataList<Dlubal.RFEM5.Node>(0, RfemNodesInput)) return;

            //reference to method processing data
            String OutputMessage = WriteRfemData(obj);

            // Finally assign the spiral to the output parameter.
            DA.SetData(0, OutputMessage);
        }

        private String WriteRfemData(List<Grasshopper.Kernel.Types.GH_ObjectWrapper> RfemNodes)
        {
            String StatusMsg = "Status Not set.";

           // List<Dlubal.RFEM5.Node> RfemNodes = new List<Dlubal.RFEM5.Node>();

           // foreach (RfemNodeType rfemNode in RFEMNodesWrapper)
           // {
           //     Dlubal.RFEM5.Node rfemNode = rfemNode.;
           //     //RfemNodeType rfemNodeWrapper = new RfemNodeType(rfemNode);
           //     RfemNodes.Add(rfemNode);
           // }


            // Gets interface to running RFEM application.
            app = Marshal.GetActiveObject("RFEM5.Application") as IApplication;
            // Locks RFEM licence
            app.LockLicense();

            // Gets interface to active RFEM model.
            model = app.GetActiveModel();

            // Gets interface to model data.
            IModelData data = model.GetModelData();

            //List<Dlubal.RFEM5.Node> RfemNodeList = new List<Dlubal.RFEM5.Node>();
            //Dlubal.RFEM5.Node[] RfemNodeArray = new Dlubal.RFEM5.Node[Rh_pt3d.Count];

            try
            {
                // Sets all objects to model data.
                data.PrepareModification();

                for (int index = 0; index < RfemNodes.Count; index++)

                {
                    
                    //Write Nodes
                    Dlubal.RFEM5.Node RfemNodeDlubal;

                   // RfemNodeDlubal.X = RfemNodes[index].x

                    RfemNodeDlubal = (Dlubal.RFEM5.Node) RfemNodes[index].Value;
                    //Dlubal.RFEM5.Node RfemNodeDlubal = RfemNodes[index].;
                    data.SetNode(RfemNodeDlubal);
                }

                data.FinishModification();
                StatusMsg = "OK";
            }

            catch (Exception ex)
            {
                StatusMsg = ex.Message;
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


            return StatusMsg;
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return null;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("5e00efaa-e9cd-4ccc-8afd-775be0e941b7"); }
        }
    }
}