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
    public class Node_Write : GH_Component
    {


        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public Node_Write()
          : base("Node RFEM", "RFEM Nd Write",
              "Create RFEM nodes from Rhino points",
              "RFEM", "Write")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        /// 

        // We'll start by declaring input parameters and assigning them starting values.
        List<Rhino.Geometry.Point3d> rhinoPointsInput = new List<Point3d>();
        Dlubal.RFEM5.NodalSupport rfemNodalSupportInput = new Dlubal.RFEM5.NodalSupport();
        string commentsInput = "";
        bool run = false;

        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            // Use the pManager object to register your input parameters.
            pManager.AddPointParameter("Point", " Rhino Points", "Input Rhino points you want to create as RFEM notes", GH_ParamAccess.list);
            pManager.AddGenericParameter("RFEM Support", "Support", "Leave empty if no support, use 'Nodal Support' node to generate support", GH_ParamAccess.item);
            pManager.AddTextParameter("Text written in RFEM comments field", "Comment", "Text written in RFEM comments field", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Run", "Toggle", "Toggles whether the nodes are written to RFEM", GH_ParamAccess.item, run);

            pManager[1].Optional = true;
            pManager[2].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        /// 
        //declaring output parameters
        List<Dlubal.RFEM5.Node> RfemNodes = new List<Dlubal.RFEM5.Node>();
        bool writeSuccess = false;

        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            // Use the pManager object to register your output parameters.
            pManager.AddGenericParameter("RFEM nodes", "RfNd", "Generated RFEM nodes", GH_ParamAccess.list);
            pManager.AddBooleanParameter("Success", "Success", "Returns 'true' if nodes successfully writen, otherwise it is 'false'", GH_ParamAccess.item);

        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object can be used to retrieve data from input parameters and 
        /// to store data in output parameters.</param>

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            // First, we need to retrieve all data from the input parameters.

            // Then we need to access the input parameters individually. 
            // When data cannot be extracted from a parameter, we should abort this method.
            if (!DA.GetDataList<Rhino.Geometry.Point3d>(0, rhinoPointsInput)) return;
            DA.GetData(1, ref rfemNodalSupportInput);
            DA.GetData(2, ref commentsInput);
            DA.GetData(3, ref run);

            // The actual functionality will be in a method defined below. This is where we run it
            if (run == true)
            {
                //clears and resets all output parameters:
                RfemNodes.Clear();
                writeSuccess = false;

                //runs the method for creating RFEM nodes
                RfemNodes = CreateRfemNodes(rhinoPointsInput, rfemNodalSupportInput, commentsInput);
                DA.SetData(1, writeSuccess);
            }
            else
            {
                DA.SetData(1, false);
            }

            // Finally assign the processed data to the output parameter.
            DA.SetDataList(0, RfemNodes);


            // clear and reset all input parameters
            rhinoPointsInput.Clear();
            commentsInput = "";
            rfemNodalSupportInput = new Dlubal.RFEM5.NodalSupport();
        }

        private List<Dlubal.RFEM5.Node> CreateRfemNodes(List<Point3d> Rh_pt3d, Dlubal.RFEM5.NodalSupport rfemNodalSupportMethodIn, string commentsListMethodIn)
        {

            // Gets interface to running RFEM application.
            IApplication app = Marshal.GetActiveObject("RFEM5.Application") as IApplication;
            // Locks RFEM licence
            app.LockLicense();

            // Gets interface to active RFEM model.
            IModel model = app.GetActiveModel();

            // Gets interface to model data.
            IModelData data = model.GetModelData();

            // Gets Max node Number
            int currentNewNodeNo = 1;
            int totalNodesCount = data.GetNodes().Count();
            if (totalNodesCount !=0)
            {
                int lastNodeNo = data.GetNode(totalNodesCount - 1, ItemAt.AtIndex).GetData().No;
                currentNewNodeNo = lastNodeNo + 1;
            }

            // Gets Max nodal support number
            int currentNewNodalSupportNo = 1;
            int totalNodalSupportsCount = data.GetNodalSupports().Count();
            if (totalNodalSupportsCount != 0)
            {
                int lastNodalSupportNo = data.GetNodalSupport(totalNodalSupportsCount - 1, ItemAt.AtIndex).GetData().No;
                currentNewNodalSupportNo = lastNodalSupportNo + 1;
            }

            //Create new array for RFEM point objects
            Dlubal.RFEM5.Node[] RfemNodeArray = new Dlubal.RFEM5.Node[Rh_pt3d.Count];
            //Create a string for list with nodes
            string createdNodesList ="";

            ///This version writes nodes one-by-one, because the API method data.SetNodes() for
            ///array appears not to be working
            try
            {
                // modification - Sets all objects to model data.
                data.PrepareModification();


                for (int i = 0; i < Rh_pt3d.Count; i++)
                {
                    RfemNodeArray[i].No = currentNewNodeNo;
                    RfemNodeArray[i].X = Rh_pt3d[i].X;
                    RfemNodeArray[i].Y = Rh_pt3d[i].Y;
                    RfemNodeArray[i].Z = Rh_pt3d[i].Z;
                    RfemNodeArray[i].Comment = commentsListMethodIn;

                    data.SetNode(RfemNodeArray[i]);

                    if (i== Rh_pt3d.Count - 1)
                    {
                        createdNodesList = createdNodesList + currentNewNodeNo.ToString();
                    }
                    else
                    {
                        createdNodesList = createdNodesList + currentNewNodeNo.ToString() + ",";
                    }

                    currentNewNodeNo++;
                }

                //addition of nodal supports
                    rfemNodalSupportMethodIn.No = currentNewNodalSupportNo;
                    rfemNodalSupportMethodIn.NodeList = createdNodesList;
                    data.SetNodalSupport(ref rfemNodalSupportMethodIn);


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
            List<Dlubal.RFEM5.Node> RfemNodeList = RfemNodeArray.OfType<Dlubal.RFEM5.Node>().ToList();

            //output 'success' as true 
            writeSuccess = true;
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
