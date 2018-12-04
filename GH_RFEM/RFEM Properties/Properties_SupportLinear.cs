using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Rhino.Geometry;

namespace GH_RFEM
{
    public class Properties_SupportLinear : GH_Component
    {

        /// <summary>
        /// Initializes a new instance of the Nodal Support class.
        /// </summary>
        public Properties_SupportLinear()
          : base("Linear Support RFEM", "Line Support",
              "Create Linear support to be applied for line",
              "RFEM", "Properties")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        /// 
        double uX = -1;
        double uY = -1;
        double uZ = -1;
        double rX = -1;
        double rY = -1;
        double rZ = -1;
        string Comment = "";

        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            // to import lists or trees of values, modify the ParamAccess flag.
            pManager.AddNumberParameter("Spring X [N/m2]", "uX", "set to -1=fixed (default), 0=free, other values create spring", GH_ParamAccess.item, uX);
            pManager.AddNumberParameter("Spring Y [N/m2]", "uY", "set to -1=fixed (default), 0=free, other values create spring", GH_ParamAccess.item, uY);
            pManager.AddNumberParameter("Spring Z [N/m2]", "uZ", "set to -1=fixed (default), 0=free, other values create spring", GH_ParamAccess.item, uZ);
            pManager.AddNumberParameter("Spring X rotation [Nm/rad/m]", "rX", "set to -1=fixed (default), 0=free, other values create spring", GH_ParamAccess.item, rX);
            pManager.AddNumberParameter("Spring Y rotation [Nm/rad/m]", "rY", "set to -1=fixed (default), 0=free, other values create spring", GH_ParamAccess.item, rY);
            pManager.AddNumberParameter("Spring Z rotation [Nm/rad/m]", "rZ", "set to -1=fixed (default), 0=free, other values create spring", GH_ParamAccess.item, rZ);
            pManager.AddTextParameter("Comment", "Comment", "This text will be written in 'comments' parameter in RFEM data", GH_ParamAccess.item, Comment);

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        /// 
        Dlubal.RFEM5.LineSupport linearSupport = new Dlubal.RFEM5.LineSupport();
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("RFEM Nodal Support definition", "Nodal Support", "RFEM nodal support definition for use with node that writes nodes", GH_ParamAccess.item);

        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {


            DA.GetData(0, ref uX);
            DA.GetData(1, ref uY);
            DA.GetData(2, ref uZ);
            DA.GetData(3, ref rX);
            DA.GetData(4, ref rY);
            DA.GetData(5, ref rZ);
            DA.GetData(6, ref Comment);

            //linearSupport.No numbers of supports not assigned here - these are assigned when writing nodes
            linearSupport.SupportConstantX = uX;
            linearSupport.SupportConstantY = uY;
            linearSupport.SupportConstantZ = uZ;
            linearSupport.RestraintConstantX = rX;
            linearSupport.RestraintConstantY = rY;
            linearSupport.RestraintConstantZ = rZ;
            linearSupport.Comment = Comment;

            DA.SetData(0, linearSupport);
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
            get { return new Guid("53faafee-fe0e-49ea-8c3d-46827609962a"); }
        }
    }
}