using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;

using Rhino.Geometry;
using Dlubal.RFEM5;

namespace GH_RFEM
{
    public class Model_Clear : GH_Component
    {

        /// <summary>
        /// Initializes a new instance of the Properties Hinge class.
        /// </summary>
        public Model_Clear()
          : base("Delete all elements from current RFEM model", "RFEM model clear",
              "Output of this element can be used to trigger all write \n fuctions - ensuring that all further operations are done in 'clean' model.",
              "RFEM", "Write")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        /// 
        bool run = false;


        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            // to import lists or trees of values, modify the ParamAccess flag.
            pManager.AddBooleanParameter("Run the element", "run", "Deletes all information from model if 'true'.", GH_ParamAccess.item, run);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        /// 
        bool success = false;
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddBooleanParameter("This element can be used as a toggle for all other 'write'  elements.", "Success (bool)", "True if all informaion from model is successfully deleted.", GH_ParamAccess.item);

        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            //get data from parameters
            DA.GetData(0, ref run);

            //perform model deletion
            if (run==true)
            {
                IModel model = Marshal.GetActiveObject("RFEM5.Model") as IModel;
                model.GetApplication().LockLicense();

                // cleans all model
                model.Clean();

                //unlocks applicaton
                model.GetApplication().UnlockLicense();

                //sets success to "true"
                success = true;

                //set data for output
                DA.SetData(0, success);
            }
            else
            {
                DA.SetData(0, false);
            }

        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                return GH_RFEM.Properties.Resources.icon_explode;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("932d3277-2c22-4e78-9119-85a87178caac"); }
        }
    }
}