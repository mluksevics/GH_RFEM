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
    public class Member_Write : GH_Component
    {


        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public Member_Write()
          : base("Member RFEM", "RFEM Mbr Write",
              "Create RFEM members from RFEM lines",
              "RFEM", "Write")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        /// 

        // We'll start by declaring input parameters and assigning them starting values.
        List<Dlubal.RFEM5.Line> rfemLinesInput = new List<Dlubal.RFEM5.Line>();
        string sectionIdInput = "UB 203x133x25 (Corus)";
        string materialIdInput = "NameID | Beton C30/37@TypeID|CONCRETE @NormID|DIN 1045-1 - 08";
        Dlubal.RFEM5.MemberHinge rfemHingeStartInput = new Dlubal.RFEM5.MemberHinge();
        Dlubal.RFEM5.MemberHinge rfemHingeEndInput = new Dlubal.RFEM5.MemberHinge();
        double rotationInput = 0;
        string commentsInput = "";
        bool run = false;

        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            // Use the pManager object to register your input parameters.
            pManager.AddGenericParameter("Line", "RFEM Lines", "Input lines created with 'RFEM Ln' component", GH_ParamAccess.list);
            pManager.AddTextParameter("Member Section", "Section[code]", "Section ID. You can use note in 'properties' subsection to select section. \n Alternatively you can use RFEM section naming conventions and assign as text to this input parameter \n e.g. 'Rectangle 500/600' or 'Circle 500'", GH_ParamAccess.item, sectionIdInput);
            pManager.AddTextParameter("Member Material", "Material[code]", "Material TextID according to Dlubal RFEM naming system (see their API documentation. \n examples:  \n NameID|Beton C30/37@TypeID|CONCRETE@NormID|EN 1992-1-1 \n NameID|Baustahl S 235@TypeID|STEEL@NormID|EN 1993-1-1 \n NameID|Pappel und Nadelholz C24@TypeID|CONIFEROUS@NormID|EN 1995-1-1", GH_ParamAccess.item, materialIdInput);
            pManager.AddGenericParameter("Hinge at member start", "Start Hinge", "Input Hinges created using 'RFEM Hinge' node", GH_ParamAccess.item);
            pManager.AddGenericParameter("Hinge at member end", "End Hinge", "Input Hinges created using 'RFEM Hinge' node", GH_ParamAccess.item);
            pManager.AddNumberParameter("Rotation (beta) angle ", "Rotation[deg]", "Input rotation of section about local X axis", GH_ParamAccess.item, rotationInput);
            pManager.AddTextParameter("Text to be written in RFEM comments field", "Comment", "Text written in RFEM comments field", GH_ParamAccess.item, commentsInput);
            pManager.AddBooleanParameter("Run", "Toggle", "Toggles whether the Surfaces are written to RFEM", GH_ParamAccess.item, run);

            pManager[1].Optional = true;
            pManager[2].Optional = true;
            pManager[3].Optional = true;
            pManager[4].Optional = true;
            pManager[5].Optional = true;
            pManager[6].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        /// 
        //declaring output parameters
        List<Dlubal.RFEM5.Member> RfemMembers = new List<Dlubal.RFEM5.Member>();
        bool writeSuccess = false;

        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            // Use the pManager object to register your output parameters.
            pManager.AddGenericParameter("RFEM members", "RfMember", "Generated RFEM members", GH_ParamAccess.list);
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
            if (!DA.GetDataList<Dlubal.RFEM5.Line>(0, rfemLinesInput)) return;
            DA.GetData(1, ref sectionIdInput);
            DA.GetData(2, ref materialIdInput);
            DA.GetData(3, ref rfemHingeStartInput);
            DA.GetData(4, ref rfemHingeEndInput);
            DA.GetData(5, ref rotationInput);
            DA.GetData(6, ref commentsInput);
            DA.GetData(7, ref run);

            // The actual functionality will be in a method defined below. This is where we run it
            if (run == true)
            {
                //clears and resets all output parameters:
                RfemMembers.Clear();
                writeSuccess = false;

                //runs the method for creating RFEM nodes
                RfemMembers = CreateRfemMembers(rfemLinesInput, sectionIdInput, materialIdInput, rfemHingeStartInput, rfemHingeEndInput, rotationInput, commentsInput);
                DA.SetData(1, writeSuccess);
            }
            else
            {
                DA.SetData(1, false);
            }

            // Finally assign the processed data to the output parameter.
            DA.SetDataList(0, RfemMembers);


            // clear and reset all input parameters
            rfemLinesInput.Clear();
            sectionIdInput = "";
            materialIdInput = "";
            rfemHingeStartInput = new Dlubal.RFEM5.MemberHinge();
            rfemHingeEndInput = new Dlubal.RFEM5.MemberHinge();
            rotationInput = 0;
            commentsInput = "";
        }

        private List<Dlubal.RFEM5.Member> CreateRfemMembers(List<Dlubal.RFEM5.Line> rfemLineMethodIn, string sectionIdMethodIn, string materialIdMethodIn,Dlubal.RFEM5.MemberHinge rfemHingeStartMethodIn, Dlubal.RFEM5.MemberHinge rfemHingeEndMethodIn, double rotationMethodIn, string commentsListMethodIn)
        {

            // Gets interface to running RFEM application.
            IApplication app = Marshal.GetActiveObject("RFEM5.Application") as IApplication;
            // Locks RFEM licence
            app.LockLicense();

            // Gets interface to active RFEM model.
            IModel model = app.GetActiveModel();

            // Gets interface to model data.
            IModelData data = model.GetModelData();

            // Gets Member min available Number
            int currentNewMemberNo = 1;
            int totalMemberCount = data.GetMemberCount();
            if (totalMemberCount != 0)
            {
                int lastMemberNo = data.GetMember(totalMemberCount - 1, ItemAt.AtIndex).GetData().No;
                currentNewMemberNo = lastMemberNo + 1;
            }

            // Gets Cross Section min available Number
            int currentNewSectionNo = 1;
            int totalSectionCount = data.GetCrossSectionCount();
            if (totalSectionCount != 0)
            {
                int lastSectionNo = data.GetCrossSection(totalSectionCount - 1, ItemAt.AtIndex).GetData().No;
                currentNewSectionNo = lastSectionNo + 1;
            }

            // Gets Max material number
            int currentNewMaterialNo = 1;
            int totalMaterialsCount = data.GetMaterials().Count();
            if (totalMaterialsCount != 0)
            {
                int lastMaterialNo = data.GetMaterial(totalMaterialsCount - 1, ItemAt.AtIndex).GetData().No;
                currentNewMaterialNo = lastMaterialNo + 1;
            }

            // Gets Max member hinge number
            int currentNewHingeNo = 1;
            int totalHingeCount = data.GetMemberHingeCount();
            if (totalHingeCount != 0)
            {
                int lastHingeNo = data.GetMemberHinge(totalHingeCount - 1, ItemAt.AtIndex).GetData().No;
                currentNewHingeNo = lastHingeNo + 1;
            }


            //define material
            Dlubal.RFEM5.Material material = new Dlubal.RFEM5.Material();
            material.No = currentNewMaterialNo;
            material.TextID = materialIdMethodIn;
            material.ModelType = MaterialModelType.IsotropicLinearElasticType;


            //define cross section
            CrossSection tempCrossSection = new CrossSection();
            tempCrossSection.No = currentNewSectionNo;
            tempCrossSection.TextID = sectionIdMethodIn;
            tempCrossSection.MaterialNo = currentNewMaterialNo;

            //define member hinge numbers
            rfemHingeStartInput.No = currentNewHingeNo;
            rfemHingeEndMethodIn.No = currentNewHingeNo+1;



            try
            {
                // modification - Sets all objects to model data.
                data.PrepareModification();

                //set material, cross section and start&end hinges
                if (sectionIdInput != "-1")
                {
                    data.SetMaterial(material);
                    data.SetCrossSection(tempCrossSection);
                }
                data.SetMemberHinge(rfemHingeStartInput);
                data.SetMemberHinge(rfemHingeEndMethodIn);


                // process all lines and create members of those
                for (int i = 0; i < rfemLineMethodIn.Count; i++)
                {
                    Dlubal.RFEM5.Member tempMember = new Dlubal.RFEM5.Member();
                    tempMember.No = currentNewMemberNo;
                    tempMember.LineNo = rfemLineMethodIn[i].No;
                    tempMember.EndCrossSectionNo = currentNewSectionNo;
                    tempMember.StartCrossSectionNo = currentNewSectionNo;
                    tempMember.TaperShape = TaperShapeType.Linear;
                    tempMember.StartHingeNo = currentNewHingeNo;
                    tempMember.EndHingeNo = currentNewHingeNo+1;
                    tempMember.Comment = commentsListMethodIn;
                    // if -1 is input as section, member is created as rigid, otherwise it is "standard"
                    if (sectionIdInput == "-1")
                    {
                        tempMember.Type = MemberType.Rigid;
                    }
                    else if (sectionIdInput == "0")
                    {
                        tempMember.Type = MemberType.NullMember;
                    }
                    else
                    {
                        tempMember.Type = MemberType.Beam;
                    }

                    data.SetMember(tempMember);

                    RfemMembers.Add(tempMember);
                    currentNewMemberNo++;
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

            //output 'success' as true 
            writeSuccess = true;
            return RfemMembers;

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
            get { return new Guid("730e2d43-e762-41ca-9c3f-fa332872e22d"); }
        }
    }
}
