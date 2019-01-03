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
    public class Surface_Write : GH_Component
    {

        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public Surface_Write()
          : base("Surface RFEM", "RFEM Srf Write",
              "Create RFEM Surfaces from Rhino Surfaces",
              "RFEM", "Write")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        /// 
        //defining input variables and default values of those
        List<Rhino.Geometry.Brep> rhino_surface = new List<Rhino.Geometry.Brep>();
        bool run = false;
        double segmentLength = 1;
        double srfThickness = 0.2;
        string srfMaterial = "NameID | Beton C30/37@TypeID|CONCRETE @NormID|DIN 1045-1 - 08";
        string commentsInput = "";

        //define grasshopper component input parameters
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            // Use the pManager object to register your input parameters.
            pManager.AddSurfaceParameter("Surface", "Surfaces", "Input Rhino planar sufaces you want to create as RFEM Surfaces", GH_ParamAccess.list);
            pManager.AddNumberParameter("Segment length", "MaxSegmentLength[m]", "Any edges with splines/circles/arcs will be simplified as segments with maximum length described in this parameter", GH_ParamAccess.item, segmentLength);
            pManager.AddNumberParameter("Surface Thickness", "Thickness[m]", "Surfaces are created as isotropic planar surface with thickness[m] assigned in this parameters \n use 0 for 'null' surfaces \n use -1 for 'rigid' surfaces", GH_ParamAccess.item, srfThickness);
            pManager.AddTextParameter("Surface Material", "Material[code]", "Material TextID according to Dlubal RFEM naming system (see their API documentation. \n examples:  \n NameID|Beton C30/37@TypeID|CONCRETE@NormID|EN 1992-1-1 \n NameID|Baustahl S 235@TypeID|STEEL@NormID|EN 1993-1-1 \n NameID|Pappel und Nadelholz C24@TypeID|CONIFEROUS@NormID|EN 1995-1-1", GH_ParamAccess.item, srfMaterial);
            pManager.AddTextParameter("Text to be written in RFEM comments field", "Comment", "Text written in RFEM comments field", GH_ParamAccess.item, commentsInput);
            pManager.AddBooleanParameter("Run", "Toggle", "Toggles whether the Surfaces are written to RFEM", GH_ParamAccess.item, run);

            pManager[1].Optional = true;
            pManager[2].Optional = true;
            pManager[3].Optional = true;
            pManager[4].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        /// 
        // defining variables for output
        bool writeSuccess = false;
        List<Dlubal.RFEM5.Surface> RfemSurfaces = new List<Dlubal.RFEM5.Surface>();

        //define grasshopper component ouput parameters
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            // Use the pManager object to register your output parameters.
            pManager.AddGenericParameter("RFEM Surfaces", "RfSrf", "RFEM Surfaces for export in RFEM", GH_ParamAccess.list);
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
            if (!DA.GetDataList<Rhino.Geometry.Brep>(0, rhino_surface)) return;
            DA.GetData(1, ref segmentLength);
            DA.GetData(2, ref srfThickness);
            DA.GetData(3, ref srfMaterial);
            DA.GetData(4, ref commentsInput);
            DA.GetData(5, ref run);

            // The actual functionality will be in a method defined below. This is where we run it
            if (run == true)
            {
                //clears and resets all output parameters.
                // this is done to ensure that if function is repeadedly run, then parameters are re-read and redefined
                RfemSurfaces.Clear();
                writeSuccess = false;

                //runs the method for creating RFEM surfaces
                RfemSurfaces = CreateRfemSurfaces(rhino_surface, srfThickness, srfMaterial, commentsInput);
                DA.SetData(1, writeSuccess);
            }
            else
            {
                // if "run" is set to false, then also the output parameter "success" is set to false
                // this ensures that as soon as "run" toogle is set "false", it automatically updates output.
                DA.SetData(1, false);
            }

            // Finally assign the processed data to the output parameter.
            DA.SetDataList(0, RfemSurfaces);

            //clears and resets all input parameters.
            // this is done to ensure that if function is repeadedly run, then parameters are re-read and redefined
            rhino_surface.Clear();
            commentsInput = "";
            srfThickness = 0;
            srfMaterial = "";
        }

        private List<Dlubal.RFEM5.Surface> CreateRfemSurfaces(List<Rhino.Geometry.Brep> Rh_Srf, double srfThicknessInMethod, string srfMaterialTextDescription, string srfCommentMethodIn)
        {
            //cycling through all Surfaces and creating RFEM objects;
            int nodeCount = 1;
            int lineCount = 1;
            int surfaceCount = 1;

            //list for curves describing surface edges
            List<Rhino.Geometry.Curve> RhSimpleLines = new List<Rhino.Geometry.Curve>();

            //lists for nodes, lines and surfaces to be created
            List<Dlubal.RFEM5.Node> RfemNodeList = new List<Dlubal.RFEM5.Node>();
            List<Dlubal.RFEM5.Line> RfemLineList = new List<Dlubal.RFEM5.Line>();
            List<Dlubal.RFEM5.Surface> RfemSurfaceList = new List<Dlubal.RFEM5.Surface>();

            //---- Interface with RFEM, getting available element numbers ----
            #region Interface with RFEM, getting available element numbers

            // Gets interface to running RFEM application.
            IApplication app = Marshal.GetActiveObject("RFEM5.Application") as IApplication;
            // Locks RFEM licence
            app.LockLicense();

            // Gets interface to active RFEM model.
            IModel model = app.GetActiveModel();

            // Gets interface to model data.
            IModelData data = model.GetModelData();

            // Gets Max node, line , surface support numbers
            int currentNewNodeNo = data.GetLastObjectNo(ModelObjectType.NodeObject) + 1;
            int currentNewLineNo = data.GetLastObjectNo(ModelObjectType.LineObject) + 1;
            int currentNewSurfaceNo = data.GetLastObjectNo(ModelObjectType.SurfaceObject) + 1;
            int currentNewMaterialNo = data.GetLastObjectNo(ModelObjectType.MaterialObject) + 1;

            #endregion

            // Defines material used for all surfaces
            Dlubal.RFEM5.Material material = new Dlubal.RFEM5.Material();
            material.No = currentNewMaterialNo;
            material.TextID = srfMaterialTextDescription;
            material.ModelType = MaterialModelType.IsotropicLinearElasticType;



            //start cycling through all surfaces
            foreach (Rhino.Geometry.Brep RhSingleSurface in Rh_Srf)
            {
                #region simplification of perimeter edges
                // clearing previous surface data before starting to work on new surface:
                RhSimpleLines.Clear();

                //simplifying edges - adding to array;
                Rhino.Geometry.Curve[] curves = RhSingleSurface.DuplicateEdgeCurves(true);
                curves = Rhino.Geometry.Curve.JoinCurves(curves);

                foreach (Rhino.Geometry.Curve RhSingleCurve in curves)
                {
                    if (RhSingleCurve.IsPolyline())
                    {
                        if (RhSingleCurve.SpanCount == 1)
                        {
                            //if this is simple linear line
                            RhSimpleLines.Add(RhSingleCurve);
                        }
                        else
                        {
                            foreach (Rhino.Geometry.Curve explodedSurface in RhSingleCurve.DuplicateSegments())
                            {
                                //if this is polyline
                                RhSimpleLines.Add(explodedSurface);
                            }
                        }
                    }

                    else
                    {

                        foreach (Rhino.Geometry.Curve explodedLine in RhSingleCurve.ToPolyline(0, 0, 3.14, 1, 0, 0, 0, segmentLength, true).DuplicateSegments())
                        {
                            //if this is curved lines
                            RhSimpleLines.Add(explodedLine);
                        }

                    }


                    #endregion

                    int surfaceNodeCounter = 0; //counts nodes witin one surface. nodeCount counts overall nodes in model
                    int surfaceLineCounter = 0; // counts lines (edges) for one surface, lineCount counts overall lines in model

                    //cycling through perimeter of the surface and defining lines surface
                    #region Defining nodes and lines for one surface

                    for (int i = 0; i < RhSimpleLines.Count; i++)
                    {
                        //defining variables needed to store geometry and RFEM info
                        Rhino.Geometry.Point3d startPoint = RhSimpleLines[i].PointAtStart;
                        Rhino.Geometry.Point3d endPoint = RhSimpleLines[i].PointAtEnd;

                        //if this is the first line for the surface
                        if (i == 0)
                        {
                            // defining start and end nodes of the line
                            Dlubal.RFEM5.Node tempCurrentStartNode = new Dlubal.RFEM5.Node();
                            Dlubal.RFEM5.Node tempCurrentEndNode = new Dlubal.RFEM5.Node();

                            tempCurrentStartNode.No = currentNewNodeNo;
                            tempCurrentStartNode.X = Math.Round(startPoint.X, 5);
                            tempCurrentStartNode.Y = Math.Round(startPoint.Y, 5);
                            tempCurrentStartNode.Z = Math.Round(startPoint.Z, 5);

                            tempCurrentEndNode.No = currentNewNodeNo + 1;
                            tempCurrentEndNode.X = Math.Round(endPoint.X, 5);
                            tempCurrentEndNode.Y = Math.Round(endPoint.Y, 5);
                            tempCurrentEndNode.Z = Math.Round(endPoint.Z, 5);

                            RfemNodeList.Add(tempCurrentStartNode);
                            RfemNodeList.Add(tempCurrentEndNode);

                            // defining line
                            Dlubal.RFEM5.Line tempCurrentLine = new Dlubal.RFEM5.Line();
                            tempCurrentLine.No = currentNewLineNo;
                            tempCurrentLine.Type = LineType.PolylineType;
                            tempCurrentLine.NodeList = $"{tempCurrentStartNode.No}, {tempCurrentEndNode.No}";

                            RfemLineList.Add(tempCurrentLine);

                            nodeCount = nodeCount + 2;
                            surfaceNodeCounter = surfaceNodeCounter + 2;
                            lineCount++;
                            surfaceLineCounter++;

                            currentNewNodeNo = currentNewNodeNo + 2;
                            currentNewLineNo++;

                        }
                        //if this is the last node for the surface
                        else if (i == RhSimpleLines.Count - 1)
                        {
                            //no need to define new node as these are both already defined
                            //create line connecting previous node with first node for surface
                            // defining start and end nodes of the line
                            Dlubal.RFEM5.Line tempCurrentLine = new Dlubal.RFEM5.Line();

                            tempCurrentLine.No = currentNewLineNo;
                            tempCurrentLine.Type = LineType.PolylineType;

                            tempCurrentLine.NodeList = $"{RfemNodeList[RfemNodeList.Count - 1].No}, {RfemNodeList[RfemNodeList.Count - surfaceNodeCounter].No}";
                            RfemLineList.Add(tempCurrentLine);
                            lineCount++;
                            surfaceLineCounter++;
                            currentNewLineNo++;

                        }
                        else
                        {

                            //if this is just a node somewhere on edges
                            //defining end node of line
                            Dlubal.RFEM5.Node tempCurrentEndNode = new Dlubal.RFEM5.Node();
                            // defining start and end nodes of the line
                            Dlubal.RFEM5.Line tempCurrentLine = new Dlubal.RFEM5.Line();


                            tempCurrentEndNode.No = currentNewNodeNo;
                            tempCurrentEndNode.X = Math.Round(endPoint.X, 5);
                            tempCurrentEndNode.Y = Math.Round(endPoint.Y, 5);
                            tempCurrentEndNode.Z = Math.Round(endPoint.Z, 5);

                            RfemNodeList.Add(tempCurrentEndNode);

                            tempCurrentLine.No = currentNewLineNo;
                            tempCurrentLine.Type = LineType.PolylineType;

                            tempCurrentLine.NodeList = $"{RfemNodeList[RfemNodeList.Count - 2].No}, {RfemNodeList[RfemNodeList.Count - 1].No}";
                            RfemLineList.Add(tempCurrentLine);

                            nodeCount++;
                            surfaceNodeCounter++;
                            lineCount++;
                            surfaceLineCounter++;

                            currentNewNodeNo++;
                            currentNewLineNo++;
                        }
                    }

                    #endregion

                    //defines surface data
                    #region Defining data of the surface to be written

                    // start with making a string with "list" of lines forming the surface
                    int surfaceFirstLine = currentNewLineNo - RhSimpleLines.Count;
                    int surfaceLastLine = surfaceFirstLine + RhSimpleLines.Count - 1;
                    string surfaceLineList = "";

                    for (int i = surfaceFirstLine; i < surfaceLastLine; i++)
                    {
                        surfaceLineList = surfaceLineList + i.ToString() + ",";
                    }
                    surfaceLineList = surfaceLineList + surfaceLastLine.ToString();

                    // defining RFEM parameter of surface
                    Dlubal.RFEM5.Surface surfaceData = new Dlubal.RFEM5.Surface();
                    surfaceData.No = currentNewSurfaceNo;
                    surfaceData.MaterialNo = material.No;
                    surfaceData.GeometryType = SurfaceGeometryType.PlaneSurfaceType;
                    surfaceData.BoundaryLineList = surfaceLineList;
                    surfaceData.Comment = srfCommentMethodIn;
                    // if -1 is input as thickness, surface is created as rigid, otherwise it is "standard"
                    if (srfThicknessInMethod == -1)
                    {
                        surfaceData.StiffnessType = SurfaceStiffnessType.RigidStiffnessType;
                    }
                    else if (srfThicknessInMethod == 0)
                    {
                        surfaceData.StiffnessType = SurfaceStiffnessType.NullStiffnessType;
                    }
                    else
                    {
                        surfaceData.StiffnessType = SurfaceStiffnessType.StandardStiffnessType;
                        surfaceData.Thickness.Constant = srfThicknessInMethod;
                    }

                    //adding surface to elements to be written
                    RfemSurfaceList.Add(surfaceData);

                    surfaceCount++;
                    currentNewSurfaceNo++;

                    #endregion

                }
            }

            //try writing the surface;
            #region Writing the surface and releasing RFEM model

            try
            {
                //prepares model for modification.
                data.PrepareModification();
                data.SetMaterial(material);

                //This version writes nodes one-by-one because the data.SetNodes() for array appears not to be working
                foreach (Node currentRfemNode in RfemNodeList)
                {
                    data.SetNode(currentRfemNode);
                }

                //This version writes lines one-by-one because the data.SetLines() for array appears not to be working
                foreach (Dlubal.RFEM5.Line currentRfemLine in RfemLineList)
                {
                    data.SetLine(currentRfemLine);
                }

                //This version writes lines one-by-one because the data.SetLines() for array appears not to be working
                foreach (Dlubal.RFEM5.Surface currentRfemSurface in RfemSurfaceList)
                {
                    data.SetSurface(currentRfemSurface);
                }

                //finishes modifications - regenerates numbering etc.
                data.FinishModification();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error - Surface Write", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //resetting counters;
            nodeCount = 1;
            lineCount = 1;
            surfaceCount = 1;


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

            #endregion



            //output 'success' as true 
            writeSuccess = true;

            ///the Surfaces below outputs created RFEM surfaces in output parameter
            ///current funcionality does not use this
            return RfemSurfaceList;

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
                return GH_RFEM.Properties.Resources.icon_surface;
            }
        }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("ff4168d8-29e0-486d-9949-9b6d3670da0e"); }
        }
    }
}
