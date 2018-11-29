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
    public class RFEM_Surface : GH_Component
    {
        //definition of RFEM and input variables used in further methods
        IApplication app = null;
        IModel model = null;
        bool run = false;
        double segmentLength = 1;

        //cycling through all Surfaces and creating RFEM objects;
        int nodeCount = 1;
        int lineCount = 1;
        int surfaceCount = 1;

        //defining variables used in cycles
        //defining variables needed to store geometry and RFEM info
        Rhino.Geometry.Point3d startPoint;
        Rhino.Geometry.Point3d endPoint;

        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public RFEM_Surface()
          : base("Surface RFEM", "RFEM Srf",
              "Create RFEM Surfaces from Rhino Surfaces",
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
            pManager.AddSurfaceParameter("Surface", "Surfaces", "Input Rhino planar sufaces you want to create as RFEM Surfaces", GH_ParamAccess.list);
            pManager.AddNumberParameter("Segment length", "MaxSegmentLength[m]", "Any edges with splines/circles/arcs will be simplified as segments with maximum length described in this parameter", GH_ParamAccess.item, 1);
            pManager.AddBooleanParameter("Run", "Toggle", "Toggles whether the Surfaces are written to RFEM", GH_ParamAccess.item, false);

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

            //pManager.AddGenericParameter("RFEM Surfaces", "RfLn", "RFEM Surfaces for export in RFEM", GH_ParamAccess.list);

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
            List<Rhino.Geometry.Brep> rhino_surface = new List<Rhino.Geometry.Brep>();
            List<Dlubal.RFEM5.Surface> RfemSurfaces = new List<Dlubal.RFEM5.Surface>();

            // Then we need to access the input parameters individually. 
            // When data cannot be extracted from a parameter, we should abort this method.
            if (!DA.GetDataList<Rhino.Geometry.Brep>(0, rhino_surface)) return;
            DA.GetData(1, ref segmentLength);
            DA.GetData(2, ref run);

            // The actual functionality will be in a method defined below. This is where we run it
            if (run == true)
            {
                RfemSurfaces = CreateRfemSurfaces(rhino_surface);
            }

            // Finally assign the processed data to the output parameter.
            //DA.SetData(0, RfemNodes);
        }

        private List<Dlubal.RFEM5.Surface> CreateRfemSurfaces(List<Rhino.Geometry.Brep> Rh_Srf)
        {
            //start by reducing the input curves to simple Surfaces with start/end points
            List<Rhino.Geometry.Curve> RhSimpleLines = new List<Rhino.Geometry.Curve>();

            // Gets interface to running RFEM application.
            app = Marshal.GetActiveObject("RFEM5.Application") as IApplication;
            // Locks RFEM licence
            app.LockLicense();

            // Gets interface to active RFEM model.
            model = app.GetActiveModel();

            // Gets interface to model data.
            IModelData data = model.GetModelData();

            //prepares model for modification.
            data.PrepareModification();

            // Creates material used for all surfaces
            Material material = new Material
            {
                No = 1,
                TextID = "NameID|Beton C30/37@TypeID|CONCRETE@NormID|DIN 1045-1 - 08"
            };

            Dlubal.RFEM5.Surface[] RfemSurfaceArray = new Dlubal.RFEM5.Surface[Rh_Srf.Count];

            foreach (Rhino.Geometry.Brep RhSingleSurface in Rh_Srf)
            {

                Rhino.Geometry.Curve[] curves = RhSingleSurface.DuplicateEdgeCurves(true);
                //double tol = doc.ModelAbsoluteTolerance * 2.1;
                curves = Rhino.Geometry.Curve.JoinCurves(curves);

                foreach(Rhino.Geometry.Curve RhSingleCurve in curves)
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

                    Dlubal.RFEM5.Node[] RfemNodeArray = new Dlubal.RFEM5.Node[RhSimpleLines.Count * 2];
                    Dlubal.RFEM5.Line[] RfemLineArray = new Dlubal.RFEM5.Line[RhSimpleLines.Count];

                    int surfaceNodeCounter = 0; //counts nodes witin one surface. nodeCount counts overall nodes in model
                    int surfaceLineCounter = 0; // counts lines (edges) for one surface, lineCount counts overall lines in model

                    for (int i = 0; i < RhSimpleLines.Count; i++)
                    {
                        startPoint = RhSimpleLines[i].PointAtStart;
                        endPoint = RhSimpleLines[i].PointAtEnd;

                        //if this is the first line for the surface
                        if (i==0)
                        {
                            RfemNodeArray[surfaceNodeCounter].No = nodeCount;
                            RfemNodeArray[surfaceNodeCounter].X = Math.Round(startPoint.X, 5);
                            RfemNodeArray[surfaceNodeCounter].Y = Math.Round(startPoint.Y, 5);
                            RfemNodeArray[surfaceNodeCounter].Z = Math.Round(startPoint.Z, 5);

                            RfemNodeArray[surfaceNodeCounter + 1].No = nodeCount + 1;
                            RfemNodeArray[surfaceNodeCounter + 1].X = Math.Round(endPoint.X, 5);
                            RfemNodeArray[surfaceNodeCounter + 1].Y = Math.Round(endPoint.Y, 5);
                            RfemNodeArray[surfaceNodeCounter + 1].Z = Math.Round(endPoint.Z, 5);

                            data.SetNode(RfemNodeArray[surfaceNodeCounter]);
                            data.SetNode(RfemNodeArray[surfaceNodeCounter + 1]);

                            RfemLineArray[surfaceLineCounter].No = lineCount;
                            RfemLineArray[surfaceLineCounter].Type = LineType.PolylineType;

                            RfemLineArray[surfaceLineCounter].NodeList = $"{RfemNodeArray[surfaceNodeCounter].No}, {RfemNodeArray[surfaceNodeCounter + 1].No}";
                            data.SetLine(RfemLineArray[surfaceLineCounter]);

                            nodeCount = nodeCount + 2;
                            surfaceNodeCounter = surfaceNodeCounter + 2;
                            lineCount++;
                            surfaceLineCounter++;

                        }
                        //if this is the last node for the surface
                        else if (i== RhSimpleLines.Count-1)
                        {
                            //no need to define new node as these are both already defined
                            //create line connecting previous node with first node for surface
                            RfemLineArray[surfaceLineCounter].NodeList = $"{RfemNodeArray[surfaceNodeCounter - 1].No}, {RfemNodeArray[0].No}";
                            data.SetLine(RfemLineArray[surfaceLineCounter]);
                            lineCount++;
                            surfaceLineCounter++;

                        }
                        else
                        {
                            //if this is just a node somewhere on edges
                            RfemNodeArray[surfaceNodeCounter].No = nodeCount;
                            RfemNodeArray[surfaceNodeCounter].X = Math.Round(endPoint.X, 5);
                            RfemNodeArray[surfaceNodeCounter].Y = Math.Round(endPoint.Y, 5);
                            RfemNodeArray[surfaceNodeCounter].Z = Math.Round(endPoint.Z, 5);

                            data.SetNode(RfemNodeArray[surfaceNodeCounter]);

                            RfemLineArray[surfaceLineCounter].No = lineCount;
                            RfemLineArray[surfaceLineCounter].Type = LineType.PolylineType;

                            RfemLineArray[surfaceLineCounter].NodeList = $"{RfemNodeArray[surfaceNodeCounter-1].No}, {RfemNodeArray[surfaceNodeCounter].No}";
                            data.SetLine(RfemLineArray[surfaceLineCounter]);

                            nodeCount++;
                            surfaceNodeCounter++;
                            lineCount++;
                            surfaceLineCounter++;

                        }

                        /*
                        try
                        {

                            
                            // finish modification - RFEM regenerates the data
                            //data.FinishModification();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                       */

                    }

                    int surfaceFirstLine = lineCount - RhSimpleLines.Count;
                    int surfaceLastLine = surfaceFirstLine + RhSimpleLines.Count - 1;
                    string surfaceLineList = "";

                    for (int i = surfaceFirstLine; i < surfaceLastLine; i++)
                    {
                        surfaceLineList = surfaceLineList + i.ToString() + ",";
                    }
                    surfaceLineList = surfaceLineList + surfaceLastLine.ToString();

                    Dlubal.RFEM5.Surface surfaceData = new Dlubal.RFEM5.Surface
                    {
                        No = surfaceCount,
                        MaterialNo = material.No,
                        GeometryType = SurfaceGeometryType.PlaneSurfaceType,
                        BoundaryLineList = surfaceLineList,
                        StiffnessType = SurfaceStiffnessType.StandardStiffnessType,
                    };
                    surfaceData.Thickness.Type = SurfaceThicknessType.ConstantThicknessType;
                    surfaceData.Thickness.Constant = 0.2;

                    try
                    {
                        // modification
                        // Sets all objects to model data.
                        //data.PrepareModification();
                        ISurface surface = data.SetSurface(surfaceData);
                        // finish modification - RFEM regenerates the data


                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }


                    RfemSurfaceArray[surfaceCount - 1] = surfaceData;

                    surfaceCount++;

                    //clear lines used in surface
                    RhSimpleLines.Clear();
                }


            }

            //resetting counters;
            nodeCount = 1;
            lineCount = 1;
            surfaceCount = 1;

            //finishes modifications - regenerates numbering etc.
            data.FinishModification();

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

            ///the Surfaces below outputs created RFEM nodes in output parameter
            ///current funcionality does not use this
            ///it uses a custom class (written within this project) RfemNodeType to wrap the Dlubal.RFEM5.Node objects.

            List<Dlubal.RFEM5.Surface> RfemSurfaceList = RfemSurfaceArray.OfType<Dlubal.RFEM5.Surface>().ToList(); // this isn't going to be fast.

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
