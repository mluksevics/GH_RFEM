using System;
using System.Drawing;
using Grasshopper.Kernel;

namespace GH_RFEM
{
    public class GH_RFEMInfo : GH_AssemblyInfo
    {
        public override string Name
        {
            get
            {
                return "GHRFEM";
            }
        }
        public override Bitmap Icon
        {
            get
            {
                //Return a 24x24 pixel bitmap to represent this GHA library.
                return null;
            }
        }
        public override string Description
        {
            get
            {
                //Return a short string describing the purpose of this GHA library.
                return "";
            }
        }
        public override Guid Id
        {
            get
            {
                return new Guid("6d83ec07-d979-4099-9e36-d45fe088f19b");
            }
        }

        public override string AuthorName
        {
            get
            {
                //Return a string identifying you or your company.
                return "";
            }
        }
        public override string AuthorContact
        {
            get
            {
                //Return a string representing your preferred contact details.
                return "";
            }
        }
    }
}
