using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;

using Dlubal.RFEM5;

namespace GH_RFEM
{

    public class RfemNodeType : GH_Goo<Dlubal.RFEM5.Node>
    {
        public RfemNodeType()
        {
            this.Value = new Dlubal.RFEM5.Node();
        }
        public RfemNodeType(Dlubal.RFEM5.Node n)
        {
            n = new Node();
            this.Value = n;
        }


        public override bool IsValid
        {
            get
            {
                return Value.IsValid;
            }
        }

        public override string TypeDescription
        {
            get
            {
                return "Container for RFEM Nodes";
            }
        }

        public override string TypeName
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public override IGH_Goo Duplicate()
        {
            throw new NotImplementedException();
        }

        public override string ToString()
        {
            return Value.No.ToString();
        }
    }

}
