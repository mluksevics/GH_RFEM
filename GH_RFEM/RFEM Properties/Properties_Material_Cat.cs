using System;
using System.Collections.Generic;
using System.Linq;
using Grasshopper.Kernel;
using System.Windows.Forms;
using Grasshopper.Kernel.Types;

namespace GH_RFEM
{
    public class MaterialCatalogue : GH_PersistentParam<GH_String>
    {
        #region constructors
        public MaterialCatalogue()
                  : base("Material Catalogue", "Mat Catalogue",
              "Select Material from catalogue",
              "RFEM", "Material")
        { }
        public MaterialCatalogue(GH_InstanceDescription componentDescription)
          : base(componentDescription)
        { }
        #endregion

        #region properties
        public static readonly Guid Id = new Guid("{1e791e58-d0b8-4bfb-945b-e0b0f55c723c}");
        public override Guid ComponentGuid
        {
            get { return Id; }
        }
        #endregion

        #region constants
        private static readonly Dictionary<string, string> _materials = new Dictionary<string, string>()
    {
      { "NameID|Beton C12/15@TypeID|CONCRETE@NormID|EN 1992-1-1", "C12/15, Concrete (EN1992)" },
      { "NameID|Beton C16/20@TypeID|CONCRETE@NormID|EN 1992-1-1", "C16/20, Concrete (EN1992)" },
      { "NameID|Beton C20/25@TypeID|CONCRETE@NormID|EN 1992-1-1", "C20/25, Concrete (EN1992)" },
      { "NameID|Beton C25/30@TypeID|CONCRETE@NormID|EN 1992-1-1", "C25/30, Concrete (EN1992)" },
      { "NameID|Beton C28/35@TypeID|CONCRETE@NormID|BS EN 1992-1-1", "C28/35, Concrete (EN1992)" },
      { "NameID|Beton C30/37@TypeID|CONCRETE@NormID|EN 1992-1-1", "C30/37, Concrete (EN1992)" },
      { "NameID|Beton C32/40@TypeID|CONCRETE@NormID|BS EN 1992-1-1", "C32/40, Concrete (EN1992)" },
      { "NameID|Beton C35/45@TypeID|CONCRETE@NormID|EN 1992-1-1", "C35/45, Concrete (EN1992)" },
      { "NameID|Beton C40/50@TypeID|CONCRETE@NormID|EN 1992-1-1", "C40/50, Concrete (EN1992)" },
      { "NameID|Beton C45/55@TypeID|CONCRETE@NormID|EN 1992-1-1", "C45/55, Concrete (EN1992)" },
      { "NameID|Beton C50/60@TypeID|CONCRETE@NormID|EN 1992-1-1", "C50/60, Concrete (EN1992)" },
      { "NameID|Beton C55/67@TypeID|CONCRETE@NormID|EN 1992-1-1", "C55/67, Concrete (EN1992)" },
      { "NameID|Beton C60/75@TypeID|CONCRETE@NormID|EN 1992-1-1", "C60/75, Concrete (EN1992)" },
      { "NameID|Beton C70/85@TypeID|CONCRETE@NormID|EN 1992-1-1", "C70/85, Concrete (EN1992)" },
      { "NameID|Baustahl S 235@TypeID|STEEL@NormID|EN 1993-1-1", "S235, Steel (EN1993)" },
      { "NameID|Baustahl S 275@TypeID|STEEL@NormID|EN 1993-1-1", "S275, Steel (EN1993)" },
      { "NameID|Baustahl S 355@TypeID|STEEL@NormID|EN 1993-1-1", "S355, Steel (EN1993)" },
      { "NameID|Baustahl S 450@TypeID|STEEL@NormID|EN 1993-1-1", "S355, Steel (EN1993)" },
      { "NameID|Pappel und Nadelholz C14@TypeID|CONIFEROUS@NormID|EN 1995-1-1", "C14, Timber (EN1995)" },
      { "NameID|Pappel und Nadelholz C16@TypeID|CONIFEROUS@NormID|EN 1995-1-1", "C16, Timber (EN1995)" },
      { "NameID|Pappel und Nadelholz C18@TypeID|CONIFEROUS@NormID|EN 1995-1-1", "C18, Timber (EN1995)" },
      { "NameID|Pappel und Nadelholz C20@TypeID|CONIFEROUS@NormID|EN 1995-1-1", "C20, Timber (EN1995)" },
      { "NameID|Pappel und Nadelholz C22@TypeID|CONIFEROUS@NormID|EN 1995-1-1", "C22, Timber (EN1995)" },
      { "NameID|Pappel und Nadelholz C24@TypeID|CONIFEROUS@NormID|EN 1995-1-1", "C24, Timber (EN1995)" },
      { "NameID|Pappel und Nadelholz C27@TypeID|CONIFEROUS@NormID|EN 1995-1-1", "C27, Timber (EN1995)" },
      { "NameID|Pappel und Nadelholz C30@TypeID|CONIFEROUS@NormID|EN 1995-1-1", "C30, Timber (EN1995)" },
      { "NameID|Pappel und Nadelholz C35@TypeID|CONIFEROUS@NormID|EN 1995-1-1", "C35, Timber (EN1995)" },
      { "NameID|Pappel und Nadelholz C40@TypeID|CONIFEROUS@NormID|EN 1995-1-1", "C40, Timber (EN1995)" },
      { "NameID|Pappel und Nadelholz C45@TypeID|CONIFEROUS@NormID|EN 1995-1-1", "C45, Timber (EN1995)" },
      { "NameID|Pappel und Nadelholz C50@TypeID|CONIFEROUS@NormID|EN 1995-1-1", "C50, Timber (EN1995)" },
      { "NameID|Laubholz D18@TypeID|LEAFY@NormID|EN 1995-1-1", "D18, Timber (EN1995)" },
      { "NameID|Laubholz D24@TypeID|LEAFY@NormID|EN 1995-1-1", "D24, Timber (EN1995)" },
      { "NameID|Laubholz D30@TypeID|LEAFY@NormID|EN 1995-1-1", "D30, Timber (EN1995)" },
      { "NameID|Laubholz D35@TypeID|LEAFY@NormID|EN 1995-1-1", "D35, Timber (EN1995)" },
      { "NameID|Laubholz D40@TypeID|LEAFY@NormID|EN 1995-1-1", "D40, Timber (EN1995)" },
      { "NameID|Laubholz D50@TypeID|LEAFY@NormID|EN 1995-1-1", "D50, Timber (EN1995)" },
      { "NameID|Laubholz D60@TypeID|LEAFY@NormID|EN 1995-1-1", "D60, Timber (EN1995)" },
      { "NameID|Laubholz D70@TypeID|LEAFY@NormID|EN 1995-1-1", "D70, Timber (EN1995)" },
    };
        #endregion

        #region menu
        /// <summary>
        /// Checks to see if a specific city code is currently the one and only integer in the volatile data.
        /// </summary>
        /// <param name="materialCodeInput">Material code to check.</param>
        /// <returns>True if the material is currently selected.</returns>
        private bool IsMaterialSelected(string materialCodeInput)
        {
            if (SourceCount > 0) return false;
            if (PersistentData.IsEmpty) return false;
            if (PersistentDataCount != 1) return false;

            GH_String materialCode = PersistentData.get_FirstItem(true);
            return materialCode?.Value == materialCodeInput;
        }

        protected override ToolStripMenuItem Menu_CustomSingleValueItem()
        {
            ToolStripMenuItem root = new ToolStripMenuItem("Pick Material");
            if (SourceCount > 0)
            {
                root.Enabled = false;
                return root;
            }

            ToolStripMenuItem[] materialTypeItems = {
        new ToolStripMenuItem("Concrete (EN1992)"),
        new ToolStripMenuItem("Steel (EN1993)"),
        new ToolStripMenuItem("Timber (EN1995)"),
      };

            foreach (KeyValuePair<string, string> pair in _materials)
                InjectMaterial(materialTypeItems, pair.Key, pair.Value);

            foreach (ToolStripMenuItem item in materialTypeItems)
                root.DropDownItems.Add(item);

            return root;
        }
        protected override ToolStripMenuItem Menu_CustomMultiValueItem()
        {
            return null;
        }

        private bool InjectMaterial(IEnumerable<ToolStripMenuItem> items, string code, string name)
        {
            int comma = name.IndexOf(",", StringComparison.Ordinal);
            if (comma < 0)
                throw new ArgumentException("Name must contain a comma with a material type.");

            string state = name.Substring(comma + 1).Trim();
            name = name.Substring(0, comma);

            foreach (ToolStripMenuItem item in items)
                if (item.Text.Equals(state, StringComparison.OrdinalIgnoreCase))
                {
                    ToolStripMenuItem materialItem = new ToolStripMenuItem(name);
                    materialItem.Tag = code;
                    materialItem.Checked = IsMaterialSelected(code);
                    materialItem.ToolTipText = string.Format("{0}, {1}", name, state);
                    materialItem.Click += MaterialItemOnClick;

                    item.DropDownItems.Add(materialItem);
                    return true;
                }
            return false;
        }
        private void MaterialItemOnClick(object sender, EventArgs eventArgs)
        {
            ToolStripMenuItem item = sender as ToolStripMenuItem;
            if (item == null)
                return;

            string code = (string)item.Tag;
            if (IsMaterialSelected(code))
                return;

            RecordPersistentDataEvent("Set material: " + item.Name);
            PersistentData.Clear();
            PersistentData.Append(new GH_String(code));
            ExpireSolution(true);
        }
        #endregion

        protected override GH_GetterResult Prompt_Singular(ref GH_String value)
        {
            return GH_GetterResult.cancel;
        }
        protected override GH_GetterResult Prompt_Plural(ref List<GH_String> values)
        {
            return GH_GetterResult.cancel;
        }
    }
}
