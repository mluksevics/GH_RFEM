using System;
using System.Collections.Generic;
using System.Linq;
using Grasshopper.Kernel;
using System.Windows.Forms;
using Grasshopper.Kernel.Types;

namespace GH_RFEM
{
    public class SectionCatalogue : GH_PersistentParam<GH_String>
    {
        #region constructors
        public SectionCatalogue()
                  : base("Section Catalogue", "Sect Catalogue",
              "Select Cross section from catalogue",
              "RFEM", "Section")
        { }
        public SectionCatalogue(GH_InstanceDescription componentDescription)
          : base(componentDescription)
        { }
        #endregion

        #region properties
        public static readonly Guid Id = new Guid("{72b990d8-507e-4cd9-8550-95a082e19e06}");
        public override Guid ComponentGuid
        {
            get { return Id; }
        }
        #endregion

        #region constants
        private static readonly Dictionary<string, string> _sections = new Dictionary<string, string>()
    {
      { "UB 203x133x25 (Corus)", "UB 203x133x25, UB" },
      { "HE-A 220 (ArcelorMittal)", "HE-A 220, HEA" },
    };
        #endregion

        #region menu
        /// <summary>
        /// Checks to see if a specific city code is currently the one and only integer in the volatile data.
        /// </summary>
        /// <param name="sectionCodeInput">Sectopm code to check.</param>
        /// <returns>True if the section is currently selected.</returns>
        private bool IsSectionSelected(string sectionCodeInput)
        {
            if (SourceCount > 0) return false;
            if (PersistentData.IsEmpty) return false;
            if (PersistentDataCount != 1) return false;

            GH_String sectionCode = PersistentData.get_FirstItem(true);
            return sectionCode?.Value == sectionCodeInput;
        }

        protected override ToolStripMenuItem Menu_CustomSingleValueItem()
        {
            ToolStripMenuItem root = new ToolStripMenuItem("Pick Section");
            if (SourceCount > 0)
            {
                root.Enabled = false;
                return root;
            }

            ToolStripMenuItem[] sectionTypeItems = {
        new ToolStripMenuItem("UB"),
        new ToolStripMenuItem("HEA"),
      };

            foreach (KeyValuePair<string, string> pair in _sections)
                InjectSection(sectionTypeItems, pair.Key, pair.Value);

            foreach (ToolStripMenuItem item in sectionTypeItems)
                root.DropDownItems.Add(item);

            return root;
        }
        protected override ToolStripMenuItem Menu_CustomMultiValueItem()
        {
            return null;
        }

        private bool InjectSection(IEnumerable<ToolStripMenuItem> items, string code, string name)
        {
            int comma = name.IndexOf(",", StringComparison.Ordinal);
            if (comma < 0)
                throw new ArgumentException("Name must contain a comma with a section type.");

            string state = name.Substring(comma + 1).Trim();
            name = name.Substring(0, comma);

            foreach (ToolStripMenuItem item in items)
                if (item.Text.Equals(state, StringComparison.OrdinalIgnoreCase))
                {
                    ToolStripMenuItem sectionItem = new ToolStripMenuItem(name);
                    sectionItem.Tag = code;
                    sectionItem.Checked = IsSectionSelected(code);
                    sectionItem.ToolTipText = string.Format("{0}, {1}", name, state);
                    sectionItem.Click += SectionItemOnClick;

                    item.DropDownItems.Add(sectionItem);
                    return true;
                }
            return false;
        }
        private void SectionItemOnClick(object sender, EventArgs eventArgs)
        {
            ToolStripMenuItem item = sender as ToolStripMenuItem;
            if (item == null)
                return;

            string code = (string)item.Tag;
            if (IsSectionSelected(code))
                return;

            RecordPersistentDataEvent("Set section: " + item.Name);
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
