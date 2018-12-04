using System;
using System.Collections.Generic;
using System.Linq;
using Grasshopper.Kernel;
using System.Windows.Forms;
using Grasshopper.Kernel.Types;

namespace BEST
{
    public class BestParameter : GH_PersistentParam<GH_Integer>
    {
        #region constructors
        public BestParameter()
                  : base("Material Catalogue", "Mat Catalogue",
              "Select Material from catalogue",
              "RFEM", "Test")
        { }
        public BestParameter(GH_InstanceDescription componentDescription)
          : base(componentDescription)
        { }
        #endregion

        #region properties
        public static readonly Guid Id = new Guid("{1F421979-9529-49E1-8B9F-5F8285A25F15}");
        public override Guid ComponentGuid
        {
            get { return Id; }
        }
        #endregion

        #region constants
        private static readonly Dictionary<int, string> _cities = new Dictionary<int, string>()
    {
      { 1, "Montgomery, Alabama" },
      { 2, "Birmingham, Alabama" },
      { 3, "Juneau, Alaska" },
      { 4, "Anchorage, Alaska" },
      { 5, "Phoenix, Arizona" },
      { 6, "Tucson, Arizona" },
      { 7, "Little Rock, Arkansas" },
      { 8, "Fort Smith, Arkansas" },
      { 9, "Olympia, Washington" },
      { 10, "Seattle, Washington" },
      { 11, "Madison, Wisconsin" },
      { 12, "Milwaukee, Wisconsin" },
      { 13, "Cheyenne, Wyoming" },
      { 14, "Rock Springs, Wyoming" },
      { 15, "Thermopolis, Wyoming" },
      { 16, "Casper, Wyoming" },
      { 17, "Sheridan, Wyoming" }
    };
        #endregion

        #region menu
        /// <summary>
        /// Checks to see if a specific city code is currently the one and only integer in the volatile data.
        /// </summary>
        /// <param name="cityCode">City code to check.</param>
        /// <returns>True if the city is currently selected.</returns>
        private bool IsCitySelected(int cityCode)
        {
            if (SourceCount > 0) return false;
            if (PersistentData.IsEmpty) return false;
            if (PersistentDataCount != 1) return false;

            GH_Integer integer = PersistentData.get_FirstItem(true);
            return integer?.Value == cityCode;
        }

        protected override ToolStripMenuItem Menu_CustomSingleValueItem()
        {
            ToolStripMenuItem root = new ToolStripMenuItem("Pick State City");
            if (SourceCount > 0)
            {
                root.Enabled = false;
                return root;
            }

            ToolStripMenuItem[] stateItems = {
        new ToolStripMenuItem("Alabama"),
        new ToolStripMenuItem("Alaska"),
        new ToolStripMenuItem("Arizona"),
        new ToolStripMenuItem("Arkansas"),
        new ToolStripMenuItem("Washington"),
        new ToolStripMenuItem("Wisconsin"),
        new ToolStripMenuItem("Wyoming")
      };

            foreach (KeyValuePair<int, string> pair in _cities)
                InjectCity(stateItems, pair.Key, pair.Value);

            foreach (ToolStripMenuItem item in stateItems)
                root.DropDownItems.Add(item);

            return root;
        }
        protected override ToolStripMenuItem Menu_CustomMultiValueItem()
        {
            return null;
        }

        private bool InjectCity(IEnumerable<ToolStripMenuItem> items, int code, string name)
        {
            int comma = name.IndexOf(",", StringComparison.Ordinal);
            if (comma < 0)
                throw new ArgumentException("Name must contain a comma with a state name.");

            string state = name.Substring(comma + 1).Trim();
            name = name.Substring(0, comma);

            foreach (ToolStripMenuItem item in items)
                if (item.Text.Equals(state, StringComparison.OrdinalIgnoreCase))
                {
                    ToolStripMenuItem cityItem = new ToolStripMenuItem(name);
                    cityItem.Tag = code;
                    cityItem.Checked = IsCitySelected(code);
                    cityItem.ToolTipText = string.Format("{0}, {1}", name, state);
                    cityItem.Click += CityItemOnClick;

                    item.DropDownItems.Add(cityItem);
                    return true;
                }
            return false;
        }
        private void CityItemOnClick(object sender, EventArgs eventArgs)
        {
            ToolStripMenuItem item = sender as ToolStripMenuItem;
            if (item == null)
                return;

            int code = (int)item.Tag;
            if (IsCitySelected(code))
                return;

            RecordPersistentDataEvent("Set city: " + item.Name);
            PersistentData.Clear();
            PersistentData.Append(new GH_Integer(code));
            ExpireSolution(true);
        }
        #endregion

        protected override GH_GetterResult Prompt_Singular(ref GH_Integer value)
        {
            return GH_GetterResult.cancel;
        }
        protected override GH_GetterResult Prompt_Plural(ref List<GH_Integer> values)
        {
            return GH_GetterResult.cancel;
        }
    }
}
