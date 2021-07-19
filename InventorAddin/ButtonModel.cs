using Inventor;
using System;

namespace InvAddIn
{
    public class ButtonModel
    {
        public delegate void ExecuteMethod(NameValueMap vs);
        public string ButtonName { get; set; }
        public string DisplayName { get; set; }
        public string InternalName { get; set; }
        public CommandTypesEnum Clasification { get; set; }
        public string AddInGIUD { get; set; }
        public string DescriptionText { get; set; }
        public string ToolTipText { get; set; }
        public string StandardIcon { get; set; }
        public string LargeIcon { get; set; }
        public ButtonDisplayEnum ButtonDisplay { get; set; }
        public string RibbonName { get; set; }
        public string TabName { get; set; }
        public string PanelName { get; set; }        
        public ExecuteMethod CommandForExecute;
        public ButtonModel(ExecuteMethod commandForExecute)
        {
            CommandForExecute = commandForExecute;
        }
    }
}
