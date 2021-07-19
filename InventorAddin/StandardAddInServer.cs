using InvAddIn;
using Inventor;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Interop;
using WpfApp1;

namespace InventorAddin
{
    enum InventorRibbonsEnum { ZeroDoc, Part, Assembly, Drawing, Presentation, iFeatures }
    enum InventorZeroDocTabsEnum { Get_Started, Tools, Add_ins, Inventor_Debug, Autodesk_A360 }
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// </summary>
    [Guid("f4f20b5e-8c55-4f8c-80ff-9ec8e7c6d656")]
    public class StandardAddInServer : ApplicationAddInServer
    {

        // Inventor application object.
        private Inventor.Application m_inventorApplication;
        private ApplicationEvents m_appEvents;
        private readonly string addinGuid = "f4f20b5e-8c55-4f8c-80ff-9ec8e7c6d656";

        public StandardAddInServer()
        {
        }

        #region ApplicationAddInServer Members

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            // This method is called by Inventor when it loads the addin.
            // The AddInSiteObject provides access to the Inventor Application object.
            // The FirstTime flag indicates if the addin is loaded for the first time.

            // Initialize AddIn members.
            m_inventorApplication = addInSiteObject.Application;
            ButtonModel button = new ButtonModel(ButtonDef1_OnExecute)
            {
                AddInGIUD = addinGuid,
                ButtonDisplay = ButtonDisplayEnum.kDisplayTextInLearningMode,
                ButtonName = "ButtonDef1",
                Clasification = CommandTypesEnum.kEditMaskCmdType,
                DescriptionText = "Test how WPF works with Inventor",
                DisplayName = "Open\nWPF",
                InternalName = "Inventor:TestWPF:ButtonDef1",
                LargeIcon = "CommIcon_48x48",
                PanelName = "Testing",
                RibbonName = "Assembly",
                StandardIcon = "CommIcon_16x16",
                TabName = "id_TabAssemble",
                ToolTipText = "Open new WPF in inventor"                
            };
            ButtonDefinition _buttonDef1 = CreateButton(button, firstTime);

            m_appEvents = m_inventorApplication.ApplicationEvents;
            m_appEvents.OnActivateDocument += new ApplicationEventsSink_OnActivateDocumentEventHandler(ApplicationEvents_OnActivateDocument);
            m_appEvents.OnDocumentChange += new ApplicationEventsSink_OnDocumentChangeEventHandler(ApplicationEvents_OnDocumentChange);
        }
        private ButtonDefinition CreateButton(ButtonModel button, bool firstTime)
        {
            ButtonDefinition output;
            ControlDefinitions oCtrlDef = m_inventorApplication.CommandManager.ControlDefinitions;

            object[] oIconPictures = GetIcons("CommIcon_16x16", "CommIcon_48x48");
            try
            {
                output = oCtrlDef["Inventor:" + button.PanelName + ":" + button.ButtonName] as ButtonDefinition;
            }
            catch (Exception)
            {
                output = oCtrlDef.AddButtonDefinition(button.DisplayName
                    , "Inventor:" + button.PanelName + ":" + button.ButtonName
                    , button.Clasification
                    , button.AddInGIUD
                    , button.DescriptionText
                    , button.ToolTipText
                    , oIconPictures[0]
                    , oIconPictures[1]
                    , button.ButtonDisplay);

                CommandCategory cmdCat = m_inventorApplication.CommandManager.CommandCategories.Add("Tester", "Inventor:" + button.PanelName + ":", button.AddInGIUD);

                cmdCat.Add(output);
            }
            if (firstTime)
            {
                try
                {
                    if (m_inventorApplication.UserInterfaceManager.InterfaceStyle == InterfaceStyleEnum.kRibbonInterface)
                    {
                        Ribbon ribbon = m_inventorApplication.UserInterfaceManager.Ribbons[button.RibbonName];
                        RibbonTab tab = ribbon.RibbonTabs[button.TabName];
                        try
                        {
                            RibbonPanel panel = tab.RibbonPanels.Add(button.PanelName, "Inventor:" + button.PanelName + ":Panel", button.AddInGIUD, "", false);
                            CommandControl control1 = panel.CommandControls.AddButton(output, true, true, "", false);
                        }
                        catch (Exception)
                        {

                        }
                    }
                    else
                    {
                        CommandBar oCommandBar = m_inventorApplication.UserInterfaceManager.CommandBars["PMxPartFeatureCmdBar"];
                        oCommandBar.Controls.AddButton(output, 0);
                    }
                }
                catch
                {
                    CommandBar oCommandBar = m_inventorApplication.UserInterfaceManager.CommandBars["PMxPartFeatureCmdBar"];
                    oCommandBar.Controls.AddButton(output, 0);
                }
            }
            //output.OnExecute += new ButtonDefinitionSink_OnExecuteEventHandler(button.OnClick);
            output.OnExecute += button.CommandForExecute;
            return output;
        }
        private object[] GetIcons(string smallIcon, string largeIcon)
        {
            object[] output = new object[2];
            Assembly assembly = Assembly.GetExecutingAssembly();
            string[] resources = assembly.GetManifestResourceNames();

            for (int i = 0; i < resources.Length; i++)
            {
                if (resources[i].Contains(smallIcon))
                {
                    Stream oSmallIconStream = assembly.GetManifestResourceStream(resources[1]);
                    Bitmap oSmallIcon = new Bitmap(oSmallIconStream);
                    output[0] = AxHostConverter.ImageToPictureDisp(oSmallIcon);
                }
                if (resources[i].Contains(largeIcon))
                {
                    Stream oLargeIconStream = assembly.GetManifestResourceStream(resources[2]);
                    Bitmap oLargeIcon = new Bitmap(oLargeIconStream);
                    output[1] = AxHostConverter.ImageToPictureDisp(oLargeIcon);
                }
            }
            return output;
        }
        void ButtonDef1_OnExecute(NameValueMap Context)
        {
            Commands.Button1Run(m_inventorApplication);
        }
        private void ApplicationEvents_OnActivateDocument(_Document DocumentObject, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
            if (BeforeOrAfter != EventTimingEnum.kAfter)
            {
                return;
            }
            HandlingCode = HandlingCodeEnum.kEventHandled;
            // TODO - Document is acivated. Make some action if it is necessary.
        }
        private void ApplicationEvents_OnDocumentChange(_Document document, EventTimingEnum BeforeOrAfter, CommandTypesEnum reasonForChange, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
            if (BeforeOrAfter != EventTimingEnum.kAfter)
            {
                return;
            }
            HandlingCode = HandlingCodeEnum.kEventHandled;
            if (reasonForChange.HasFlag(CommandTypesEnum.kShapeEditCmdType))
            {
                // TODO - respond on used commands that can affect the geometry of the model.
            }
            if (reasonForChange.HasFlag(CommandTypesEnum.kQueryOnlyCmdType))
            {
                // TODO - respond on used commands that purely query data. These do not 'dirty' the document.
            }
            if (reasonForChange.HasFlag(CommandTypesEnum.kFileOperationsCmdType))
            {
                // TODO - respond on used commands that manage file operations - e.g. File Save.
            }
            if (reasonForChange.HasFlag(CommandTypesEnum.kFilePropertyEditCmdType))
            {
                // TODO - respond on used commands that edit File Properties (a.k.a Document Properties).
            }
            if (reasonForChange.HasFlag(CommandTypesEnum.kUpdateWithReferencesCmdType))
            {
                // TODO - respond on used commands that cause this document to recalculate its contents with respect to changes that may have occurred in files it is referencing.
            }
            if (reasonForChange.HasFlag(CommandTypesEnum.kNonShapeEditCmdType))
            {
                // TODO - respond on used commands that edit data (other than File Properties) that is not directly related to the geometry of the model (e.g. color, style).
            }
            if (reasonForChange.HasFlag(CommandTypesEnum.kReferencesChangeCmdType))
            {
                // TODO - respond on used commands that cause this document to change which files it references.
            }
            if (reasonForChange.HasFlag(CommandTypesEnum.kSchemaChangeCmdType))
            {
                // TODO - respond on used commands that change the format of the data, but do not change it otherwise (e.g. from the format of one Inventor release to another).
            }
            if (reasonForChange.HasFlag(CommandTypesEnum.kEditMaskCmdType))
            {
                // TODO - respond on used commands that cause the document to become 'dirty'. Includes ShapeEdit, FilePropertyEdit, NonShapeEdit and UpdateWithReferences commands.
            }
        }
        public void Deactivate()
        {
            // This method is called by Inventor when the AddIn is unloaded.
            // The AddIn will be unloaded either manually by the user or
            // when the Inventor session is terminated

            // Release objects.
            m_inventorApplication = null;

            m_appEvents.OnActivateDocument -= new ApplicationEventsSink_OnActivateDocumentEventHandler(ApplicationEvents_OnActivateDocument);
            m_appEvents.OnDocumentChange -= new ApplicationEventsSink_OnDocumentChangeEventHandler(ApplicationEvents_OnDocumentChange);
            m_appEvents = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        public void ExecuteCommand(int commandID)
        {
            // Note:this method is now obsolete, you should use the 
            // ControlDefinition functionality for implementing commands.
        }
        public object Automation
        {
            // This property is provided to allow the AddIn to expose an API 
            // of its own to other programs. Typically, this  would be done by
            // implementing the AddIn's API interface in a class and returning 
            // that class object through this property.

            get
            {
                // TODO: Add ApplicationAddInServer.Automation getter implementation
                return null;
            }
        }
        #endregion
        internal class AxHostConverter : AxHost
        {
            private AxHostConverter() : base("")
            {
            }
            public static stdole.IPictureDisp ImageToPictureDisp(Image image)
            {
                return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
            }
        }
        internal class Commands
        {
            public static void Button1Run(Inventor.Application inventorApplication)
            {
                MainWindow mainWindow = new MainWindow();

                _ = new WindowInteropHelper(mainWindow)
                {
                    Owner = new IntPtr(inventorApplication.MainFrameHWND)
                };

                mainWindow.ShowDialog();
            }
        }
    }
}
