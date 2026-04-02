using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;
using Apparel_Dynamic_1._0.Helper;
using Apparel_Dynamic_1._0.Resources;

namespace Apparel_Dynamic_1._0.Modules
{
    class SalesQuotation
    {
        public SalesQuotation()
        {
            Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            Application.SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
        }

        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.FormTypeEx == "149" && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                {
                    Global.G_Form = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction == true)
                    {
                        try
                        {
                            Global.G_Form.Freeze(true);

                            // *** OTT *****

                            //Static Text
                            Global.G_Form.Items.Add("STOTNTRY", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                            Global.oStatic = (SAPbouiCOM.StaticText)Global.G_Form.Items.Item("STOTNTRY").Specific;
                            Global.oStatic.Caption = "OTT";
                            Global.G_Form.Items.Item("STOTNTRY").Top = Global.G_Form.Items.Item("70").Top + 15;
                            Global.G_Form.Items.Item("STOTNTRY").Left = Global.G_Form.Items.Item("70").Left;
                            Global.G_Form.Items.Item("STOTNTRY").Width = Global.G_Form.Items.Item("70").Width;
                            Global.G_Form.Items.Item("STOTNTRY").FromPane = 0;
                            Global.G_Form.Items.Item("STOTNTRY").ToPane = 0;

                         
                            // EditText (ETOTTNTRY)
                            Global.G_Form.Items.Add("ETOTNTRY", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            Global.G_Form.Items.Item("ETOTNTRY").Top = Global.G_Form.Items.Item("STOTNTRY").Top;
                            Global.G_Form.Items.Item("ETOTNTRY").Left = Global.G_Form.Items.Item("14").Left;
                            Global.G_Form.Items.Item("ETOTNTRY").Width = (Global.G_Form.Items.Item("14").Width / 2) - 5;
                            Global.G_Form.Items.Item("ETOTNTRY").FromPane = 0;
                            Global.G_Form.Items.Item("ETOTNTRY").ToPane = 0;
                            Global.G_Form.Items.Item("ETOTNTRY").SetAutoManagedAttribute(
                                SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                            Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETOTNTRY").Specific;
                            Global.oEdit.DataBind.SetBound(true, "OQUT", "U_OTTENTRY");
                            Global.G_Form.Items.Item("STOTNTRY").LinkTo = "ETOTNTRY";

                            // Linked Button
                            Global.G_Form.Items.Add("LNOTNTRY", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            Global.oLinkButton = (SAPbouiCOM.LinkedButton)Global.G_Form.Items.Item("LNOTNTRY").Specific;
                            // Global.oLinkButton.LinkedObjectType = "ESPL_UDO_DD_GATENTRY";
                            Global.G_Form.Items.Item("LNOTNTRY").Top = Global.G_Form.Items.Item("STOTNTRY").Top;
                            Global.G_Form.Items.Item("LNOTNTRY").Left = Global.G_Form.Items.Item("ETOTNTRY").Left - 19;
                            Global.G_Form.Items.Item("LNOTNTRY").Height = Global.G_Form.Items.Item("LNOTNTRY").Height - 1;
                            Global.G_Form.Items.Item("LNOTNTRY").LinkTo = "ETOTNTRY";

                            // EditText (ETOTTNO)
                            Global.G_Form.Items.Add("ETOTTNO", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            Global.G_Form.Items.Item("ETOTTNO").Top = Global.G_Form.Items.Item("STOTNTRY").Top;
                            Global.G_Form.Items.Item("ETOTTNO").Left = Global.G_Form.Items.Item("ETOTNTRY").Left + Global.G_Form.Items.Item("ETOTNTRY").Width + 5;
                            Global.G_Form.Items.Item("ETOTTNO").Width = (Global.G_Form.Items.Item("14").Width / 2) - 5;
                            Global.G_Form.Items.Item("ETOTTNO").FromPane = 0;
                            Global.G_Form.Items.Item("ETOTTNO").ToPane = 0;
                            Global.G_Form.Items.Item("ETOTTNO").Enabled = false;
                            Global.G_Form.Items.Item("ETOTTNO").SetAutoManagedAttribute(
                                SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                            Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETOTTNO").Specific;
                            Global.oEdit.DataBind.SetBound(true, "OQUT", "U_OTTNO");
                            Global.G_Form.Items.Item("ETOTNTRY").LinkTo = "ETOTTNO";

                            //// Choose From List (CFL)
                            //Global.oCfls = Global.G_Form.ChooseFromLists;
                            //Global.oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)
                            //    Global.G_UI_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                            //Global.oCFLCreationParams.MultiSelection = false;
                            //Global.oCFLCreationParams.ObjectType = "FIL_D_OTT";
                            //Global.oCFLCreationParams.UniqueID = "CFL_OTT";

                            //Global.oCfl = Global.oCfls.Add(Global.oCFLCreationParams);

                            //Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETOTNTRY").Specific;
                            //Global.oEdit.ChooseFromListUID = "CFL_OTT";
                            //Global.oEdit.ChooseFromListAlias = "DocEntry";

                            //***** Sales Contract***** 
                            //Static Text
                            Global.G_Form.Items.Add("STSCNTRY", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                            Global.oStatic = (SAPbouiCOM.StaticText)Global.G_Form.Items.Item("STSCNTRY").Specific;
                            Global.oStatic.Caption = "Sales Contract";
                            Global.G_Form.Items.Item("STSCNTRY").Top = Global.G_Form.Items.Item("254000012").Top + 18;
                            Global.G_Form.Items.Item("STSCNTRY").Left = Global.G_Form.Items.Item("254000012").Left;
                            Global.G_Form.Items.Item("STSCNTRY").Width = Global.G_Form.Items.Item("254000012").Width;
                            Global.G_Form.Items.Item("STSCNTRY").FromPane = 0;
                            Global.G_Form.Items.Item("STSCNTRY").ToPane = 0;


                            // EditText (ETOTTNTRY)
                            Global.G_Form.Items.Add("ETSCNTRY", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            Global.G_Form.Items.Item("ETSCNTRY").Top = Global.G_Form.Items.Item("STSCNTRY").Top;
                            Global.G_Form.Items.Item("ETSCNTRY").Left = Global.G_Form.Items.Item("254000015").Left;
                            Global.G_Form.Items.Item("ETSCNTRY").Width = (Global.G_Form.Items.Item("254000015").Width / 2) - 5;
                            Global.G_Form.Items.Item("ETSCNTRY").FromPane = 0;
                            Global.G_Form.Items.Item("ETSCNTRY").ToPane = 0;
                            Global.G_Form.Items.Item("ETSCNTRY").SetAutoManagedAttribute(
                                SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                            Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETSCNTRY").Specific;
                            Global.oEdit.DataBind.SetBound(true, "OQUT", "U_SCNTRY");
                            Global.G_Form.Items.Item("STSCNTRY").LinkTo = "ETSCNTRY";

                            // Linked Button
                            Global.G_Form.Items.Add("LNSCNTRY", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            Global.oLinkButton = (SAPbouiCOM.LinkedButton)Global.G_Form.Items.Item("LNSCNTRY").Specific;
                            // Global.oLinkButton.LinkedObjectType = "ESPL_UDO_DD_GATENTRY";
                            Global.G_Form.Items.Item("LNSCNTRY").Top = Global.G_Form.Items.Item("STSCNTRY").Top;
                            Global.G_Form.Items.Item("LNSCNTRY").Left = Global.G_Form.Items.Item("ETSCNTRY").Left - 19;
                            Global.G_Form.Items.Item("LNSCNTRY").Height = Global.G_Form.Items.Item("LNSCNTRY").Height - 1;
                            Global.G_Form.Items.Item("LNSCNTRY").LinkTo = "ETSCNTRY";

                            // EditText (ETSCNO)
                            Global.G_Form.Items.Add("ETSCNO", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            Global.G_Form.Items.Item("ETSCNO").Top = Global.G_Form.Items.Item("STSCNTRY").Top;
                            Global.G_Form.Items.Item("ETSCNO").Left = Global.G_Form.Items.Item("ETSCNTRY").Left + Global.G_Form.Items.Item("ETSCNTRY").Width + 5;
                            Global.G_Form.Items.Item("ETSCNO").Width = (Global.G_Form.Items.Item("254000015").Width / 2) - 5;
                            Global.G_Form.Items.Item("ETSCNO").FromPane = 0;
                            Global.G_Form.Items.Item("ETSCNO").ToPane = 0;
                            Global.G_Form.Items.Item("ETSCNO").Enabled = false;
                            Global.G_Form.Items.Item("ETSCNO").SetAutoManagedAttribute(
                                SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                            Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETSCNO").Specific;
                            Global.oEdit.DataBind.SetBound(true, "OQUT", "U_SCNO");
                            Global.G_Form.Items.Item("ETSCNTRY").LinkTo = "ETSCNO";

                            //// Choose From List (CFL)
                            //Global.oCfls = Global.G_Form.ChooseFromLists;
                            //Global.oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)
                            //    Global.G_UI_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                            //Global.oCFLCreationParams.MultiSelection = false;
                            //Global.oCFLCreationParams.ObjectType = "FIL_D_OSCM";
                            //Global.oCFLCreationParams.UniqueID = "CFL_SLCN";

                            //Global.oCfl = Global.oCfls.Add(Global.oCFLCreationParams);

                            //Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETSCNTRY").Specific;
                            //Global.oEdit.ChooseFromListUID = "CFL_SLCN";
                            //Global.oEdit.ChooseFromListAlias = "DocEntry";

                            // ************** Style No. **************

                            Global.G_Form.Items.Add("STStyleNo", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                            Global.oStatic = (SAPbouiCOM.StaticText)Global.G_Form.Items.Item("STStyleNo").Specific;
                            Global.oStatic.Caption = "Style No.";

                            Global.G_Form.Items.Item("STStyleNo").Top = Global.G_Form.Items.Item("2002").Top + 18;
                            Global.G_Form.Items.Item("STStyleNo").Left = Global.G_Form.Items.Item("2002").Left;
                            Global.G_Form.Items.Item("STStyleNo").Width = Global.G_Form.Items.Item("2002").Width;
                            Global.G_Form.Items.Item("STStyleNo").Height = Global.G_Form.Items.Item("2002").Height;
                            Global.G_Form.Items.Item("STStyleNo").FromPane = 0;
                            Global.G_Form.Items.Item("STStyleNo").ToPane = 0;

                            //// Linked Button
                            //Global.G_Form.Items.Add("LNStyleNo", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);

                            //// EditText (ETStyleNo)
                            //Global.G_Form.Items.Add("ETStyleNo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            //Global.G_Form.Items.Item("ETStyleNo").Top = Global.G_Form.Items.Item("STStyleNo").Top;
                            //Global.G_Form.Items.Item("ETStyleNo").Left = Global.G_Form.Items.Item("2003").Left;
                            //Global.G_Form.Items.Item("ETStyleNo").Width = (Global.G_Form.Items.Item("2003").Width / 2) - 5;
                            //Global.G_Form.Items.Item("ETStyleNo").Height = Global.G_Form.Items.Item("2003").Height;
                            //Global.G_Form.Items.Item("ETStyleNo").FromPane = 0;
                            //Global.G_Form.Items.Item("ETStyleNo").ToPane = 0;
                            //Global.G_Form.Items.Item("ETStyleNo").SetAutoManagedAttribute(
                            //    SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1,
                            //    SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                            //Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETStyleNo").Specific;
                            //Global.oEdit.DataBind.SetBound(true, "OQUT", "U_StyleNo");
                            //Global.G_Form.Items.Item("STStyleNo").LinkTo = "ETStyleNo";

                            //// Linked Button Setup
                            //Global.oLinkButton = (SAPbouiCOM.LinkedButton)Global.G_Form.Items.Item("LNStyleNo").Specific;
                            //// Global.oLinkButton.LinkedObjectType = "ESPL_UDO_DD_GATENTRY"; // Optional
                            //Global.G_Form.Items.Item("LNStyleNo").Top = Global.G_Form.Items.Item("STStyleNo").Top;
                            //Global.G_Form.Items.Item("LNStyleNo").Left = Global.G_Form.Items.Item("ETStyleNo").Left - 19;
                            //Global.G_Form.Items.Item("LNStyleNo").Height = Global.G_Form.Items.Item("LNStyleNo").Height - 1;
                            //Global.G_Form.Items.Item("LNStyleNo").LinkTo = "ETStyleNo";

                            //// Choose From List (CFL)
                            //Global.oCfls = Global.G_Form.ChooseFromLists;
                            //Global.oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)
                            //    Global.G_UI_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                            //Global.oCFLCreationParams.MultiSelection = false;
                            //Global.oCFLCreationParams.ObjectType = "FIL_M_OPSM";
                            //Global.oCFLCreationParams.UniqueID = "CFL_OPSM";

                            //Global.oCfl = Global.oCfls.Add(Global.oCFLCreationParams);

                            //Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETStyleNo").Specific;
                            //Global.oEdit.ChooseFromListUID = "CFL_OPSM";
                            //Global.oEdit.ChooseFromListAlias = "Code";

                            //// Button (BTStyleNo)
                            //Global.G_Form.Items.Add("BTStyleNo", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            //Global.G_Form.Items.Item("BTStyleNo").Top = Global.G_Form.Items.Item("STStyleNo").Top;
                            //Global.G_Form.Items.Item("BTStyleNo").Left = Global.G_Form.Items.Item("ETStyleNo").Left + Global.G_Form.Items.Item("ETStyleNo").Width + 5;
                            //Global.G_Form.Items.Item("BTStyleNo").Width = (Global.G_Form.Items.Item("2003").Width / 2) - 5;
                            //Global.G_Form.Items.Item("BTStyleNo").Height = Global.G_Form.Items.Item("ETStyleNo").Height;
                            //Global.G_Form.Items.Item("BTStyleNo").FromPane = 0;
                            //Global.G_Form.Items.Item("BTStyleNo").ToPane = 0;

                            //Global.oButton = (SAPbouiCOM.Button)Global.G_Form.Items.Item("BTStyleNo").Specific;
                            //Global.oButton.Caption = "Get Info";

                            //Global.G_Form.Items.Item("BTStyleNo").SetAutoManagedAttribute(
                            //    SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1,
                            //    SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                            //// ************ BUYER'S STYLE NO ************
                            //Global.G_Form.Items.Add("STBSTYLENO", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                            //Global.oStatic = (SAPbouiCOM.StaticText)Global.G_Form.Items.Item("STBSTYLENO").Specific;
                            //Global.oStatic.Caption = "Buyers Style No.";

                            //Global.G_Form.Items.Item("STBSTYLENO").Top = Global.G_Form.Items.Item("STStyleNo").Top + 18; // - 2000 
                            //Global.G_Form.Items.Item("STBSTYLENO").Left = Global.G_Form.Items.Item("STStyleNo").Left;
                            //Global.G_Form.Items.Item("STBSTYLENO").Width = Global.G_Form.Items.Item("STStyleNo").Width;
                            //Global.G_Form.Items.Item("STBSTYLENO").FromPane = 0;
                            //Global.G_Form.Items.Item("STBSTYLENO").ToPane = 0;

                            //// Add EditText for Buyer's Style No
                            //Global.G_Form.Items.Add("ETBSTYLENO", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            //Global.G_Form.Items.Item("ETBSTYLENO").Top = Global.G_Form.Items.Item("STBSTYLENO").Top;
                            //Global.G_Form.Items.Item("ETBSTYLENO").Left = Global.G_Form.Items.Item("ETStyleNo").Left;
                            //Global.G_Form.Items.Item("ETBSTYLENO").Width = Global.G_Form.Items.Item("ETStyleNo").Width;
                            //Global.G_Form.Items.Item("ETBSTYLENO").Height = Global.G_Form.Items.Item("ETStyleNo").Height;
                            //Global.G_Form.Items.Item("ETBSTYLENO").FromPane = 0;
                            //Global.G_Form.Items.Item("ETBSTYLENO").ToPane = 0;
                            //Global.G_Form.Items.Item("ETBSTYLENO").SetAutoManagedAttribute(
                            //    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                            //    1,
                            //    SAPbouiCOM.BoModeVisualBehavior.mvb_False
                            //);
                            //Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETBSTYLENO").Specific;
                            //Global.oEdit.DataBind.SetBound(true, "OQUT", "U_BSTYLENO");
                            //Global.G_Form.Items.Item("STBSTYLENO").LinkTo = "ETBSTYLENO";

                            //// Add EditText for Buyer's Style Name
                            //Global.G_Form.Items.Add("ETBSTYLENM", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            //Global.G_Form.Items.Item("ETBSTYLENM").Top = Global.G_Form.Items.Item("ETBSTYLENO").Top;
                            //Global.G_Form.Items.Item("ETBSTYLENM").Left = Global.G_Form.Items.Item("ETBSTYLENO").Left + Global.G_Form.Items.Item("ETBSTYLENO").Width + 5;
                            //Global.G_Form.Items.Item("ETBSTYLENM").Width = Global.G_Form.Items.Item("BTStyleNo").Width;
                            //Global.G_Form.Items.Item("ETBSTYLENM").Height = Global.G_Form.Items.Item("BTStyleNo").Height;
                            //Global.G_Form.Items.Item("ETBSTYLENM").FromPane = 0;
                            //Global.G_Form.Items.Item("ETBSTYLENM").ToPane = 0;
                            //Global.G_Form.Items.Item("ETBSTYLENM").SetAutoManagedAttribute(
                            //    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                            //    1,
                            //    SAPbouiCOM.BoModeVisualBehavior.mvb_False
                            //);
                            //Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETBSTYLENM").Specific;
                            //Global.oEdit.DataBind.SetBound(true, "OQUT", "U_BSTYLENM");
                            //Global.G_Form.Items.Item("ETBSTYLENO").LinkTo = "ETBSTYLENM";
                            //// ************ BUYER'S STYLE NO ************

                            //// ************* TYPE ********************
                            //Global.G_Form.Items.Add("STTYPE", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                            //Global.oStatic = (SAPbouiCOM.StaticText)Global.G_Form.Items.Item("STTYPE").Specific;
                            //Global.oStatic.Caption = "Order Type.";

                            //Global.G_Form.Items.Item("STTYPE").Top = Global.G_Form.Items.Item("254000012").Top + 18;  // - 2000
                            //Global.G_Form.Items.Item("STTYPE").Left = Global.G_Form.Items.Item("254000012").Left;
                            //Global.G_Form.Items.Item("STTYPE").Width = Global.G_Form.Items.Item("254000012").Width;
                            //Global.G_Form.Items.Item("STTYPE").FromPane = 0;
                            //Global.G_Form.Items.Item("STTYPE").ToPane = 0;

                            //// Add ComboBox for Type
                            //Global.G_Form.Items.Add("CBTYPE", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                            //Global.G_Form.Items.Item("CBTYPE").Top = Global.G_Form.Items.Item("STTYPE").Top;
                            //Global.G_Form.Items.Item("CBTYPE").Left = Global.G_Form.Items.Item("254000015").Left;
                            //Global.G_Form.Items.Item("CBTYPE").Width = Global.G_Form.Items.Item("254000015").Width;
                            //Global.G_Form.Items.Item("CBTYPE").FromPane = 0;
                            //Global.G_Form.Items.Item("CBTYPE").ToPane = 0;
                            //Global.G_Form.Items.Item("CBTYPE").DisplayDesc = true;
                            //Global.G_Form.Items.Item("CBTYPE").SetAutoManagedAttribute(
                            //    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                            //    1,
                            //    SAPbouiCOM.BoModeVisualBehavior.mvb_False
                            //);

                            //Global.oCombo = (SAPbouiCOM.ComboBox)Global.G_Form.Items.Item("CBTYPE").Specific;
                            //Global.oCombo.DataBind.SetBound(true, "OQUT", "U_TYPE");
                            //Global.G_Form.Items.Item("STTYPE").LinkTo = "CBTYPE";
                            //// ************* TYPE ********************
                            //// ********** CREATE FOLDER TAB21 *************
                            //Global.G_Form.Items.Add("TAB21", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                            //SAPbouiCOM.Folder oFolder = (SAPbouiCOM.Folder)Global.G_Form.Items.Item("TAB21").Specific;
                            //oFolder.Caption = "Size and Colour";
                            ////G_Form.DataSources.UserDataSources.Add("FolderDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                            //try
                            //{
                            //    oFolder.DataBind.SetBound(true, "", "FolderDS");
                            //}
                            //catch (Exception)
                            //{
                            //}
                            //oFolder.AutoPaneSelection = false;
                            //oFolder.Pane = 21;
                            //oFolder.GroupWith("112");

                            //Global.G_Form.Items.Item("112").Width = 80;
                            //Global.G_Form.Items.Item("114").Width = 80;
                            //Global.G_Form.Items.Item("138").Width = 80;
                            //Global.G_Form.Items.Item("2013").Width = 150;
                            //Global.G_Form.Items.Item("1320002137").Width = 80;
                            //Global.G_Form.Items.Item("TAB21").Width = 80;

                            //Global.G_Form.Items.Item("TAB21").Left = Global.G_Form.Items.Item("112").Left + Global.G_Form.Items.Item("112").Width;
                            //Global.G_Form.Items.Item("TAB21").Top = Global.G_Form.Items.Item("112").Top;
                            //Global.G_Form.Items.Item("TAB21").Visible = true;

                            //Global.G_Form.Items.Item("114").Left = Global.G_Form.Items.Item("TAB21").Left + Global.G_Form.Items.Item("TAB21").Width;
                            //Global.G_Form.Items.Item("138").Left = Global.G_Form.Items.Item("114").Left + Global.G_Form.Items.Item("114").Width;
                            //Global.G_Form.Items.Item("2013").Left = Global.G_Form.Items.Item("138").Left + Global.G_Form.Items.Item("138").Width;
                            //Global.G_Form.Items.Item("1320002137").Left = Global.G_Form.Items.Item("2013").Left + Global.G_Form.Items.Item("2013").Width;

                            //// ********** GRID **********
                            //Global.G_Form.Items.Add("SCGrid", SAPbouiCOM.BoFormItemTypes.it_GRID);
                            //SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)Global.G_Form.Items.Item("SCGrid").Specific;
                            //Global.G_Form.Items.Item("SCGrid").Width = (Global.G_Form.Items.Item("118").Left - 10) - (Global.G_Form.Items.Item("117").Left + 10);
                            //Global.G_Form.Items.Item("SCGrid").Top = Global.G_Form.Items.Item("44").Top;
                            //Global.G_Form.Items.Item("SCGrid").Left = Global.G_Form.Items.Item("44").Left + 15;
                            //Global.G_Form.Items.Item("SCGrid").Height = (Global.G_Form.Items.Item("116").Top - 10) - Global.G_Form.Items.Item("38").Top;
                            //Global.G_Form.Items.Item("SCGrid").FromPane = 21;
                            //Global.G_Form.Items.Item("SCGrid").ToPane = 21;

                            //// Add DataTable
                            //Global.G_Form.DataSources.DataTables.Add("DT_0");

                            //// ********** Qty Static **********
                            //Global.G_Form.Items.Add("STQty", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                            //SAPbouiCOM.StaticText oStatic = (SAPbouiCOM.StaticText)Global.G_Form.Items.Item("STQty").Specific;
                            //oStatic.Caption = "Qty.";

                            //Global.G_Form.Items.Item("STQty").Top = Global.G_Form.Items.Item("38").Top + Global.G_Form.Items.Item("38").Height + 15;
                            //Global.G_Form.Items.Item("STQty").Left = Global.G_Form.Items.Item("86").Left;
                            //Global.G_Form.Items.Item("STQty").Width = Global.G_Form.Items.Item("86").Width;
                            //Global.G_Form.Items.Item("STQty").FromPane = 1;
                            //Global.G_Form.Items.Item("STQty").ToPane = 1;

                            //// ********** Qty EditText **********
                            //Global.G_Form.Items.Add("ETQty", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            //Global.G_Form.Items.Item("ETQty").Top = Global.G_Form.Items.Item("STQty").Top;
                            //Global.G_Form.Items.Item("ETQty").Left = Global.G_Form.Items.Item("46").Left;
                            //Global.G_Form.Items.Item("ETQty").Width = Global.G_Form.Items.Item("ETStyleNo").Width;
                            //Global.G_Form.Items.Item("ETQty").FromPane = 1;
                            //Global.G_Form.Items.Item("ETQty").ToPane = 1;
                            //Global.G_Form.Items.Item("ETQty").SetAutoManagedAttribute(
                            //    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                            //    1,
                            //    SAPbouiCOM.BoModeVisualBehavior.mvb_False
                            //);
                            //SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETQty").Specific;
                            //oEdit.DataBind.SetBound(true, "OQUT", "U_TotalQty");
                            //Global.G_Form.Items.Item("STQty").LinkTo = "ETQty";

                            //// ********** Get Item Button **********
                            //Global.G_Form.Items.Add("BTQty", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            //Global.G_Form.Items.Item("BTQty").Top = Global.G_Form.Items.Item("ETQty").Top;
                            //Global.G_Form.Items.Item("BTQty").Left = Global.G_Form.Items.Item("ETStyleNo").Left + Global.G_Form.Items.Item("ETStyleNo").Width + 5;
                            //Global.G_Form.Items.Item("BTQty").Width = (Global.G_Form.Items.Item("46").Width / 2) - 5;
                            //Global.G_Form.Items.Item("BTQty").Height = Global.G_Form.Items.Item("ETQty").Height;
                            //Global.G_Form.Items.Item("BTQty").FromPane = 1;
                            //Global.G_Form.Items.Item("BTQty").ToPane = 1;

                            //SAPbouiCOM.Button oButton = (SAPbouiCOM.Button)Global.G_Form.Items.Item("BTQty").Specific;
                            //oButton.Caption = "Get Item";
                            //Global.G_Form.Items.Item("BTQty").SetAutoManagedAttribute(
                            //    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                            //    1,
                            //    SAPbouiCOM.BoModeVisualBehavior.mvb_False
                            //);

                            //// ********** Size Colour ENTRY Hidden EditText **********
                            //Global.G_Form.Items.Add("ETCRSZNTRY", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            //Global.G_Form.Items.Item("ETCRSZNTRY").Top = -2000;
                            //Global.G_Form.Items.Item("ETCRSZNTRY").Enabled = false;
                            //oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETCRSZNTRY").Specific;
                            //oEdit.DataBind.SetBound(true, "OQUT", "U_CRSZNTRY");

                            //// Click to refresh form state
                            //Global.G_Form.Items.Item("4").Click();


                            Global.G_Form.Freeze(false);
                        }
                        catch (Exception ex)
                        {
                            Global.G_Form.Freeze(false);
                        }
                    }
                }

            }catch(Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage("Error in Draft Order for SAP Screen - " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
        }
        private void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

    }
}
