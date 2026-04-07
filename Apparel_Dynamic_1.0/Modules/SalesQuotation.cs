using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;
using Apparel_Dynamic_1._0.Helper;
using Apparel_Dynamic_1._0.Resources.Master;

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
                            Global.G_Form.Items.Add("STOTTNO", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                            Global.oStatic = (SAPbouiCOM.StaticText)Global.G_Form.Items.Item("STOTTNO").Specific;
                            Global.oStatic.Caption = "OTT";
                            Global.G_Form.Items.Item("STOTTNO").Top = Global.G_Form.Items.Item("70").Top + 15;
                            Global.G_Form.Items.Item("STOTTNO").Left = Global.G_Form.Items.Item("70").Left;
                            Global.G_Form.Items.Item("STOTTNO").Width = Global.G_Form.Items.Item("70").Width;
                            Global.G_Form.Items.Item("STOTTNO").FromPane = 0;
                            Global.G_Form.Items.Item("STOTTNO").ToPane = 0;

                         
                            // EditText (ETOTTNTRY)
                            Global.G_Form.Items.Add("ETOTTNO", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            Global.G_Form.Items.Item("ETOTTNO").Top = Global.G_Form.Items.Item("STOTTNO").Top;
                            Global.G_Form.Items.Item("ETOTTNO").Left = Global.G_Form.Items.Item("14").Left;
                            Global.G_Form.Items.Item("ETOTTNO").Width = (Global.G_Form.Items.Item("14").Width / 2) - 5;
                            Global.G_Form.Items.Item("ETOTTNO").FromPane = 0;
                            Global.G_Form.Items.Item("ETOTTNO").ToPane = 0;
                            Global.G_Form.Items.Item("ETOTTNO").SetAutoManagedAttribute(
                                SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                            Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETOTTNO").Specific;
                            Global.oEdit.DataBind.SetBound(true, "OQUT", "U_OTTNO");
                            Global.G_Form.Items.Item("STOTTNO").LinkTo = "ETOTTNO";

                            // Linked Button
                            Global.G_Form.Items.Add("LNOTNTRY", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            Global.oLinkButton = (SAPbouiCOM.LinkedButton)Global.G_Form.Items.Item("LNOTNTRY").Specific;
                            // Global.oLinkButton.LinkedObjectType = "ESPL_UDO_DD_GATENTRY";
                            Global.G_Form.Items.Item("LNOTNTRY").Top = Global.G_Form.Items.Item("STOTTNO").Top;
                            Global.G_Form.Items.Item("LNOTNTRY").Left = Global.G_Form.Items.Item("ETOTTNO").Left - 19;
                            Global.G_Form.Items.Item("LNOTNTRY").Height = Global.G_Form.Items.Item("LNOTNTRY").Height - 1;
                            Global.G_Form.Items.Item("LNOTNTRY").LinkTo = "ETOTTNO";

                                                                                    // EditText (ETOTNTRY)

                            Global.G_Form.Items.Add("ETOTNTRY", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            Global.G_Form.Items.Item("ETOTNTRY").Top = Global.G_Form.Items.Item("STOTTNO").Top;
                            Global.G_Form.Items.Item("ETOTNTRY").Left = Global.G_Form.Items.Item("ETOTTNO").Left + Global.G_Form.Items.Item("ETOTTNO").Width + 2;
                            Global.G_Form.Items.Item("ETOTNTRY").Width = (Global.G_Form.Items.Item("14").Width / 2) + 5 ;
                            Global.G_Form.Items.Item("ETOTNTRY").FromPane = 0;
                            Global.G_Form.Items.Item("ETOTNTRY").ToPane = 0;
                            Global.G_Form.Items.Item("ETOTNTRY").Enabled = false;
                            Global.G_Form.Items.Item("ETOTNTRY").SetAutoManagedAttribute(
                                SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                            Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETOTNTRY").Specific;
                            Global.oEdit.DataBind.SetBound(true, "OQUT", "U_OTTENTRY");
                            Global.G_Form.Items.Item("ETOTTNO").LinkTo = "ETOTNTRY";

                            // Choose From List (CFL)
                            Global.oCfls = Global.G_Form.ChooseFromLists;
                            Global.oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)
                                Global.G_UI_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                            Global.oCFLCreationParams.MultiSelection = false;
                            Global.oCFLCreationParams.ObjectType = "FIL_D_OTT";
                            Global.oCFLCreationParams.UniqueID = "CFL_OTT";

                            Global.oCfl = Global.oCfls.Add(Global.oCFLCreationParams);

                            Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETOTTNO").Specific;
                            Global.oEdit.ChooseFromListUID = "CFL_OTT";
                            Global.oEdit.ChooseFromListAlias = "DocEntry";

                                                                            //***** Sales Contract***** 
                            //Static Text
                            Global.G_Form.Items.Add("STSCNO", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                            Global.oStatic = (SAPbouiCOM.StaticText)Global.G_Form.Items.Item("STSCNO").Specific;
                            Global.oStatic.Caption = "Sales Contract";
                            Global.G_Form.Items.Item("STSCNO").Top = Global.G_Form.Items.Item("254000012").Top + 18;
                            Global.G_Form.Items.Item("STSCNO").Left = Global.G_Form.Items.Item("254000012").Left;
                            Global.G_Form.Items.Item("STSCNO").Width = Global.G_Form.Items.Item("254000012").Width;
                            Global.G_Form.Items.Item("STSCNO").FromPane = 0;
                            Global.G_Form.Items.Item("STSCNO").ToPane = 0;


                                                                                // EditText (ETSCNO)

                            Global.G_Form.Items.Add("ETSCNO", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            Global.G_Form.Items.Item("ETSCNO").Top = Global.G_Form.Items.Item("STSCNO").Top;
                            Global.G_Form.Items.Item("ETSCNO").Left = Global.G_Form.Items.Item("254000015").Left;
                            Global.G_Form.Items.Item("ETSCNO").Width = (Global.G_Form.Items.Item("254000015").Width / 2) - 5;
                            Global.G_Form.Items.Item("ETSCNO").FromPane = 0;
                            Global.G_Form.Items.Item("ETSCNO").ToPane = 0;
                            Global.G_Form.Items.Item("ETSCNO").SetAutoManagedAttribute(
                                SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                            Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETSCNO").Specific;
                            Global.oEdit.DataBind.SetBound(true, "OQUT", "U_SCNO");
                            Global.G_Form.Items.Item("STSCNO").LinkTo = "ETSCNO";

                            // Linked Button
                            Global.G_Form.Items.Add("LNSCNTRY", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            Global.oLinkButton = (SAPbouiCOM.LinkedButton)Global.G_Form.Items.Item("LNSCNTRY").Specific;
                            // Global.oLinkButton.LinkedObjectType = "ESPL_UDO_DD_GATENTRY";
                            Global.G_Form.Items.Item("LNSCNTRY").Top = Global.G_Form.Items.Item("STSCNO").Top;
                            Global.G_Form.Items.Item("LNSCNTRY").Left = Global.G_Form.Items.Item("ETSCNO").Left - 19;
                            Global.G_Form.Items.Item("LNSCNTRY").Height = Global.G_Form.Items.Item("LNSCNTRY").Height - 1;
                            Global.G_Form.Items.Item("LNSCNTRY").LinkTo = "ETSCNTRY";

                                                                                // EditText (ETSCNTRY)

                            Global.G_Form.Items.Add("ETSCNTRY", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            Global.G_Form.Items.Item("ETSCNTRY").Top = Global.G_Form.Items.Item("STSCNO").Top;
                            Global.G_Form.Items.Item("ETSCNTRY").Left = Global.G_Form.Items.Item("ETSCNO").Left + Global.G_Form.Items.Item("ETSCNO").Width + 2;
                            Global.G_Form.Items.Item("ETSCNTRY").Width = (Global.G_Form.Items.Item("254000015").Width / 2) + 5;
                            Global.G_Form.Items.Item("ETSCNTRY").FromPane = 0;
                            Global.G_Form.Items.Item("ETSCNTRY").ToPane = 0;
                            Global.G_Form.Items.Item("ETSCNTRY").Enabled = false;
                            Global.G_Form.Items.Item("ETSCNTRY").SetAutoManagedAttribute(
                                SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                            Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETSCNTRY").Specific;
                            Global.oEdit.DataBind.SetBound(true, "OQUT", "U_SCNTRY");
                            Global.G_Form.Items.Item("ETSCNO").LinkTo = "ETSCNTRY";

                            // Choose From List (CFL)
                            Global.oCfls = Global.G_Form.ChooseFromLists;
                            Global.oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)
                                Global.G_UI_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                            Global.oCFLCreationParams.MultiSelection = false;
                            Global.oCFLCreationParams.ObjectType = "FIL_D_OSCM";
                            Global.oCFLCreationParams.UniqueID = "CFL_SLCN";

                            Global.oCfl = Global.oCfls.Add(Global.oCFLCreationParams);

                            Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETSCNO").Specific;
                            Global.oEdit.ChooseFromListUID = "CFL_SLCN";
                            Global.oEdit.ChooseFromListAlias = "U_SCNO";

                                                                // ************** Style No. **************

                            Global.G_Form.Items.Add("STSTYLNO", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                            Global.oStatic = (SAPbouiCOM.StaticText)Global.G_Form.Items.Item("STSTYLNO").Specific;
                            Global.oStatic.Caption = "Style No.";

                            Global.G_Form.Items.Item("STSTYLNO").Top = Global.G_Form.Items.Item("2002").Top + 18;
                            Global.G_Form.Items.Item("STSTYLNO").Left = Global.G_Form.Items.Item("2002").Left;
                            Global.G_Form.Items.Item("STSTYLNO").Width = Global.G_Form.Items.Item("2002").Width;
                            Global.G_Form.Items.Item("STSTYLNO").Height = Global.G_Form.Items.Item("2002").Height;
                            Global.G_Form.Items.Item("STSTYLNO").FromPane = 0;
                            Global.G_Form.Items.Item("STSTYLNO").ToPane = 0;

                            //// Linked Button
                            Global.G_Form.Items.Add("LNSTYLNO", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);

                            // EditText (ETStyleNo)
                            Global.G_Form.Items.Add("ETSTYLNO", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            Global.G_Form.Items.Item("ETSTYLNO").Top = Global.G_Form.Items.Item("STSTYLNO").Top;
                            Global.G_Form.Items.Item("ETSTYLNO").Left = Global.G_Form.Items.Item("2003").Left;
                            Global.G_Form.Items.Item("ETSTYLNO").Width = (Global.G_Form.Items.Item("2003").Width / 2) - 5;
                            Global.G_Form.Items.Item("ETSTYLNO").Height = Global.G_Form.Items.Item("2003").Height;
                            Global.G_Form.Items.Item("ETSTYLNO").FromPane = 0;
                            Global.G_Form.Items.Item("ETSTYLNO").ToPane = 0;
                            Global.G_Form.Items.Item("ETSTYLNO").SetAutoManagedAttribute(
                                SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1,
                                SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                            Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETSTYLNO").Specific;
                            Global.oEdit.DataBind.SetBound(true, "OQUT", "U_STYLECODE");
                            Global.G_Form.Items.Item("STSTYLNO").LinkTo = "ETSTYLNO";

                            // Linked Button Setup
                            Global.oLinkButton = (SAPbouiCOM.LinkedButton)Global.G_Form.Items.Item("LNSTYLNO").Specific;
                            // Global.oLinkButton.LinkedObjectType = "ESPL_UDO_DD_GATENTRY"; // Optional
                            Global.G_Form.Items.Item("LNSTYLNO").Top = Global.G_Form.Items.Item("STSTYLNO").Top;
                            Global.G_Form.Items.Item("LNSTYLNO").Left = Global.G_Form.Items.Item("ETSTYLNO").Left - 19;
                            Global.G_Form.Items.Item("LNSTYLNO").Height = Global.G_Form.Items.Item("LNSTYLNO").Height - 1;
                            Global.G_Form.Items.Item("LNSTYLNO").LinkTo = "ETSTYLNO";

                            // Choose From List (CFL)
                            Global.oCfls = Global.G_Form.ChooseFromLists;
                            Global.oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)
                                Global.G_UI_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                            Global.oCFLCreationParams.MultiSelection = false;
                            Global.oCFLCreationParams.ObjectType = "FIL_D_OPSM";
                            Global.oCFLCreationParams.UniqueID = "CFL_OPSM";

                            Global.oCfl = Global.oCfls.Add(Global.oCFLCreationParams);

                            Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETSTYLNO").Specific;
                            Global.oEdit.ChooseFromListUID = "CFL_OPSM";
                            Global.oEdit.ChooseFromListAlias = "U_STYLECODE";

                                                                            // Button (BTNSTYLD)
                            Global.G_Form.Items.Add("BTNSTYLD", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            Global.G_Form.Items.Item("BTNSTYLD").Top = Global.G_Form.Items.Item("STSTYLNO").Top;
                            Global.G_Form.Items.Item("BTNSTYLD").Left = Global.G_Form.Items.Item("ETSTYLNO").Left + Global.G_Form.Items.Item("ETSTYLNO").Width + 5;
                            Global.G_Form.Items.Item("BTNSTYLD").Width = (Global.G_Form.Items.Item("2003").Width / 2) - 5;
                            Global.G_Form.Items.Item("BTNSTYLD").Height = Global.G_Form.Items.Item("ETSTYLNO").Height;
                            Global.G_Form.Items.Item("BTNSTYLD").FromPane = 0;
                            Global.G_Form.Items.Item("BTNSTYLD").ToPane = 0;

                            Global.oButton = (SAPbouiCOM.Button)Global.G_Form.Items.Item("BTNSTYLD").Specific;
                            Global.oButton.Caption = "Load Info";

                            Global.G_Form.Items.Item("BTNSTYLD").SetAutoManagedAttribute(
                                SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1,
                                SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                            //// ************ STYLE DESC ************
                            ///
                            Global.G_Form.Items.Add("STSTYLDS", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                            Global.oStatic = (SAPbouiCOM.StaticText)Global.G_Form.Items.Item("STSTYLDS").Specific;
                            Global.oStatic.Caption = "Style Desc.";

                            Global.G_Form.Items.Item("STSTYLDS").Top = Global.G_Form.Items.Item("STSTYLNO").Top + 18; // - 2000 
                            Global.G_Form.Items.Item("STSTYLDS").Left = Global.G_Form.Items.Item("STSTYLNO").Left;
                            Global.G_Form.Items.Item("STSTYLDS").Width = Global.G_Form.Items.Item("STSTYLNO").Width;
                            Global.G_Form.Items.Item("STSTYLDS").FromPane = 0;
                            Global.G_Form.Items.Item("STSTYLDS").ToPane = 0;
                            

                            // Add EditText for Buyer's Style No
                            Global.G_Form.Items.Add("ETSTYLDS", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            Global.G_Form.Items.Item("ETSTYLDS").Top = Global.G_Form.Items.Item("STSTYLDS").Top;
                            Global.G_Form.Items.Item("ETSTYLDS").Left = Global.G_Form.Items.Item("STSTYLNO").Left + Global.G_Form.Items.Item("STSTYLDS").Width;
                            Global.G_Form.Items.Item("ETSTYLDS").Width = Global.G_Form.Items.Item("ETSTYLNO").Width;
                            Global.G_Form.Items.Item("ETSTYLDS").Height = Global.G_Form.Items.Item("STSTYLNO").Height;
                            Global.G_Form.Items.Item("ETSTYLDS").FromPane = 0;
                            Global.G_Form.Items.Item("ETSTYLDS").ToPane = 0;
                            Global.G_Form.Items.Item("ETSTYLDS").Enabled = false;
                            Global.G_Form.Items.Item("ETSTYLDS").SetAutoManagedAttribute(
                                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                                1,
                                SAPbouiCOM.BoModeVisualBehavior.mvb_False
                            );
                            Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETSTYLDS").Specific;
                            Global.oEdit.DataBind.SetBound(true, "OQUT", "U_STYLENM");
                            Global.G_Form.Items.Item("STSTYLDS").LinkTo = "ETSTYLDS";

                            // Add EditText for Style docentry
                            Global.G_Form.Items.Add("ETSLNTRY", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            Global.G_Form.Items.Item("ETSLNTRY").Top = Global.G_Form.Items.Item("ETSTYLDS").Top;
                            Global.G_Form.Items.Item("ETSLNTRY").Left = Global.G_Form.Items.Item("ETSTYLDS").Left + Global.G_Form.Items.Item("ETSTYLDS").Width + 5;
                            Global.G_Form.Items.Item("ETSLNTRY").Width = Global.G_Form.Items.Item("BTNSTYLD").Width;
                            Global.G_Form.Items.Item("ETSLNTRY").Height = Global.G_Form.Items.Item("BTNSTYLD").Height;
                            Global.G_Form.Items.Item("ETSLNTRY").FromPane = 0;
                            Global.G_Form.Items.Item("ETSLNTRY").ToPane = 0;
                            Global.G_Form.Items.Item("ETSLNTRY").Enabled = false;
                            Global.G_Form.Items.Item("ETSLNTRY").SetAutoManagedAttribute(
                                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                                1,
                                SAPbouiCOM.BoModeVisualBehavior.mvb_False
                            );
                            Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETSLNTRY").Specific;
                            Global.oEdit.DataBind.SetBound(true, "OQUT", "U_STYLENTRY");
                            Global.G_Form.Items.Item("ETSTYLDS").LinkTo = "ETSLNTRY";
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
                            //// ************* TYPE ********************///
                            ///

                            // ********** CREATE FOLDER TAB21 *************
                            Global.G_Form.Items.Add("TAB21", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                            SAPbouiCOM.Folder oFolder = (SAPbouiCOM.Folder)Global.G_Form.Items.Item("TAB21").Specific;
                            oFolder.Caption = "Size and Colour";
                            //G_Form.DataSources.UserDataSources.Add("FolderDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                            try
                            {
                                oFolder.DataBind.SetBound(true, "", "FolderDS");
                            }
                            catch (Exception)
                            {
                            }
                            oFolder.AutoPaneSelection = false;
                            oFolder.Pane = 21;
                            oFolder.GroupWith("112");

                            Global.G_Form.Items.Item("112").Width = 80;
                            Global.G_Form.Items.Item("114").Width = 80;
                            Global.G_Form.Items.Item("138").Width = 80;
                            Global.G_Form.Items.Item("2013").Width = 150;
                            Global.G_Form.Items.Item("1320002137").Width = 80;
                            Global.G_Form.Items.Item("TAB21").Width = 80;

                            Global.G_Form.Items.Item("TAB21").Left = Global.G_Form.Items.Item("112").Left + Global.G_Form.Items.Item("112").Width;
                            Global.G_Form.Items.Item("TAB21").Top = Global.G_Form.Items.Item("112").Top;
                            Global.G_Form.Items.Item("TAB21").Visible = true;

                            Global.G_Form.Items.Item("114").Left = Global.G_Form.Items.Item("TAB21").Left + Global.G_Form.Items.Item("TAB21").Width;
                            Global.G_Form.Items.Item("138").Left = Global.G_Form.Items.Item("114").Left + Global.G_Form.Items.Item("114").Width;
                            Global.G_Form.Items.Item("2013").Left = Global.G_Form.Items.Item("138").Left + Global.G_Form.Items.Item("138").Width;
                            Global.G_Form.Items.Item("1320002137").Left = Global.G_Form.Items.Item("2013").Left + Global.G_Form.Items.Item("2013").Width;

                            // ********** GRID **********
                            Global.G_Form.Items.Add("SCGrid", SAPbouiCOM.BoFormItemTypes.it_GRID);
                            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)Global.G_Form.Items.Item("SCGrid").Specific;
                            Global.G_Form.Items.Item("SCGrid").Width = (Global.G_Form.Items.Item("118").Left - 10) - (Global.G_Form.Items.Item("117").Left + 10);
                            Global.G_Form.Items.Item("SCGrid").Top = Global.G_Form.Items.Item("44").Top;
                            Global.G_Form.Items.Item("SCGrid").Left = Global.G_Form.Items.Item("44").Left + 15;
                            Global.G_Form.Items.Item("SCGrid").Height = (Global.G_Form.Items.Item("116").Top - 10) - Global.G_Form.Items.Item("38").Top;
                            Global.G_Form.Items.Item("SCGrid").FromPane = 21;
                            Global.G_Form.Items.Item("SCGrid").ToPane = 21;

                            // Add DataTable
                            Global.G_Form.DataSources.DataTables.Add("DT_0");

                            // ********** Qty Static **********
                            Global.G_Form.Items.Add("STQty", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                            SAPbouiCOM.StaticText oStatic = (SAPbouiCOM.StaticText)Global.G_Form.Items.Item("STQty").Specific;
                            oStatic.Caption = "Qty.";

                            Global.G_Form.Items.Item("STQty").Top = Global.G_Form.Items.Item("38").Top + Global.G_Form.Items.Item("38").Height + 15;
                            Global.G_Form.Items.Item("STQty").Left = Global.G_Form.Items.Item("86").Left;
                            Global.G_Form.Items.Item("STQty").Width = Global.G_Form.Items.Item("86").Width;
                            Global.G_Form.Items.Item("STQty").FromPane =1;
                            Global.G_Form.Items.Item("STQty").ToPane = 1;

                            // ********** Qty EditText **********
                            Global.G_Form.Items.Add("ETQty", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            Global.G_Form.Items.Item("ETQty").Top = Global.G_Form.Items.Item("STQty").Top;
                            Global.G_Form.Items.Item("ETQty").Left = Global.G_Form.Items.Item("46").Left;
                            Global.G_Form.Items.Item("ETQty").Width = Global.G_Form.Items.Item("ETSTYLNO").Width;
                            Global.G_Form.Items.Item("ETQty").FromPane = 1;
                            Global.G_Form.Items.Item("ETQty").ToPane = 1;
                            Global.G_Form.Items.Item("ETQty").Enabled = false;
                            Global.G_Form.Items.Item("ETQty").SetAutoManagedAttribute(
                                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                                1,
                                SAPbouiCOM.BoModeVisualBehavior.mvb_False
                            );
                            SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETQty").Specific;
                            oEdit.DataBind.SetBound(true, "OQUT", "U_TOTALQTY");
                            Global.G_Form.Items.Item("STQty").LinkTo = "ETQty";

                            // ********** Get Item Button **********
                            Global.G_Form.Items.Add("BTQty", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            Global.G_Form.Items.Item("BTQty").Top = Global.G_Form.Items.Item("ETQty").Top;
                            Global.G_Form.Items.Item("BTQty").Left = Global.G_Form.Items.Item("ETSTYLNO").Left + Global.G_Form.Items.Item("ETSTYLNO").Width + 5;
                            Global.G_Form.Items.Item("BTQty").Width = (Global.G_Form.Items.Item("46").Width / 2) - 5;
                            Global.G_Form.Items.Item("BTQty").Height = Global.G_Form.Items.Item("ETQty").Height;
                            Global.G_Form.Items.Item("BTQty").FromPane = 1;
                            Global.G_Form.Items.Item("BTQty").ToPane = 1;

                            SAPbouiCOM.Button oButton = (SAPbouiCOM.Button)Global.G_Form.Items.Item("BTQty").Specific;
                            oButton.Caption = "Get Item";
                            Global.G_Form.Items.Item("BTQty").SetAutoManagedAttribute(
                                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                                1,
                                SAPbouiCOM.BoModeVisualBehavior.mvb_False
                            );

                            // ********** Size Colour ENTRY Hidden EditText **********
                            Global.G_Form.Items.Add("ETCRSZNTRY", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            Global.G_Form.Items.Item("ETCRSZNTRY").Top = -2000;
                            Global.G_Form.Items.Item("ETCRSZNTRY").Enabled = false;
                            oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETCRSZNTRY").Specific;
                            oEdit.DataBind.SetBound(true, "OQUT", "U_CRSZNTRY");

                            // Click to refresh form state
                            Global.G_Form.Items.Item("4").Click();


                            Global.G_Form.Freeze(false);
                        }
                        catch (Exception ex)
                        {
                            Global.G_Form.Freeze(false);
                        }

                    }
                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction == false)
                    {
                        if (!string.IsNullOrWhiteSpace(Global.G_Form.DataSources.DBDataSources.Item("OQUT").GetValue("U_CRSZNTRY", 0).ToString().Trim()))
                        {
                            LoadGrid(ref Global.G_Form, true);
                        }

                    }
                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && pVal.BeforeAction == false)
                    {
                        try
                        {
                            Global.G_Form.Freeze(true);

                            // Adjust Grid Size
                            Global.G_Form.Items.Item("SCGrid").Width =
                                    (Global.G_Form.Items.Item("118").Left - 10) - (Global.G_Form.Items.Item("117").Left + 10);
                            Global.G_Form.Items.Item("SCGrid").Height =
                                    (Global.G_Form.Items.Item("116").Top - 10) - Global.G_Form.Items.Item("38").Top;

                            // FILENTRY Row
                            //oForm.Items.Item("STFILENTRY").Top = oForm.Items.Item("86").Top + 18;
                            //oForm.Items.Item("ETFILENTRY").Top = oForm.Items.Item("STFILENTRY").Top;
                            //oForm.Items.Item("ETFILENO").Top = oForm.Items.Item("STFILENTRY").Top;
                            //oForm.Items.Item("LNFILENTRY").Top = oForm.Items.Item("STFILENTRY").Top;
                            //oForm.Items.Item("LNFILENTRY").Left = oForm.Items.Item("2003").Left - 20;

                            // Second Row
                            //Global.G_Form.Items.Item("2002").Top = Global.G_Form.Items.Item("STFILENTRY").Top + 18;
                            //Global.G_Form.Items.Item("2003").Top = Global.G_Form.Items.Item("2002").Top;

                            // Style No Row
                            Global.G_Form.Items.Item("STSTYLNO").Top = Global.G_Form.Items.Item("2002").Top + 18;
                            Global.G_Form.Items.Item("ETSTYLNO").Top = Global.G_Form.Items.Item("STSTYLNO").Top;
                            Global.G_Form.Items.Item("BTNSTYLD").Top = Global.G_Form.Items.Item("STSTYLNO").Top;
                            Global.G_Form.Items.Item("LNSTYLNO").Top = Global.G_Form.Items.Item("STSTYLNO").Top;

                            //// Buyer Style No Row
                            //Global.G_Form.Items.Item("STBSTYLENO").Top = Global.G_Form.Items.Item("STStyleNo").Top + 18;
                            //Global.G_Form.Items.Item("ETBSTYLENO").Top = Global.G_Form.Items.Item("STBSTYLENO").Top;
                            //Global.G_Form.Items.Item("ETBSTYLENM").Top = Global.G_Form.Items.Item("STBSTYLENO").Top;
                        }
                        catch (Exception ex)
                        {
                            Application.SBO_Application.StatusBar.SetText("Resize error: " + ex.Message,
                                 SAPbouiCOM.BoMessageTime.bmt_Short,
                                 SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        finally
                        {
                            Global.G_Form.Freeze(false);
                        }
                    }
                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK
                           && pVal.ItemUID == "TAB21" && pVal.BeforeAction == false)
                    {
                        Global.G_Form.PaneLevel = 21;

                    }
                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK
                           && pVal.ItemUID == "LNSTYLNO" && pVal.BeforeAction == true)
                    {
                        Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETSTYLNO").Specific;
                        openStyleMaster(Global.oEdit.Value);
                        Global.G_Form = Application.SBO_Application.Forms.Item(pVal.FormUID);
                        BubbleEvent = false;
                        return;
                    }
                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK
                            && pVal.ItemUID == "LNOTNTRY" && pVal.BeforeAction == true)
                    {
                        Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETOTNTRY").Specific;
                        //openOTT(Global.oEdit.Value);
                        Global.G_Form = Application.SBO_Application.Forms.Item(pVal.FormUID);
                        BubbleEvent = false;
                        return;
                    }
                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK
                            && pVal.ItemUID == "LNSCNTRY" && pVal.BeforeAction == true)
                    {
                        Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETSCNTRY").Specific;
                        //openSalesContract(Global.oEdit.Value);
                        Global.G_Form = Application.SBO_Application.Forms.Item(pVal.FormUID);
                        BubbleEvent = false;
                        return;
                    }
                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE
                            && pVal.ItemUID == "SCGrid"   && pVal.BeforeAction == false)
                    {
                        try
                        {
                            Global.G_Form.Freeze(true);
                            if (pVal.Row < 0 || string.IsNullOrWhiteSpace(pVal.ColUID))
                                return;

                            Global.oGrid = (SAPbouiCOM.Grid)Global.G_Form.Items.Item("SCGrid").Specific;
                            Global.oDataTable = Global.G_Form.DataSources.DataTables.Item("DT_0");

                            int dtRow = Global.oGrid.GetDataTableRowIndex(pVal.Row);
                            if (dtRow < 0)
                                return;

                            if (pVal.ColUID == "Total" || pVal.ColUID == "Colour Name" || pVal.ColUID == "Colour Code")
                                return;

                            double total = 0;

                            for (int j = 0; j < Global.oDataTable.Columns.Count; j++)
                            {
                                string currentCol = Global.oDataTable.Columns.Item(j).Name;

                                if (currentCol == "Colour Name" || currentCol == "Colour Code" || currentCol == "Total")
                                    continue;

                                object val = Global.oDataTable.Columns.Item(j).Cells.Item(dtRow).Value;

                                double cellValue = 0;
                                if (val != null && !string.IsNullOrWhiteSpace(val.ToString()))
                                    double.TryParse(val.ToString(), out cellValue);

                                total += cellValue;
                            }

                            Global.oDataTable.Columns.Item("Total").Cells.Item(dtRow).Value = total.ToString();
                            Global.G_Form.Freeze(false);
                        }
                        catch (Exception ex)
                        {
                            Global.G_Form.Freeze(false);

                            Application.SBO_Application.StatusBar.SetText(
                                "Error in SCGrid Validate: " + ex.Message,
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                    }

                    
                    //else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && pVal.ItemUID == "ETOTTNTRY" && pVal.BeforeAction == true)
                    //{
                    //    SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                    //    SAPbouiCOM.DataTable oDataTable;
                    //    oDataTable = oCFLEvento.SelectedObjects;
                    //    SAPbouiCOM.ChooseFromList oCfl = Global.G_Form.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                    //    SAPbouiCOM.Conditions oCons = null;
                    //    SAPbouiCOM.Condition oCon = null;

                    //    oCons = new SAPbouiCOM.Conditions();

                    //    string qStr = @"SELECT DISTINCT A.""DocEntry"" 
                    //FROM ""@FIL_DR_PLANOTT"" A 
                    //WHERE  A.""U_STYLENO"" = '" + Global.G_Form.DataSources.DBDataSources.Item("OQUT").GetValue("U_StyleNo", 0).ToString().TrimEnd() + @"' 
                    //  AND A.""U_BUYRSTCD"" = '" + Global.G_Form.DataSources.DBDataSources.Item("OQUT").GetValue("U_BSTYLENO", 0).ToString().TrimEnd() + @"'";

                    //    SAPbobsCOM.Recordset POSSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //    POSSet.DoQuery(qStr);

                    //    if (POSSet.RecordCount > 0)
                    //    {
                    //        while (!POSSet.EoF)
                    //        {
                    //            oCon = oCons.Add();
                    //            oCon.Alias = "DocEntry";
                    //            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    //            oCon.CondVal = POSSet.Fields.Item("DocEntry").Value.ToString();

                    //            POSSet.MoveNext();
                    //            if (!POSSet.EoF)
                    //                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                    //        }
                    //    }
                    //    else
                    //    {
                    //        oCon = oCons.Add();
                    //        oCon.Alias = "DocEntry";
                    //        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    //        oCon.CondVal = "-1";
                    //    }

                    //    oCfl.SetConditions(oCons);
                    //}
                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && pVal.ItemUID == "ETOTTNO" && pVal.BeforeAction == false)
                    {
                        try
                        {
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                            SAPbouiCOM.DataTable oDataTable;
                            oDataTable = oCFLEvento.SelectedObjects;

                            if (oDataTable != null)
                            {
                                Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item(pVal.ItemUID).Specific;
                                Global.oEdit.Value = oDataTable.GetValue(Global.oEdit.ChooseFromListAlias, 0).ToString().Trim();

                                try
                                {
                                    //Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETOTTNO").Specific;
                                    //Global.oEdit.Value = oDataTable.GetValue("DocNum", 0).ToString().Trim();

                                    Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETOTNTRY").Specific;
                                    Global.oEdit.Value = oDataTable.GetValue("DocEntry", 0).ToString().Trim();
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }
                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && pVal.ItemUID == "ETSCNO" && pVal.BeforeAction == false)
                    {
                        try
                        {
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                            SAPbouiCOM.DataTable oDataTable;
                            oDataTable = oCFLEvento.SelectedObjects;

                            if (oDataTable != null)
                            {
                                Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item(pVal.ItemUID).Specific;
                                Global.oEdit.Value = oDataTable.GetValue(Global.oEdit.ChooseFromListAlias, 0).ToString().Trim();

                                try
                                {
                                    //Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETOTTNO").Specific;
                                    //Global.oEdit.Value = oDataTable.GetValue("DocNum", 0).ToString().Trim();

                                    Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETSCNTRY").Specific;
                                    Global.oEdit.Value = oDataTable.GetValue("DocEntry", 0).ToString().Trim();
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }
                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && pVal.ItemUID == "ETSTYLNO" && pVal.BeforeAction == false)
                    {
                        try
                        {
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                            SAPbouiCOM.DataTable oDataTable;
                            oDataTable = oCFLEvento.SelectedObjects;

                            if (oDataTable != null)
                            {
                                Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item(pVal.ItemUID).Specific;
                                Global.oEdit.Value = oDataTable.GetValue(Global.oEdit.ChooseFromListAlias, 0).ToString().Trim();


                                Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETSTYLDS").Specific;
                                Global.oEdit.Value = oDataTable.GetValue("U_STYLENM", 0).ToString().Trim();

                                Global.oEdit = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETSLNTRY").Specific;
                                Global.oEdit.Value = oDataTable.GetValue("DocEntry", 0).ToString().Trim();

                            }
                        }
                        catch { }
                    }
                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "BTNSTYLD" && pVal.BeforeAction == false)
                    {
                        try
                        {
                            SAPbouiCOM.EditText oETSTYLNO =
                                (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETSTYLNO").Specific;

                            if (string.IsNullOrWhiteSpace(oETSTYLNO.Value))
                            {
                                Application.SBO_Application.StatusBar.SetText(
                                    "Please select a Style Master.",
                                    SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                                return;
                            }

                            Global.G_Form.Freeze(true);

                            LoadGrid(ref Global.G_Form, false);

                            Global.G_Form.Freeze(false);
                        }
                        catch (Exception ex)
                        {
                            Global.G_Form.Freeze(false);
                        }
                    }
                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "BTQty" && pVal.BeforeAction == false)
                    {
                        try
                        {
                            Global.G_Form.Freeze(true);
                            LoadMatrix(ref Global.G_Form);
                            Global.G_Form.Freeze(false);
                        }
                        catch (Exception ex)
                        {
                            Global.G_Form.Freeze(false);
                        }
                    }
                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "1" && pVal.BeforeAction == true && Global.G_Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        try
                        {
                            Global.G_Form.Freeze(true);

                            // ##### Insert Size
                            int SizeEntry = 0;
                            var oItem = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETCRSZNTRY").Specific;
                            if (!string.IsNullOrEmpty(oItem.Value))
                                SizeEntry = Convert.ToInt32(oItem.Value);

                            try
                            {
                                Global.oComp.StartTransaction();

                                InsertSizeDetails(Global.G_Form, "", "", SizeEntry);

                                Global.oComp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            }
                            catch (Exception ex)
                            {
                                try
                                {
                                    Global.oComp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                }
                                catch { }

                                try
                                {
                                    Global.oComp.StartTransaction();
                                    InsertSizeDetails(Global.G_Form, "", "", SizeEntry);
                                    Global.oComp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                }
                                catch (Exception ex2)
                                {
                                    Global.G_UI_Application.StatusBar.SetText(ex2.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                                                                        SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    try
                                    {
                                        Global.oComp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    }
                                    catch { }
                                }
                            }

                            Global.G_Form.Freeze(false);
                        }
                        catch (Exception ex)
                        {
                            Global.G_Form.Freeze(false);
                        }

                    }
                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "1" && pVal.BeforeAction == false && Global.G_Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        try
                        {
                            Global.G_Form.Freeze(true);
                            LoadGrid(ref Global.G_Form, false);
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
            if (BusinessObjectInfo.BeforeAction == false && BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD)
            {
                Global.G_Form = Global.G_UI_Application.Forms.Item(BusinessObjectInfo.FormUID);
                switch (BusinessObjectInfo.FormTypeEx)
                {
                    case "149":
                        {
                            try
                            {
                                Global.G_Form.Freeze(true);
                                LoadGrid(ref Global.G_Form, true);
                                Global.G_Form.Freeze(false);
                            }
                            catch (Exception ex)
                            {
                                Global.G_Form.Freeze(false);
                            }


                            break;
                        }
                }
            }
            else if (BusinessObjectInfo.BeforeAction == true && BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            {
                Global.G_Form = Global.G_UI_Application.Forms.Item(BusinessObjectInfo.FormUID);
                switch (BusinessObjectInfo.FormTypeEx)
                {
                    case "149":
                        {
                            try
                            {
                                Global.G_Form.Freeze(true);

                                string ObjectType = BusinessObjectInfo.Type;
                                //XmlDocument xmlobjEntry = new XmlDocument();
                                //xmlobjEntry.LoadXml(BusinessObjectInfo.ObjectKey);
                                //int ObjectEntry = Convert.ToInt32(xmlobjEntry.SelectSingleNode("//DocEntry").InnerText);

                                int SizeEntry = string.IsNullOrEmpty(((SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETCRSZNTRY").Specific).Value)
                                    ? 0
                                    : Convert.ToInt32(((SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETCRSZNTRY").Specific).Value);

                                try
                                {
                                    Global.oComp.StartTransaction();

                                    if (ObjectType == "112")
                                    {
                                        InsertSizeDetails(Global.G_Form, "", "", SizeEntry);
                                    }

                                    Global.oComp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                }
                                catch (Exception ex)
                                {
                                    try
                                    {
                                        Global.oComp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    }
                                    catch { }

                                    try
                                    {
                                        Global.oComp.StartTransaction();

                                        if (ObjectType == "112")
                                        {
                                            InsertSizeDetails(Global.G_Form, "", "", SizeEntry);
                                        }

                                        Global.oComp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                    }
                                    catch (Exception ex2)
                                    {
                                        Global.G_UI_Application.StatusBar.SetText(ex2.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                                             SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                                        try
                                        {
                                            Global.oComp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                        }
                                        catch { }
                                    }
                                }

                                Global.G_Form.Freeze(false);
                            }
                            catch (Exception ex)
                            {
                                Global.G_Form.Freeze(false);
                            }


                            break;
                        }
                }
            }

        }


        public void openStyleMaster(string styleCode)
        {
            try
            {
                StyleMaster styleMaster = new StyleMaster();
                styleMaster.Show();
                Global.G_Form = Application.SBO_Application.Forms.Item("FIL_FRM_STYLMSTR");
                Global.G_Form.Freeze(true);
                Global.G_Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                Global.G_Form.Items.Item("ETSLCODE").Enabled = true;
                SAPbouiCOM.EditText cETSLCODE = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETSLCODE").Specific;
                cETSLCODE.Value = styleCode;
                Global.G_Form.Items.Item("1").Click();
                Global.G_Form.Items.Item("FOLSIZE").Click();
                Global.G_Form.Freeze(false);
            }
            catch (Exception ex)
            {
                Global.G_Form.Freeze(false);

                Application.SBO_Application.StatusBar.SetText("Error in open StyleMaster: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        //infuture
        //public void openOTT(string OTTEntry)
        //{
        //    try
        //    {
        //        OTT oTT = new OTT();
        //        oTT.Show();
        //        Global.G_Form = Application.SBO_Application.Forms.Item("FIL_FRM_OTT");
        //        Global.G_Form.Freeze(true);
        //        Global.G_Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
        //        Global.G_Form.Items.Item("ETDOCTRY").Enabled = true;
        //        SAPbouiCOM.EditText cETOTTID = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETDOCTRY").Specific;
        //        cETOTTID.Value = OTTEntry;
        //        Global.G_Form.Items.Item("1").Click();
        //        Global.G_Form.Freeze(false);
        //    }
        //    catch (Exception ex)
        //    {
        //        Global.G_Form.Freeze(false);

        //        Application.SBO_Application.StatusBar.SetText("Error in open OTT: " + ex.Message,
        //            SAPbouiCOM.BoMessageTime.bmt_Short,
        //            SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //    }
        //}


        //public void openSalesContract(string OTTEntry)
        //{
        //    try
        //    {
        //        SalesContract sc = new SalesContract();
        //        sc.Show();
        //        Global.G_Form = Application.SBO_Application.Forms.Item("FIL_FRM_SLCNTRCT");
        //        Global.G_Form.Freeze(true);
        //        Global.G_Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
        //        Global.G_Form.Items.Item("ETDOCTRY").Enabled = true;
        //        SAPbouiCOM.EditText cETOTTID = (SAPbouiCOM.EditText)Global.G_Form.Items.Item("ETDOCTRY").Specific;
        //        cETOTTID.Value = OTTEntry;
        //        Global.G_Form.Items.Item("1").Click();
        //        Global.G_Form.Freeze(false);
        //    }
        //    catch (Exception ex)
        //    {
        //        Global.G_Form.Freeze(false);

        //        Application.SBO_Application.StatusBar.SetText("Error in open OTT: " + ex.Message,
        //            SAPbouiCOM.BoMessageTime.bmt_Short,
        //            SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //    }
        //}
        private void LoadGrid(ref SAPbouiCOM.Form pform, bool FromDataEvent)
        {
            try
            {
                string qStr = string.Empty;

                if (FromDataEvent == false)
                {
                    SAPbouiCOM.EditText oETSLNTRY =
                        (SAPbouiCOM.EditText)pform.Items.Item("ETSLNTRY").Specific;

                    qStr = @"
                            SELECT A.""U_SIZECODE""
                            FROM ""@FIL_DR_PSMST"" A
                            INNER JOIN ""@FIL_MR_STM1"" B
                                ON B.""U_SIZECODE"" = A.""U_SIZECODE""
                            WHERE A.""DocEntry"" = '" + oETSLNTRY.Value.Trim() + @"'
                            GROUP BY A.""U_SIZECODE""
                            ORDER BY A.""U_SIZECODE""";

                    SAPbobsCOM.Recordset SizerSet =
                        (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    SizerSet.DoQuery(qStr);

                    string SizeString = "";

                    while (!SizerSet.EoF)
                    {
                        string sizeCode = SizerSet.Fields.Item("U_SIZECODE").Value.ToString().Trim();
                        string safeSizeCodeAlias = sizeCode.Replace("\"", "");

                        SizeString += @"0 AS """ + safeSizeCodeAlias + @"""";

                        SizerSet.MoveNext();

                        if (!SizerSet.EoF)
                            SizeString += ",";
                    }

                    qStr = @"
                            SELECT ""U_COLORNAME"" AS ""Colour Name"",
                                   ""U_COLORCODE"" AS ""Colour Code"""
                                   + (string.IsNullOrWhiteSpace(SizeString) ? "" : "," + SizeString) + @",
                                   ""Total"" AS ""Total""
                            FROM
                            (
                                SELECT 0 AS ""LineId"",
                                       B.""LineId"" AS ""RCount"",
                                       B.""U_COLORCODE"",
                                       B.""U_COLORNAME"",
                                       0 AS ""Total""
                                FROM ""@FIL_DR_PSMST"" A
                                INNER JOIN ""@FIL_DR_PSMCO"" B
                                    ON A.""DocEntry"" = B.""DocEntry""
                                WHERE A.""DocEntry"" = '" + oETSLNTRY.Value.Trim() + @"'
                                GROUP BY B.""LineId"", B.""U_COLORCODE"", B.""U_COLORNAME""
                            ) V1
                            ORDER BY V1.""RCount""";
                }
                else
                {
                    string docEntry = pform.DataSources.DBDataSources.Item("OQUT")
                                      .GetValue("U_CRSZNTRY", 0).ToString().Trim();

                    qStr = @"
                            SELECT A.""U_SIZECODE""
                            FROM ""@FIL_DR_OQUT"" A
                            INNER JOIN ""OQUT"" B
                                ON B.""U_CRSZNTRY"" = A.""DocEntry""
                            INNER JOIN ""QUT1"" C
                                ON C.""DocEntry"" = B.""DocEntry""
                               AND A.""U_COLORCODE"" = C.""U_FGCOLOUR""
                               AND A.""U_SIZECODE"" = C.""U_FGSIZE""
                            INNER JOIN ""@FIL_DR_PSMST"" D
                                ON D.""DocEntry"" = C.""U_STYLENTRY""
                               AND D.""U_SIZECODE"" = C.""U_FGSIZE""
                            INNER JOIN ""@FIL_MR_STM1"" E
                                ON D.""U_SIZECODE"" = E.""U_SIZECODE""
                            WHERE A.""DocEntry"" = '" + docEntry + @"'
                            GROUP BY A.""U_SIZECODE""
                            ORDER BY A.""U_SIZECODE""";

                    SAPbobsCOM.Recordset SizerSet1 =
                        (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    SizerSet1.DoQuery(qStr);

                    string SizeString = "";

                    while (!SizerSet1.EoF)
                    {
                        string sizeCode = SizerSet1.Fields.Item("U_SIZECODE").Value.ToString().Trim();
                        string safeSizeCodeValue = sizeCode.Replace("'", "''");
                        string safeSizeCodeAlias = sizeCode.Replace("\"", "");

                        SizeString += @"SUM(CASE
                                    WHEN A.""U_SIZECODE"" = '" + safeSizeCodeValue + @"'
                                    THEN A.""U_Qty""
                                    ELSE 0
                                END) AS """ + safeSizeCodeAlias + @"""";

                        SizerSet1.MoveNext();

                        if (!SizerSet1.EoF)
                            SizeString += ",";
                    }

                    qStr = @"
                            SELECT A.""U_COLORNAME"" AS ""Colour Name"",
                                   A.""U_COLORCODE"" AS ""Colour Code"",
                                   SUM(A.""U_Qty"") AS ""Total"""
                                           + (string.IsNullOrWhiteSpace(SizeString) ? "" : "," + SizeString) + @"
                            FROM ""@FIL_DR_OQUT"" A
                            WHERE A.""DocEntry"" = '" + docEntry + @"'
                            GROUP BY A.""U_COLORNAME"", A.""U_COLORCODE""";
                }

                Global.oGrid = (SAPbouiCOM.Grid)pform.Items.Item("SCGrid").Specific;
                Global.oDataTable = pform.DataSources.DataTables.Item("DT_0");

                Global.oDataTable.ExecuteQuery(qStr);
                Global.oGrid.DataTable = Global.oDataTable;

                Global.oGrid.Columns.Item("Colour Name").Editable = false;
                Global.oGrid.Columns.Item("Colour Code").Editable = false;
                Global.oGrid.Columns.Item("Total").Editable = false;

                Global.oGrid.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error in LoadGrid: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        private void LoadMatrix(ref SAPbouiCOM.Form pForm)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)pForm.Items.Item("38").Specific;
                SAPbouiCOM.DataTable oDataTable = pForm.DataSources.DataTables.Item("DT_0");

                double TotalQty = 0;
                double.TryParse(((SAPbouiCOM.EditText)pForm.Items.Item("ETQty").Specific).Value, out TotalQty);

                string styleEntry = pForm.DataSources.DBDataSources.Item("OQUT").GetValue("U_STYLENTRY", 0).Trim();
                string styleCode = pForm.DataSources.DBDataSources.Item("OQUT").GetValue("U_STYLECODE", 0).Trim();

                if (string.IsNullOrWhiteSpace(styleEntry) || string.IsNullOrWhiteSpace(styleCode))
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Style Entry / Style Code is missing.",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return;
                }

                for (int j = 0; j < oDataTable.Rows.Count; j++)
                {
                    string colourCode = oDataTable.Columns.Item("Colour Code").Cells.Item(j).Value.ToString().Trim();

                    string qStr = @"SELECT 
                                A.""ItemCode"",
                                A.""U_COLORCODE"",
                                B.""U_COLORNAME"",
                                A.""U_SIZECODE"",
                                C.""U_SIZENAME""
                            FROM ""OITM"" A
                            INNER JOIN ""@FIL_DR_PSMCO"" B 
                                ON B.""DocEntry"" = '" + styleEntry + @"'
                               AND B.""U_COLORCODE"" = A.""U_COLORCODE""
                            INNER JOIN ""@FIL_DR_PSMST"" C 
                                ON C.""DocEntry"" = '" + styleEntry + @"'
                               AND C.""U_SIZECODE"" = A.""U_SIZECODE""
                            WHERE A.""U_STYLECODE"" = '" + styleCode + @"'
                              AND A.""U_COLORCODE"" = '" + colourCode + @"'
                            ORDER BY B.""LineId"", C.""LineId""";

                    SAPbobsCOM.Recordset rs =
                        (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rs.DoQuery(qStr);

                    while (!rs.EoF)
                    {
                        string sizeCode = rs.Fields.Item("U_SIZECODE").Value.ToString().Trim();

                        double cellValue = 0;
                        object cellObj = oDataTable.Columns.Item(sizeCode).Cells.Item(j).Value;

                        if (cellObj != null && cellObj.ToString().Trim() != "")
                            double.TryParse(cellObj.ToString(), out cellValue);

                        if (cellValue > 0)
                        {
                            SAPbouiCOM.DBDataSource dsQUT1 = pForm.DataSources.DBDataSources.Item("QUT1");

                            if (oMatrix.VisualRowCount == 0)
                            {
                                oMatrix.AddRow();
                            }
                            else
                            {
                                string lastItem = "";
                                try
                                {
                                    lastItem = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1")
                                        .Cells.Item(oMatrix.VisualRowCount).Specific).Value.Trim();
                                }
                                catch { }

                                if (!string.IsNullOrWhiteSpace(lastItem))
                                    oMatrix.AddRow();
                            }

                            oMatrix.FlushToDataSource();

                            int dsRow = dsQUT1.Size - 1;

                            dsQUT1.SetValue("ItemCode", dsRow, rs.Fields.Item("ItemCode").Value.ToString());
                            dsQUT1.SetValue("U_FGCOLOUR", dsRow, rs.Fields.Item("U_COLORCODE").Value.ToString());
                            dsQUT1.SetValue("U_FGCOLRNM", dsRow, rs.Fields.Item("U_COLORNAME").Value.ToString());
                            dsQUT1.SetValue("U_FGSIZE", dsRow, rs.Fields.Item("U_SIZECODE").Value.ToString());
                            dsQUT1.SetValue("U_STYLECODE", dsRow, styleCode);
                            dsQUT1.SetValue("Quantity", dsRow, cellValue.ToString());

                            oMatrix.LoadFromDataSource();

                            TotalQty += cellValue;
                        }

                        rs.MoveNext();
                    }
                }

                ((SAPbouiCOM.EditText)pForm.Items.Item("ETQty").Specific).Value = TotalQty.ToString();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error in Load Matrix: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        private void InsertSizeDetails(SAPbouiCOM.Form pForm, string ObjType, string ObjEntry, int SizeEntry)
        {
            SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string qStr = string.Empty;

            // --- 1. Get next SizeEntry if 0 ---
            if (SizeEntry == 0)
            {
                qStr = "SELECT IFNULL(MAX(\"DocEntry\"),0)+1 \"SizeEntry\" FROM \"@FIL_DR_OQUT\"";
                rSet.DoQuery(qStr);
                SizeEntry = Convert.ToInt32(rSet.Fields.Item("SizeEntry").Value);
            }
            else
            {
                if (ObjType != "112")
                {
                    qStr = $"SELECT * FROM \"OQUT\" WHERE \"U_CRSZNTRY\"='{SizeEntry.ToString().TrimEnd()}'";
                    rSet.DoQuery(qStr);

                    if (rSet.RecordCount > 0)
                    {
                        qStr = "SELECT IFNULL(MAX(\"DocEntry\"),0)+1 \"SizeEntry\" FROM \"@FIL_DR_OQUT\"";
                        rSet.DoQuery(qStr);
                        SizeEntry = Convert.ToInt32(rSet.Fields.Item("SizeEntry").Value);
                    }
                }
            }

            // --- 2. Loop through DataTable rows ---
            Global.oEdit = (SAPbouiCOM.EditText)pForm.Items.Item("ETStyleNo").Specific;
            SAPbouiCOM.DataTable oDataTable = pForm.DataSources.DataTables.Item("DT_0");
            double TotalQty = 0;
            string qStr1 = string.Empty;
            int i = 1;

            for (int j = 0; j < oDataTable.Rows.Count; j++)
            {
                qStr =
                        "Select  A.\"ItemCode\", A.\"U_ColourCode\",B.\"U_ColourName\", A.\"U_SizeCode\" " + Environment.NewLine +
                        " From   \"OITM\" A   " + Environment.NewLine +
                        "    Inner Join \"@FIL_MR_PSMCO\" B On A.\"U_StyleNo\"=B.\"Code\" And A.\"U_ColourCode\"=B.\"U_ColourCode\"  " + Environment.NewLine +
                        "    Inner Join \"@FIL_MR_PSMST\" C On A.\"U_StyleNo\"=C.\"Code\" And A.\"U_SizeCode\"=C.\"U_SizeCode\"  " + Environment.NewLine +
                        $" Where   A.\"U_StyleNo\"='{Global.oEdit.Value}' And B.\"U_ColourAppl\"='Y' And C.\"U_SizeAppl\"='Y'  " + Environment.NewLine +
                        $"   AND A.\"U_ColourCode\"='{oDataTable.Columns.Item("Colour Code").Cells.Item(j).Value}' " + Environment.NewLine +
                        " Order by B.\"LineId\", C.\"LineId\" ";



                SAPbobsCOM.Recordset ColourSizerSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                ColourSizerSet.DoQuery(qStr);

                while (!ColourSizerSet.EoF)
                {
                    string colourCode = ColourSizerSet.Fields.Item("U_ColourCode").Value.ToString();
                    string colourName = ColourSizerSet.Fields.Item("U_ColourName").Value.ToString();
                    string sizeCode = ColourSizerSet.Fields.Item("U_SizeCode").Value.ToString();
                    //string colourName = oDataTable.Columns.Item("Colour Name").Cells.Item(j).Value?.ToString() ?? "";
                    string qty = oDataTable.Columns.Item(sizeCode).Cells.Item(j).Value?.ToString() ?? "0";
                    string totalQty = oDataTable.Columns.Item("Total").Cells.Item(j).Value?.ToString() ?? "0";


                    qStr1 += Environment.NewLine +
                                 $" INSERT INTO \"@FIL_DR_OQUT\"(\"DocEntry\", \"LineId\", \"Object\", \"U_ColourCode\", \"U_ColourName\", \"U_SizeCode\", \"U_Qty\", \"U_TotalQty\") " +
                                 $" SELECT '{SizeEntry}', '{i}', '23', '{colourCode}', '{colourName}', '{sizeCode}', '{qty}', '{totalQty}' FROM DUMMY; ";

                    i++;
                    ColourSizerSet.MoveNext();
                }
            }

            // --- 3. Update or cleanup records ---
            if (!string.IsNullOrEmpty(qStr1))
            {
                if (ObjType == "112")
                {
                    qStr1 += Environment.NewLine +
                             $"UPDATE {(ObjType == "112" ? "\"ODRF\"" : "\"OQUT\"")} SET \"U_CRSZNTRY\"='{SizeEntry}' WHERE \"DocEntry\"='{ObjEntry}'";
                }
            }

            if (ObjType != "112")
            {
                rSet.DoQuery($"SELECT * FROM \"OQUT\" WHERE \"U_CRSZNTRY\"='{SizeEntry.ToString().TrimEnd()}'");
                if (rSet.RecordCount == 0)
                {
                    rSet.DoQuery($"DELETE FROM \"@FIL_DR_OQUT\" WHERE \"DocEntry\"='{SizeEntry.ToString().TrimEnd()}'");
                }
            }

            if (ObjType == "112")
            {
                rSet.DoQuery($"SELECT * FROM \"ODRF\" WHERE \"DocEntry\"='{ObjEntry.ToString().TrimEnd()}'");
                if (rSet.RecordCount > 0)
                {
                    rSet.DoQuery($"SELECT * FROM \"OQUT\" WHERE \"U_CRSZNTRY\"='{SizeEntry.ToString().TrimEnd()}'");
                    if (rSet.RecordCount == 0)
                    {
                        rSet.DoQuery($"DELETE FROM \"@FIL_DR_OQUT\" WHERE \"DocEntry\"='{SizeEntry.ToString().TrimEnd()}'");
                    }
                }
            }

            // --- 4. Execute accumulated inserts ---
            if (!string.IsNullOrEmpty(qStr1))
            {
                SAPbobsCOM.Recordset execSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string qStr2 = "DO " + Environment.NewLine + "BEGIN ";
                qStr1 = qStr2 + qStr1 + Environment.NewLine + "END;";
                execSet.DoQuery(qStr1);


                ((SAPbouiCOM.EditText)pForm.Items.Item("ETCRSZNTRY").Specific).Value = SizeEntry.ToString();
            }
        }
    }
}
