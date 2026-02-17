using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Text;
using Apparel_Dynamic_1._0.Helper;
using Apparel_Dynamic_1._0.Resources.Setup;

namespace Apparel_Dynamic_1._0
{
    class Menu
    {
        public void BasicStart()
        {
            string _AddonId = "2001";
            if (CompanyConnection(_AddonId))
            {
                //Main Module
                CreateMainMenu("43520", "APP", "Apparel", 18, 2, true);

                //Parent Section
                CreateMainMenu("APP", "APP_STP", "Setup", 0, 2, false);
                CreateMainMenu("APP", "APP_MST", "Master", 1, 2, false);
                CreateMainMenu("APP", "APP_TRN", "Transaction", 2, 2, false);

                //Apparel -> Setup
                CreateMainMenu("APP_STP", "APP_STP_SMPLTYPE", "Sample Type",1,1, false);
                CreateMainMenu("APP_STP", "APP_STP_GENTYMSTR", "Gender Type Master",2,1, false);
                CreateMainMenu("APP_STP", "APP_STP_CLRMSTR", "Color Master ",3,1, false);
                CreateMainMenu("APP_STP", "APP_STP_SIZEMSTR", "Size Master",4,1, false);
                CreateMainMenu("APP_STP", "APP_STP_PSTNMSTR", "Position Master ", 5, 1, false);
                CreateMainMenu("APP_STP", "APP_STP_SMPLSTUS", "Sample Status", 6, 1, false);
                CreateMainMenu("APP_STP", "APP_STP_COMSTGES", "Components Stages", 7, 1, false);

                //CreateMainMenu("APP_STP", "APP_STP_PRODLN", "Product Line ", 7, 1, false);
                //CreateMainMenu("APP_STP", "APP_STP_PRODGRP", "Product Group ", 8, 1, false);
                //CreateMainMenu("APP_STP", "APP_STP_PRODCAT", "Product Category", 9, 1, false);
                //CreateMainMenu("APP_STP", "APP_STP_PRODTYPE", "Product Type", 10, 1, false);

              
               
                //Apparel ->  Master
                CreateMainMenu("APP_MST", "APP_MST_ROUTNG", "Route Master", 1, 1, false);


                //Apparel -> Transcation
                CreateMainMenu("APP_TRN", "APP_TRN_COM", "Commercial ", 0, 2, false);
                CreateMainMenu("APP_TRN", "APP_TRN_SAM", "Sampling", 1, 2, false);
                CreateMainMenu("APP_TRN", "APP_TRN_MRD", "Merchandising ", 2, 2, false);

                //Apparel -> Transaction -> Commercial
                CreateMainMenu("APP_TRN_COM", "APP_TRN_COM_EXP", "Export LC ", 0, 2, false);
                CreateMainMenu("APP_TRN_COM", "APP_TRN_COM_IMP", "Import LC ", 1, 2, false);

                //Apparel -> Transaction -> Commercial -> Export LC
                CreateMainMenu("APP_TRN_COM_EXP", "APP_TRN_COM_EXP_MLC", "Master LC", 1, 1, false);
                CreateMainMenu("APP_TRN_COM_EXP", "APP_TRN_COM_EXP_AMD", "Master LC Amendment", 2, 1, false);

                //Apparel -> Transaction -> Commercial -> Import  LC
                CreateMainMenu("APP_TRN_COM_IMP", "APP_TRN_COM_IMP_B2B_LC", "Import LC/TT/RTGS LC(B2B)", 0, 1, false);
                CreateMainMenu("APP_TRN_COM_IMP", "APP_TRN_COM_IMP_B2B_AMD", "Import LC/TT/RTGS LC Ammendment (B2B)", 1, 1, false);

                //Apparel -> Transaction -> Sampling
                CreateMainMenu("APP_TRN_SAM", "APP_TRN_SAM_SM", "Sample Master", 1, 1, false);
                CreateMainMenu("APP_TRN_SAM", "APP_TRN_SAM_SPC", "Sample PreCositng", 2, 1, false);

                //Apparel -> Transaction -> Merchandising 
                CreateMainMenu("APP_TRN_MRD", "APP_TRN_MRD_SM", "Style Master", 1, 1, false);
                CreateMainMenu("APP_TRN_MRD", "APP_TRN_MRD_SCON", "Sales Contract ", 2, 2, false);

                //Apparel -> Transaction -> Merchandising->Sales Contract
                CreateMainMenu("APP_TRN_MRD_SCON", "APP_TRN_MRD_SCON_SC", "Sales Contract", 1, 1, false);
                CreateMainMenu("APP_TRN_MRD_SCON", "APP_TRN_MRD_SCON_SC_AMD", "Sales Contract Ammendment", 2, 1, false);

                //Apparel -> Transaction -> Merchandising -> OTT
                CreateMainMenu("APP_TRN_MRD", "APP_TRN_MRD_OTT", "OTT", 3, 1, false);

                //Apparel -> Transaction -> Merchandising -> Draft Order
                CreateMainMenu("APP_TRN_MRD", "APP_TRN_MRD_DRO", "Draft Order", 4, 1, false);

                //Apparel -> Transaction -> Merchandising -> CAD Fabric Consuption - FHR
                CreateMainMenu("APP_TRN_MRD", "APP_TRN_MRD_CAD", "CAD", 5, 1, false);

                //Apparel -> Transaction -> Merchandising -> Costing
                CreateMainMenu("APP_TRN_MRD", "APP_TRN_MRD_CST", "Procurement Costing", 6, 1, false);

                //Apparel -> Transaction -> Merchandising -> Material requirements planning  - FHR
                CreateMainMenu("APP_TRN_MRD", "APP_TRN_MRD_MRP", "MRP", 7, 1, false);


            }

        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                //___________________________________________________________Setup_______________________________________________
                
                //Sample Type
                if (pVal.BeforeAction && pVal.MenuUID == "APP_STP_SMPLTYPE")
                {
                        string formUID = "FIL_FRM_SMPLTYPE";
                        if (IsFormOpen(formUID))
                        {
                            Global.G_UI_Application.Forms.Item(formUID).Select();
                            Global.G_UI_Application.StatusBar.SetText("Form already opened once.",
                                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            return;
                        }
                    SampleType activeForm = new SampleType();
                    activeForm.Show();
                }
                // Gender 
                else if (pVal.BeforeAction && pVal.MenuUID == "APP_STP_GENTYMSTR")
                {
                    string formUID = "FIL_FRM_GENTYMSTR";
                    if (IsFormOpen(formUID))
                    {
                        Global.G_UI_Application.Forms.Item(formUID).Select();
                        Global.G_UI_Application.StatusBar.SetText("Form already opened once.",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        return;
                    }
                    Gender activeForm = new Gender();
                    activeForm.Show();
                }
                //Size
                else if (pVal.BeforeAction && pVal.MenuUID == "APP_STP_SIZEMSTR")
                {
                    string formUID = "FIL_FRM_SIZEMSTR";
                    if (IsFormOpen(formUID))
                    {
                        Global.G_UI_Application.Forms.Item(formUID).Select();
                        Global.G_UI_Application.StatusBar.SetText("Form already opened once.",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        return;
                    }
                    Size activeForm = new Size();
                    activeForm.Show();
                }
                //Colour
                else if (pVal.BeforeAction && pVal.MenuUID == "APP_STP_CLRMSTR")
                {
                    string formUID = "FIL_FRM_CLR_MSTR";
                    if (IsFormOpen(formUID))
                    {
                        Global.G_UI_Application.Forms.Item(formUID).Select();
                        Global.G_UI_Application.StatusBar.SetText("Form already opened once.",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        return;
                    }
                    Colour activeForm = new Colour();
                    activeForm.Show();
                }
                //Component Stages
                else if (pVal.BeforeAction && pVal.MenuUID == "APP_STP_COMSTGES")
                {
                    string formUID = "FIL_FRM_COMSTG";
                    if (IsFormOpen(formUID))
                    {
                        Global.G_UI_Application.Forms.Item(formUID).Select();
                        Global.G_UI_Application.StatusBar.SetText("Form already opened once.",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        return;
                    }
                    ComponentStages activeForm = new ComponentStages();
                    activeForm.Show();
                }

                //___________________________________________________________Master________________________________________________


                //___________________________________________________________Standard______________________________________________
                //ADD
                else if (!pVal.BeforeAction && pVal.MenuUID == "1282")
                {
                    SAPbouiCOM.Form oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                    string formtype = oForm.UniqueID.ToString();
                    switch (formtype)
                    {
                       
                        case "FIL_FRM_COMSTG":
                            {
                                SetComStageEnable(oForm);
                                break;
                            }
                        case "FIL_FRM_SMPLTYPE":
                            {
                                SAPbouiCOM.Item oUomItem = oForm.Items.Item("ETCODE");
                                oUomItem.Enabled = true;
                                break;
                            }
                        case "FIL_FRM_GENTYMSTR":
                            {
                                SAPbouiCOM.Item oUomItem = oForm.Items.Item("ETCODE");
                                oUomItem.Enabled = true;
                                break;
                            }
                        case "FIL_FRM_SMPLMSTR":
                            {
                                //Initial State
                                SAPbouiCOM.Matrix MTSZ = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXSIZE").Specific;
                                SAPbouiCOM.Matrix MTXCLR = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXCOLOR").Specific;
                                SAPbouiCOM.Matrix MTXITM = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXITEM").Specific;
                                SAPbouiCOM.Matrix MTXBYRS = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXBUYER").Specific;
                                SAPbouiCOM.Matrix MTXATTAC = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXATTCH").Specific;

                                MTSZ.AutoResizeColumns();
                                MTXCLR.AutoResizeColumns();
                                MTXITM.AutoResizeColumns();
                                MTXBYRS.AutoResizeColumns();
                                MTXATTAC.AutoResizeColumns();

                                //Series Initialization
                                SAPbouiCOM.DBDataSource oDBH = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@FIL_DH_SMPLMAST");   //DEFINE  DATASOURCES.
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                {
                                    SAPbouiCOM.ComboBox ocmb = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBSERIES").Specific;
                                    Global.GFunc.LoadComboBoxSeries(ocmb, "FIL_D_SMPLMAST");  //Object Type
                                    string ocmbvalue = ocmb.Selected.Value;
                                    long docno = oForm.BusinessObject.GetNextSerialNumber(ocmbvalue, "FIL_D_SMPLMAST");

                                    oDBH.SetValue("DocNum", 0, docno.ToString()); // only set the value in string.
                                }
                                //Date
                                ((SAPbouiCOM.EditText)oForm.Items.Item("ETSLCNDT").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
                                //Sample Code Enable
                                SAPbouiCOM.Item oSampleCode = oForm.Items.Item("ETSLCODE");
                                oSampleCode.Enabled = true;
                                //button disable 
                                SAPbouiCOM.Item oBtnItmTx = oForm.Items.Item("BTNITMTX");
                                SAPbouiCOM.Item oBtnItmCr = oForm.Items.Item("BTNITMCR");
                                oBtnItmCr.Enabled = false;
                                oBtnItmTx.Enabled = false;

                                break;
                            }
                        case "FIL_FRM_SMPLPCST":
                            {
                                //Series Initialization
                                SAPbouiCOM.DBDataSource oDBH = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@FIL_DH_PRECOSTING");
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                {
                                    SAPbouiCOM.ComboBox ocmb = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBSERIES").Specific;
                                    Global.GFunc.LoadComboBoxSeries(ocmb, "FIL_D_PRECOSTING");
                                    string ocmbvalue = ocmb.Selected.Value;
                                    long docno = oForm.BusinessObject.GetNextSerialNumber(ocmbvalue, "FIL_D_PRECOSTING");
                                    oDBH.SetValue("DocNum", 0, docno.ToString());
                                }
                                //Date
                                ((SAPbouiCOM.EditText)oForm.Items.Item("ETDATE").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
                                ((SAPbouiCOM.EditText)oForm.Items.Item("ETVERSON").Specific).Value = "1"; //Default version 

                                //Enable off
                                SetItemsEnabled(oForm, false, "ETSMPLNM", "ETBUYER", "ETBYRNM", "ETDOCNUM", "ETDATE", "ETVERSON");

                                break;
                            }
                    }
                }
                //Find Mode
                else if (!pVal.BeforeAction && pVal.MenuUID == "1281")
                {
                    SAPbouiCOM.Form oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                    string formtype = oForm.UniqueID.ToString();
                    switch (formtype)
                    {
 
                        case "FIL_FRM_SMPLMSTR":
                            {
                                SAPbouiCOM.Item oSampleCode = oForm.Items.Item("ETSLCODE");
                                oSampleCode.Enabled = true;
                                //button  
                                SAPbouiCOM.Item oBtnItmTx = oForm.Items.Item("BTNITMTX");
                                SAPbouiCOM.Item oBtnItmCr = oForm.Items.Item("BTNITMCR");
                                oBtnItmCr.Enabled = false;
                                oBtnItmTx.Enabled = false;
                                SampleEnableButtons(oForm);
                                break;
                            }
                        case "FIL_FRM_SMPLPCST":
                            {
                                //Enable off
                                SetItemsEnabled(oForm, true, "ETSMPLNM", "ETBUYER", "ETBYRNM", "ETDOCNUM", "ETDATE", "ETVERSON");

                                break;
                            }
                    }
                }
                //First
                else if (!pVal.BeforeAction && pVal.MenuUID == "1288")
                {
                    SAPbouiCOM.Form oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                    string formtype = oForm.UniqueID.ToString();
                    switch (formtype)
                    {
                        
                        case "FIL_FRM_COMSTG":
                            {
                                SetComStageEnable(oForm);
                                break;
                            }
                        case "FIL_FRM_SMPLTYPE":
                            {
                                SAPbouiCOM.Item oUomItem = oForm.Items.Item("ETCODE");
                                oUomItem.Enabled = false;
                                break;
                            }
                        case "FIL_FRM_GENTYMSTR":
                            {
                                SAPbouiCOM.Item oUomItem = oForm.Items.Item("ETCODE");
                                oUomItem.Enabled = false;
                                break;
                            }
                        case "FIL_FRM_SMPLPCST":
                            {
                                //Enable off
                                SetItemsEnabled(oForm, false, "ETSMPLNM", "ETBYRNM", "ETDOCNUM", "ETDATE", "ETVERSON");
                                SetItemsEnabled(oForm, true, "BTNLCSTH", "BTNVRNUP");
                                AddLineIfLastRowHasValue(oForm, "MTXCMPNT", "@FIL_DR_PRECOSTCOMP", "U_ROUTSTAG");
                                string route = GetRouteFromSampleCode(oForm);
                                LoadRouteWiseComboToMatrixColumn(oForm, "MTXCMPNT", "CLRSTGCD", route);
                                break;
                            }

                    }
                }
                //Previous
                else if (!pVal.BeforeAction && pVal.MenuUID == "1289")
                {
                    SAPbouiCOM.Form oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                    string formtype = oForm.UniqueID.ToString();
                    switch (formtype)
                    {
                        
                        case "FIL_FRM_SMPLMSTR":
                            {
                                //Matrix Open 
                                EnsureLine(oForm, "MTXSIZE", "@FIL_DR_SMPLSIZE");
                                EnsureLine(oForm, "MTXCOLOR", "@FIL_DR_SMPLCOLO");
                                EnsureLine(oForm, "MTXBUYER", "@FIL_DR_SMPLBUYER");
                                AddLineIfLastRowHasValue(oForm, "MTXCOLOR", "@FIL_DR_SMPLCOLO", "U_COLOCODE");
                                AddLineIfLastRowHasValue(oForm, "MTXSIZE", "@FIL_DR_SMPLSIZE", "U_SIZECODE");
                                AddLineIfLastRowHasValue(oForm, "MTXBUYER", "@FIL_DR_SMPLBUYER", "U_CARDCODE");

                                SAPbouiCOM.Item oSampleCode = oForm.Items.Item("ETSLCODE");
                                oSampleCode.Enabled = false;
                                //button  
                                SAPbouiCOM.Item oBtnItmTx = oForm.Items.Item("BTNITMTX");
                                SAPbouiCOM.Item oBtnItmCr = oForm.Items.Item("BTNITMCR");
                                oBtnItmCr.Enabled = false;
                                oBtnItmTx.Enabled = false;
                                SampleEnableButtons(oForm);
                                break;
                            }

                        case "FIL_FRM_COMSTG":
                            {
                                SetComStageEnable(oForm);
                                break;
                            }
                        case "FIL_FRM_SMPLTYPE":
                            {
                                SAPbouiCOM.Item oUomItem = oForm.Items.Item("ETCODE");
                                oUomItem.Enabled = false;
                                break;
                            }
                        case "FIL_FRM_GENTYMSTR":
                            {
                                SAPbouiCOM.Item oUomItem = oForm.Items.Item("ETCODE");
                                oUomItem.Enabled = false;
                                break;
                            }
                        case "FIL_FRM_SMPLPCST":
                            {
                                //Enable off
                                SetItemsEnabled(oForm, false, "ETSMPLNM", "ETBYRNM", "ETDOCNUM", "ETDATE", "ETVERSON");
                                SetItemsEnabled(oForm, true, "BTNLCSTH", "BTNVRNUP");
                                AddLineIfLastRowHasValue(oForm, "MTXCMPNT", "@FIL_DR_PRECOSTCOMP", "U_ROUTSTAG");
                                string route = GetRouteFromSampleCode(oForm);
                                LoadRouteWiseComboToMatrixColumn(oForm, "MTXCMPNT", "CLRSTGCD", route);
                                break;
                            }
                    }
                }
                //Next
                else if (!pVal.BeforeAction && pVal.MenuUID == "1290")
                {
                    SAPbouiCOM.Form oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                    string formtype = oForm.UniqueID.ToString();
                    switch (formtype)
                    {
                       
                        case "FIL_FRM_SMPLMSTR":
                            {
                                //Matrix Open 
                                EnsureLine(oForm, "MTXSIZE", "@FIL_DR_SMPLSIZE");
                                EnsureLine(oForm, "MTXCOLOR", "@FIL_DR_SMPLCOLO");
                                EnsureLine(oForm, "MTXBUYER", "@FIL_DR_SMPLBUYER");
                                AddLineIfLastRowHasValue(oForm, "MTXCOLOR", "@FIL_DR_SMPLCOLO", "U_COLOCODE");
                                AddLineIfLastRowHasValue(oForm, "MTXSIZE", "@FIL_DR_SMPLSIZE", "U_SIZECODE");
                                AddLineIfLastRowHasValue(oForm, "MTXBUYER", "@FIL_DR_SMPLBUYER", "U_CARDCODE");

                                SAPbouiCOM.Item oSampleCode = oForm.Items.Item("ETSLCODE");
                                oSampleCode.Enabled = false;
                                //button  
                                SAPbouiCOM.Item oBtnItmTx = oForm.Items.Item("BTNITMTX");
                                SAPbouiCOM.Item oBtnItmCr = oForm.Items.Item("BTNITMCR");
                                oBtnItmCr.Enabled = false;
                                oBtnItmTx.Enabled = false;
                                SampleEnableButtons(oForm);
                                break;
                            }

                        case "FIL_FRM_COMSTG":
                            {
                                SetComStageEnable(oForm);
                                break;
                            }
                        case "FIL_FRM_SMPLTYPE":
                            {
                                SAPbouiCOM.Item oUomItem = oForm.Items.Item("ETCODE");
                                oUomItem.Enabled = false;
                                break;
                            }
                        case "FIL_FRM_GENTYMSTR":
                            {
                                SAPbouiCOM.Item oUomItem = oForm.Items.Item("ETCODE");
                                oUomItem.Enabled = false;
                                break;
                            }
                        case "FIL_FRM_SMPLPCST":
                            {
                                //Enable off
                                SetItemsEnabled(oForm, false, "ETSMPLNM", "ETBYRNM", "ETDOCNUM", "ETDATE", "ETVERSON");
                                SetItemsEnabled(oForm, true, "BTNLCSTH", "BTNVRNUP");
                                AddLineIfLastRowHasValue(oForm, "MTXCMPNT", "@FIL_DR_PRECOSTCOMP", "U_ROUTSTAG");
                                string route = GetRouteFromSampleCode(oForm);
                                LoadRouteWiseComboToMatrixColumn(oForm, "MTXCMPNT", "CLRSTGCD", route);
                                break;
                            }
                    }
                }
                //Last
                else if (!pVal.BeforeAction && pVal.MenuUID == "1291")
                {
                    SAPbouiCOM.Form oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                    string formtype = oForm.UniqueID.ToString();
                    switch (formtype)
                    {
                        case "FIL_FRM_SMPLMSTR":
                            {
                                //Matrix Open 
                                EnsureLine(oForm, "MTXSIZE", "@FIL_DR_SMPLSIZE");
                                EnsureLine(oForm, "MTXCOLOR", "@FIL_DR_SMPLCOLO");
                                EnsureLine(oForm, "MTXBUYER", "@FIL_DR_SMPLBUYER");
                                AddLineIfLastRowHasValue(oForm, "MTXCOLOR", "@FIL_DR_SMPLCOLO", "U_COLOCODE");
                                AddLineIfLastRowHasValue(oForm, "MTXSIZE", "@FIL_DR_SMPLSIZE", "U_SIZECODE");
                                AddLineIfLastRowHasValue(oForm, "MTXBUYER", "@FIL_DR_SMPLBUYER", "U_CARDCODE");

                                SAPbouiCOM.Item oSampleCode = oForm.Items.Item("ETSLCODE");
                                oSampleCode.Enabled = false;
                                //button  
                                SAPbouiCOM.Item oBtnItmTx = oForm.Items.Item("BTNITMTX");
                                SAPbouiCOM.Item oBtnItmCr = oForm.Items.Item("BTNITMCR");
                                oBtnItmCr.Enabled = false;
                                oBtnItmTx.Enabled = false;
                                SampleEnableButtons(oForm);
                                break;
                            }

                        case "FIL_FRM_COMSTG":
                            {
                                SetComStageEnable(oForm);
                                break;
                            }
                        case "FIL_FRM_SMPLTYPE":
                            {
                                SAPbouiCOM.Item oUomItem = oForm.Items.Item("ETCODE");
                                oUomItem.Enabled = false;
                                break;
                            }
                        case "FIL_FRM_GENTYMSTR":
                            {
                                SAPbouiCOM.Item oUomItem = oForm.Items.Item("ETCODE");
                                oUomItem.Enabled = false;
                                break;
                            }
                        case "FIL_FRM_SMPLPCST":
                            {
                                //Enable off
                                SetItemsEnabled(oForm, false, "ETSMPLNM", "ETBYRNM", "ETDOCNUM", "ETDATE", "ETVERSON");
                                SetItemsEnabled(oForm, true, "BTNLCSTH", "BTNVRNUP");
                                AddLineIfLastRowHasValue(oForm, "MTXCMPNT", "@FIL_DR_PRECOSTCOMP", "U_ROUTSTAG");
                                string route = GetRouteFromSampleCode(oForm);
                                LoadRouteWiseComboToMatrixColumn(oForm, "MTXCMPNT", "CLRSTGCD", route);

                                break;
                            }
                    }
                }
                //duplicate
                else if (!pVal.BeforeAction && pVal.MenuUID == "1287")
                {
                    SAPbouiCOM.Form oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                    string formtype = oForm.UniqueID.ToString();
                    switch (formtype)
                    {


                    }
                }
                else if (pVal.BeforeAction && pVal.MenuUID == "FIL_DUPL")
                {
                    SAPbouiCOM.Form oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                    string formtype = oForm.UniqueID.ToString();
                    switch (formtype)
                    {
                       
                    }
                }
                //Delete Row 
                else if (pVal.BeforeAction && pVal.MenuUID == "1293")
                {
                    SAPbouiCOM.Form oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                    string formtype = oForm.UniqueID.ToString();
                    switch (formtype)
                    {
                        
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

        //_____________________________________________________ Method for Working Purpose________________________________________
        private string GetRouteFromSampleCode(SAPbouiCOM.Form oForm)
        {
            try
            {
                string sampleCode = ((SAPbouiCOM.EditText)oForm.Items.Item("ETSMPLCD").Specific).Value.Trim();

                if (string.IsNullOrEmpty(sampleCode))
                    return string.Empty;

                // Prepare SQL Query
                string query = $@"SELECT TOP 1 ""U_ROUTSTAG"" 
                          FROM ""@FIL_DH_SMPLMAST"" 
                          WHERE ""U_SMPLCODE"" = '{sampleCode}'";

                SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)
                    Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRec.DoQuery(query);

                if (!oRec.EoF)
                {
                    return oRec.Fields.Item("U_ROUTSTAG").Value.ToString();
                }

                return string.Empty;
            }
            catch (Exception ex)
            {
                return string.Empty;
            }
        }


        private void LoadRouteWiseComboToMatrixColumn(SAPbouiCOM.Form oForm, string matrixId, string comboColId, string routeCode)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixId).Specific;
            SAPbouiCOM.Column oCol = oMatrix.Columns.Item(comboColId);

            if (oCol.Type != SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                throw new Exception($"Column {comboColId} is not a ComboBox column.");

            //Remove Prev Values
            while (oCol.ValidValues.Count > 0)
                oCol.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);


            string query = $@"Select Distinct ""U_Code"", ""U_Name""
                                from ""@FIL_MR_RSM1""
                                where ""Code"" = '{routeCode}'";

            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(query);
            oCol.ValidValues.Add("", "");

            while (!rs.EoF)
            {
                string v = rs.Fields.Item("U_Code").Value.ToString().Trim();
                string d = rs.Fields.Item("U_Name").Value.ToString().Trim();

                if (!ValidValueExists(oCol, v))
                    oCol.ValidValues.Add(v, d);

                rs.MoveNext();
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
        }

        private bool ValidValueExists(SAPbouiCOM.Column col, string value)
        {
            for (int i = 0; i < col.ValidValues.Count; i++)
            {
                if (col.ValidValues.Item(i).Value == value)
                    return true;
            }
            return false;
        }


        private void SampleEnableButtons(SAPbouiCOM.Form oForm)
        {
            try
            {
                // Only run in VIEW or OK mode
                if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_VIEW_MODE &&
                    oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    return;

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXITEM").Specific;
                SAPbouiCOM.EditText oSampleId = (SAPbouiCOM.EditText)oForm.Items.Item("ETSLCODE").Specific;
                SAPbouiCOM.EditText oDocEntry = (SAPbouiCOM.EditText)oForm.Items.Item("ETDOCTRY").Specific;
                string sampleId = oSampleId.Value.Trim();
                string docEntry = oDocEntry.Value.Trim();

                SAPbouiCOM.Item oBtnItmTx = oForm.Items.Item("BTNITMTX");
                SAPbouiCOM.Item oBtnItmCr = oForm.Items.Item("BTNITMCR");

                bool matrixEmpty = false;

                // --- Check if matrix is empty or has only 1 blank row ---
                if (oMatrix.RowCount == 0 || oMatrix.VisualRowCount == 0)
                {
                    matrixEmpty = true;
                }
                else if (oMatrix.RowCount == 1)
                {
                    SAPbouiCOM.EditText cell = (SAPbouiCOM.EditText)oMatrix.Columns.Item("CLCLRCOD").Cells.Item(1).Specific;
                    if (string.IsNullOrEmpty(cell.Value.Trim()))
                        matrixEmpty = true;
                }

                // --- Case 1: Matrix empty or one blank row ---
                if (matrixEmpty)
                {
                    oBtnItmTx.Enabled = !string.IsNullOrEmpty(sampleId);
                    oBtnItmCr.Enabled = false;
                    return;
                }

                // --- Case 2: Matrix has data ---
                oBtnItmTx.Enabled = false;
                bool enableBtnItmCr = false;

                // 1️⃣ Check DB only if StyleID is not empty //****need to verify****
                if (!string.IsNullOrEmpty(sampleId))
                {
                    SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)
                        Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string query = $"SELECT 1 FROM \"@FIL_DR_SMPLITEM\" WHERE \"DocEntry\" = '{docEntry}'";
                    oRec.DoQuery(query);
                    enableBtnItmCr = !oRec.EoF; // True if record exists
                }

                // 2️⃣ Check if all checkboxes in CLCREAT column are checked
                bool allChecked = true;
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    SAPbouiCOM.CheckBox chk = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("CLCREAT").Cells.Item(i).Specific;
                    if (!chk.Checked)
                    {
                        allChecked = false;
                        break;
                    }
                }

                // Final decision for BTNITMCR
                oBtnItmCr.Enabled = enableBtnItmCr && !allChecked;
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error in StyleEnableButtons: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
        }


        public bool IsFormOpen(string formUID)
        {
            try
            {
                foreach (SAPbouiCOM.Form form in Application.SBO_Application.Forms)
                {
                    if (form.UniqueID == formUID)
                    {
                        // Check if it's visible and not closed
                        if (form.Visible)
                        {
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Global.G_UI_Application.StatusBar.SetText("Error checking form: " + ex.Message,
                   SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            return false; // Form is not open
        }



        private void SetItemsEnabled(SAPbouiCOM.Form oForm, bool enabled, params string[] itemIds)
        {
            foreach (string itemId in itemIds)
            {
                try
                {
                    oForm.Items.Item(itemId).Enabled = enabled;
                }
                catch
                {

                }
            }
        }
        public static void EnsureLine(SAPbouiCOM.Form oForm, string matrixID, string dbTable)
        {
            SAPbouiCOM.Matrix matrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixID).Specific;
            SAPbouiCOM.DBDataSource db = oForm.DataSources.DBDataSources.Item(dbTable);
            if (matrix.RowCount == 0)
            {
                Global.GFunc.SetNewLine(matrix, db, 1, "");
            }
        }

        public static void AddLineIfLastRowHasValue(
           SAPbouiCOM.Form oForm,
           string matrixID,
           string dbTable,
           string columnName
           )
        {
            try
            {
                SAPbouiCOM.Matrix matrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixID).Specific;
                SAPbouiCOM.DBDataSource db = oForm.DataSources.DBDataSources.Item(dbTable);
                matrix.FlushToDataSource();
                int dbRowCount = db.Size;
                if (dbRowCount == 0)
                {
                    Global.GFunc.SetNewLine(matrix, db, 1, "");
                    return;
                }
                int lastDbRow = dbRowCount - 1;
                string lastValue = db.GetValue(columnName, lastDbRow).Trim();
                if (!string.IsNullOrEmpty(lastValue) && !lastValue.Equals("0.0"))
                {
                    Global.GFunc.SetNewLine(matrix, db, dbRowCount + 1, "");
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox("AddLineIfLastRowHasValue Error: " + ex.Message);
            }
        }

        private void SetComStageEnable(SAPbouiCOM.Form oForm)
        {
            if (oForm == null)
                return;

            SAPbouiCOM.Item itmCode = oForm.Items.Item("ETCODE");
            SAPbouiCOM.Item itmUOM = oForm.Items.Item("ETUOM");
            SAPbouiCOM.ComboBox cbUOMApply = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBUOMAPL").Specific;


            itmCode.Enabled = (oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE);


            string uomApplyVal = "";

            if (cbUOMApply.Selected != null)
                uomApplyVal = cbUOMApply.Selected.Value;

            if (uomApplyVal == "Y")
            {
                itmUOM.Enabled = true;
            }
            else
            {
                itmUOM.Enabled = false;
                ((SAPbouiCOM.EditText)itmUOM.Specific).Value = "";
            }
        }

        //________________________________________________________ Connection Related Method_____________________________________________ 

        private bool CompanyConnection(string _AddonId)
        {

            try
            {
                string sErrorMsg;
                string cookie;
                string connStr;
                // Global.ocomp.
                Global.oComp = new SAPbobsCOM.Company();
                cookie = Global.oComp.GetContextCookie();
                //    Global.oCompany = new SAPbobsCOM.Company();
                //   cookie =Global.oCompany.GetContextCookie();
                connStr = Application.SBO_Application.Company.GetConnectionContext(cookie);
                Global.oComp.SetSboLoginContext(connStr);
                ////   if (Global.CF.IsSAPHANA())
                ////  {
                ////   Global.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
                //// }
                //// else
                //// {
                //Global.ocomp.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019;
                // }
                // Global.oCompany.Connect();
                Global.G_UI_Application = Application.SBO_Application;
                Global.oComp = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany(); // Reassign the ocomp with the session we conencted with sap b1
                                                                                                       // sErrorMsg = Global.oCompany.GetLastErrorDescription();
                if (!ValidateLicense(_AddonId))
                {
                    Application.SBO_Application.StatusBar.SetText("License validation failed for Apparel Add-On", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    System.Windows.Forms.Application.Exit();
                    return false;
                }
                else
                {
                    Application.SBO_Application.StatusBar.SetText("Apparel Add-On Connected Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    return true;
                }
            }
            catch
            {
                Application.SBO_Application.MessageBox(Global.oComp.GetLastErrorDescription().ToString(), 1, "OK", "", "");
                return false;
            }
        }

        private static bool ValidateLicense(string addonId)
        {
            try
            {
                // Example: license info stored in custom UDT [@MYADDON_LICENSE]
                // Fields: U_AddonId, U_CompanyDB, U_LicenseKey, U_ExpiryDate

                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sql = string.Format("SELECT A.{0}U_LICSEKEY{0},A.{0}U_EXPIRDATE{0} from {0}@FIL_MH_LICENSE{0} A  Where A.{0}U_ADDONID{0}='" + addonId + "' AND A.{0}U_COMNYDB{0}='" + Global.oComp.CompanyDB.ToString() + "' ", '"');
                rs.DoQuery(sql);

                if (rs.RecordCount == 0)
                    return false;

                string licenseKey = rs.Fields.Item("U_LICSEKEY").Value.ToString();
                DateTime expiryDate = Convert.ToDateTime(rs.Fields.Item("U_EXPIRDATE").Value);

                // Validate Expiry
                if (DateTime.Now > expiryDate)
                    return false;

                // Validate Key (example check — implement your own hash/algorithm)
                string expectedKey = GenerateExpectedKey(Global.oComp.CompanyDB.ToString(), addonId);
                if (licenseKey != expectedKey)
                    return false;

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string GenerateExpectedKey(string companyDb, string addonId)
        {
            // Example license key algorithm (you can implement stronger encryption)
            return (companyDb + "_" + addonId + "_2025").ToUpper();
        }

        public void CreateMainMenu(string ParentMenuID, string MenuID, string MenuName, int Position, int imenutype, bool flgimg)
        {
            try
            {
                SAPbouiCOM.Menus oMenus = null;
                SAPbouiCOM.MenuItem oMenuItem = null;
                oMenus = Application.SBO_Application.Menus;
                // ✅ Skip if already exists
                if (oMenus.Exists(MenuID))
                {
                    return;
                }
                SAPbouiCOM.MenuCreationParams oCreationPackage = null;
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oMenuItem = Application.SBO_Application.Menus.Item(ParentMenuID);

                switch (imenutype)
                {
                    case 2:
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                        break;
                    case 1:
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        break;
                    case 3:
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_SEPERATOR;
                        break;
                }

                oCreationPackage.UniqueID = MenuID;
                oCreationPackage.String = MenuName;
                oCreationPackage.Enabled = true;
                oCreationPackage.Position = Position;

                string path = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath).ToString();
                if (flgimg == true)
                {
                    string Apparel = string.Concat(path, @"AddFiles\Logo\test.jpg");
                    oCreationPackage.Image = Apparel;
                }

                oMenus = oMenuItem.SubMenus;

                try
                {
                    oMenus.AddEx(oCreationPackage);
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.MessageBox("Menu already Exists " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox("Unexpected error in CreateMainMenu: " + ex.Message);
            }
        }

    }
}
