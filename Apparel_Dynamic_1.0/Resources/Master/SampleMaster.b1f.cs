using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Apparel_Dynamic_1._0.Helper;

namespace Apparel_Dynamic_1._0.Resources.Master
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Master.SampleMaster", "Resources/Master/SampleMaster.b1f")]
    class SampleMaster : UserFormBase
    {
        public SampleMaster()
        {
        }

        private SAPbouiCOM.StaticText STSLCODE, STSLDESC, STGSM, STSLTYPE, STGENDER, STMERCOD, STNO,
                                     STSLCNDT, STITMGRP, STRUTSTG, STUOM, STCRDCOD;



        private SAPbouiCOM.EditText ETSLCODE, ETSLDESC, ETGSM, ETSLTYPE, ETGENDER, ETMERCOD, ETDOCTRY,
                                    ETDOCNUM, ETSLCNDT, ETITMGNM, ETRUTSTG, ETUOM, ETCRDCOD,
                                    ETMERNAM, ETCRDNAM, ETRUTSNM, ETITMGRP, ETSLTYNM;


        private SAPbouiCOM.LinkedButton LKMERCOD, LKCRDCOD;
        private SAPbouiCOM.Folder FOLCLRSZ, FOLITEM, FOLCRDDT, FOLATTCH, FOLCOLR;
        private SAPbouiCOM.Matrix MTXCOLOR, MTXSIZE, MTXITEM, MTXBUYER, MTXATTCH;
        private SAPbouiCOM.Button ADDButton, CancelButton, BTNITMTX, BTNITMCR, BRWSBTN, DISPBTN, DELBTN;

        private SAPbouiCOM.ComboBox CBSERIES;

        private string sampleCode = "";

        // 🔹 Global storage for both matrices’ previous checkbox states
        private Dictionary<int, bool> prevSztTypeCheckboxStates = new Dictionary<int, bool>();
        private Dictionary<int, bool> prevColorCheckboxStates = new Dictionary<int, bool>();


        public override void OnInitializeComponent()
        {
            this.STSLCODE = ((SAPbouiCOM.StaticText)(this.GetItem("STSLCODE").Specific));
            this.STSLDESC = ((SAPbouiCOM.StaticText)(this.GetItem("STSLDESC").Specific));
            this.STGSM = ((SAPbouiCOM.StaticText)(this.GetItem("STGSM").Specific));
            this.STSLTYPE = ((SAPbouiCOM.StaticText)(this.GetItem("STSLTYPE").Specific));
            this.STGENDER = ((SAPbouiCOM.StaticText)(this.GetItem("STGENDER").Specific));
            this.STMERCOD = ((SAPbouiCOM.StaticText)(this.GetItem("STMERCOD").Specific));
            this.STSLCNDT = ((SAPbouiCOM.StaticText)(this.GetItem("STSLCNDT").Specific));
            this.STITMGRP = ((SAPbouiCOM.StaticText)(this.GetItem("STITMGRP").Specific));
            this.STRUTSTG = ((SAPbouiCOM.StaticText)(this.GetItem("STRUTSTG").Specific));
            this.STUOM = ((SAPbouiCOM.StaticText)(this.GetItem("STUOM").Specific));
            this.STCRDCOD = ((SAPbouiCOM.StaticText)(this.GetItem("STCRDCOD").Specific));
            this.ETSLCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETSLCODE").Specific));
            this.ETSLCODE.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.ETSLCODE_LostFocusAfter);
            this.ETSLDESC = ((SAPbouiCOM.EditText)(this.GetItem("ETSLDESC").Specific));
            this.ETGSM = ((SAPbouiCOM.EditText)(this.GetItem("ETGSM").Specific));
            this.ETSLTYPE = ((SAPbouiCOM.EditText)(this.GetItem("ETSLTYPE").Specific));
            this.ETSLTYPE.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETSLTYPE_ChooseFromListAfter);
            this.ETSLTYPE.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETSLTYPE_ChooseFromListBefore);
            this.ETGENDER = ((SAPbouiCOM.EditText)(this.GetItem("ETGENDER").Specific));
            this.ETGENDER.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETGENDER_ChooseFromListAfter);
            this.ETGENDER.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETGENDER_ChooseFromListBefore);
            this.ETMERCOD = ((SAPbouiCOM.EditText)(this.GetItem("ETMERCOD").Specific));
            this.ETMERCOD.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETMERCOD_ChooseFromListAfter);
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            this.ETDOCNUM = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCNUM").Specific));
            this.ETSLCNDT = ((SAPbouiCOM.EditText)(this.GetItem("ETSLCNDT").Specific));
            this.ETITMGNM = ((SAPbouiCOM.EditText)(this.GetItem("ETITMGNM").Specific));
            this.ETRUTSTG = ((SAPbouiCOM.EditText)(this.GetItem("ETRUTSTG").Specific));
            this.ETRUTSTG.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETRUTSTG_ChooseFromListAfter);
            this.ETUOM = ((SAPbouiCOM.EditText)(this.GetItem("ETUOM").Specific));
            this.ETCRDCOD = ((SAPbouiCOM.EditText)(this.GetItem("ETCRDCOD").Specific));
            this.ETCRDCOD.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETCRDCOD_ChooseFromListAfter);
            this.ETMERNAM = ((SAPbouiCOM.EditText)(this.GetItem("ETMERNAM").Specific));
            this.ETCRDNAM = ((SAPbouiCOM.EditText)(this.GetItem("ETCRDNAM").Specific));
            this.ETRUTSNM = ((SAPbouiCOM.EditText)(this.GetItem("ETRUTSNM").Specific));
            this.ETITMGRP = ((SAPbouiCOM.EditText)(this.GetItem("ETITMGRP").Specific));
            this.ETITMGRP.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETITMGRP_ChooseFromListBefore);
            this.ETITMGRP.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETITMGRP_ChooseFromListAfter);
            this.ETSLTYNM = ((SAPbouiCOM.EditText)(this.GetItem("ETSLTYNM").Specific));
            this.LKMERCOD = ((SAPbouiCOM.LinkedButton)(this.GetItem("LKMERCOD").Specific));
            this.LKCRDCOD = ((SAPbouiCOM.LinkedButton)(this.GetItem("LKCRDCOD").Specific));
            this.FOLCLRSZ = ((SAPbouiCOM.Folder)(this.GetItem("FOLSIZE").Specific));
            this.FOLITEM = ((SAPbouiCOM.Folder)(this.GetItem("FOLITEM").Specific));
            this.FOLCRDDT = ((SAPbouiCOM.Folder)(this.GetItem("FOLCRDDT").Specific));
            this.FOLATTCH = ((SAPbouiCOM.Folder)(this.GetItem("FOLATTCH").Specific));
            this.MTXCOLOR = ((SAPbouiCOM.Matrix)(this.GetItem("MTXCOLOR").Specific));
            this.MTXCOLOR.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.MTXCOLOR_ChooseFromListBefore);
            this.MTXCOLOR.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.MTXCOLOR_ChooseFromListAfter);
            this.MTXSIZE = ((SAPbouiCOM.Matrix)(this.GetItem("MTXSIZE").Specific));
            this.MTXSIZE.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.MTXSIZE_ChooseFromListBefore);
            this.MTXSIZE.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.MTXSIZE_ChooseFromListAfter);
            this.MTXITEM = ((SAPbouiCOM.Matrix)(this.GetItem("MTXITEM").Specific));
            this.MTXBUYER = ((SAPbouiCOM.Matrix)(this.GetItem("MTXBUYER").Specific));
            this.MTXBUYER.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.MTXBUYER_ChooseFromListAfter);
            this.MTXATTCH = ((SAPbouiCOM.Matrix)(this.GetItem("MTXATTCH").Specific));
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.ADDButton.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.ADDButton_PressedAfter);
            this.ADDButton.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.ADDButton_PressedBefore);
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.BTNITMTX = ((SAPbouiCOM.Button)(this.GetItem("BTNITMTX").Specific));
            this.BTNITMTX.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.BTNITMTX_PressedAfter);
            this.BTNITMCR = ((SAPbouiCOM.Button)(this.GetItem("BTNITMCR").Specific));
            this.BTNITMCR.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.BTNITMCR_PressedBefore);
            this.BTNITMCR.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.BTNITMCR_PressedAfter);
            this.FOLCOLR = ((SAPbouiCOM.Folder)(this.GetItem("FOLCOLR").Specific));
            this.CBSERIES = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSERIES").Specific));
            this.STNO = ((SAPbouiCOM.StaticText)(this.GetItem("STNO").Specific));
            this.BRWSBTN = ((SAPbouiCOM.Button)(this.GetItem("BRWSBTN").Specific));
            this.BRWSBTN.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.BRWSBTN_ClickAfter);
            this.DISPBTN = ((SAPbouiCOM.Button)(this.GetItem("DISPBTN").Specific));
            this.DISPBTN.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.DISPBTN_ClickAfter);
            this.DELBTN = ((SAPbouiCOM.Button)(this.GetItem("DELBTN").Specific));
            this.DELBTN.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.DELBTN_ClickAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {


        }



        private void OnCustomInitialize()
        {

        }

        private void ETSLCODE_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    string code = ((SAPbouiCOM.EditText)oForm.Items.Item("ETSLCODE").Specific).Value.Trim();
                    string UCode = Global.GFunc.ToUpperCase(code);
                    ((SAPbouiCOM.EditText)oForm.Items.Item("ETSLCODE").Specific).Value = UCode;
                    if (!string.IsNullOrEmpty(UCode))
                    {
                        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string query = $@"SELECT 1 FROM ""@FIL_DH_SMPLMAST"" WHERE ""U_SMPLCODE"" = '{UCode.Replace("'", "''")}'";
                        oRS.DoQuery(query);
                        if (!oRS.EoF)
                        {
                            Application.SBO_Application.StatusBar.SetText("Code already exists!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            ((SAPbouiCOM.EditText)oForm.Items.Item("ETSLCODE").Specific).Value = "";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Error: Sample Code  " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void BTNITMCR_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                SampleEnableButtons(ref oForm);
            }

        }
        private void BTNITMCR_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                SAPbouiCOM.ProgressBar oProgressBar = null;
                try
                {
                    oForm.Freeze(true);

                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXITEM").Specific;
                    oProgressBar = Global.G_UI_Application.StatusBar.CreateProgressBar("Processing...", oMatrix.RowCount + 1, true);
                    oProgressBar.Text = "Processing...";
                    oProgressBar.Value = 0;

                    SAPbouiCOM.DBDataSource pDataSource = oForm.DataSources.DBDataSources.Item("@FIL_DR_SMPLITEM");
                    pDataSource.Clear();
                    oMatrix.FlushToDataSource();

                    Global.oComp.StartTransaction();

                    for (int i = pDataSource.Size - 1; i >= 0; i--)
                    {
                        string createFlag = pDataSource.GetValue("U_CREATEFLAG", i).Trim();
                        string itemCode = pDataSource.GetValue("U_ITEMCODE", i).Trim();

                        if (createFlag == "N" && !string.IsNullOrEmpty(itemCode))
                        {
                            oProgressBar.Text = "Item Creating for " + itemCode;
                            SAPbobsCOM.Items checkItem = (SAPbobsCOM.Items)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                            if (checkItem.GetByKey(itemCode))
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(checkItem);
                                continue;
                            }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(checkItem);

                            SAPbobsCOM.Items newItem = (SAPbobsCOM.Items)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                            newItem.ItemCode = itemCode;
                            newItem.ItemName = pDataSource.GetValue("U_ITEMNAME", i).Trim();
                            newItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tYES;
                            newItem.SalesItem = SAPbobsCOM.BoYesNoEnum.tYES;
                            newItem.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tYES;

                            newItem.ItemsGroupCode = Convert.ToInt32(((SAPbouiCOM.EditText)oForm.Items.Item("ETITMGRP").Specific).Value.Trim());
                            newItem.PlanningSystem = SAPbobsCOM.BoPlanningSystem.bop_MRP;
                            newItem.ProcurementMethod = SAPbobsCOM.BoProcurementMethod.bom_Make;
                            newItem.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Manual;
                            //newItem.LeadTime = Convert.ToInt32(pDataSource.GetValue("U_LeadTime", i).Trim());

                            string uom = pDataSource.GetValue("U_UOM", i).Trim();
                            newItem.SalesUnit = uom;
                            newItem.InventoryUOM = uom;
                            newItem.PurchaseUnit = uom;

                            //string defaultWhs = oForm.DataSources.DBDataSources.Item("@DTS_OPSM").GetValue("U_WhsCode", 0).Trim();
                            //newItem.DefaultWarehouse = defaultWhs;

                            newItem.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tYES;
                            newItem.SRIAndBatchManageMethod = SAPbobsCOM.BoManageMethod.bomm_OnEveryTransaction;

                            // User-defined fields
                            newItem.UserFields.Fields.Item("U_StyleNo").Value = ((SAPbouiCOM.EditText)oForm.Items.Item("ETSLCODE").Specific).Value.Trim();
                            newItem.UserFields.Fields.Item("U_SizeCode").Value = pDataSource.GetValue("U_SIZECODE", i).Trim();
                            newItem.UserFields.Fields.Item("U_ColourCode").Value = pDataSource.GetValue("U_COLOCODE", i).Trim();
                            //newItem.UserFields.Fields.Item("U_BrandCode").Value = pDataSource.GetValue("U_BrandCode", i).Trim();
                            //newItem.UserFields.Fields.Item("U_SeasonCode").Value = pDataSource.GetValue("U_SeasonCode", i).Trim();
                            //newItem.UserFields.Fields.Item("U_BOMReq").Value = pDataSource.GetValue("U_BOMReq", i).Trim();
                            //newItem.UserFields.Fields.Item("U_SizeGroup").Value = pDataSource.GetValue("U_SizeGroupCode", i).Trim();
                            //newItem.UserFields.Fields.Item("U_MAINFG").Value = pDataSource.GetValue("U_MAINFG", i).Trim();

                            //// Price list
                            //int priceListId = Convert.ToInt32(pDataSource.GetValue("U_PListId", i).Trim());
                            //double priceValue = Convert.ToDouble(pDataSource.GetValue("U_Price", i).Trim() == "0" ? "0.01" : pDataSource.GetValue("U_Price", i).Trim());

                            //for (int j = 0; j < newItem.PriceList.Count; j++)
                            //{
                            //    newItem.PriceList.SetCurrentLine(j);
                            //    if (newItem.PriceList.PriceList == priceListId)
                            //    {
                            //        newItem.PriceList.Price = priceValue;
                            //        break;
                            //    }
                            //}

                            int lRetCode = newItem.Add();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(newItem);

                            if (lRetCode != 0)
                            {
                                int errCode;
                                string errMsg;
                                Global.oComp.GetLastError(out errCode, out errMsg);
                                Global.G_UI_Application.StatusBar.SetText(errMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                                if (Global.oComp.InTransaction)
                                    Global.oComp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                break;
                            }
                        }
                        else if (createFlag == "Y" && !string.IsNullOrEmpty(itemCode))
                        {
                            oProgressBar.Text = "Item Updating for " + itemCode;

                            SAPbobsCOM.Items existingItem = (SAPbobsCOM.Items)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                            if (existingItem.GetByKey(itemCode))
                            {
                                existingItem.ItemName = pDataSource.GetValue("U_ITEMNAME", i).Trim();
                                existingItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tYES;
                                existingItem.SalesItem = SAPbobsCOM.BoYesNoEnum.tYES;
                                existingItem.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tYES;

                                existingItem.ItemsGroupCode = Convert.ToInt32(((SAPbouiCOM.EditText)oForm.Items.Item("ETITMGRP").Specific).Value.Trim());
                                existingItem.PlanningSystem = SAPbobsCOM.BoPlanningSystem.bop_MRP;
                                existingItem.ProcurementMethod = SAPbobsCOM.BoProcurementMethod.bom_Make;
                                existingItem.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Manual;
                                //existingItem.LeadTime = Convert.ToInt32(pDataSource.GetValue("U_LeadTime", i).Trim());

                                string uom = pDataSource.GetValue("U_UOM", i).Trim();
                                existingItem.SalesUnit = uom;
                                existingItem.InventoryUOM = uom;
                                existingItem.PurchaseUnit = uom;

                                //existingItem.DefaultWarehouse = oForm.DataSources.DBDataSources.Item("@DTS_OPSM").GetValue("U_WhsCode", 0).Trim();
                                existingItem.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tYES;
                                existingItem.SRIAndBatchManageMethod = SAPbobsCOM.BoManageMethod.bomm_OnEveryTransaction;

                                existingItem.UserFields.Fields.Item("U_StyleNo").Value = ((SAPbouiCOM.EditText)oForm.Items.Item("ETSLCODE").Specific).Value.Trim();
                                existingItem.UserFields.Fields.Item("U_SizeCode").Value = pDataSource.GetValue("U_SIZECODE", i).Trim();
                                existingItem.UserFields.Fields.Item("U_ColourCode").Value = pDataSource.GetValue("U_COLOCODE", i).Trim();
                                //existingItem.UserFields.Fields.Item("U_BrandCode").Value = pDataSource.GetValue("U_BrandCode", i).Trim();
                                //existingItem.UserFields.Fields.Item("U_SeasonCode").Value = pDataSource.GetValue("U_SeasonCode", i).Trim();
                                //existingItem.UserFields.Fields.Item("U_BOMReq").Value = pDataSource.GetValue("U_BOMReq", i).Trim();
                                //existingItem.UserFields.Fields.Item("U_SizeGroup").Value = pDataSource.GetValue("U_SizeGroupCode", i).Trim();
                                //existingItem.UserFields.Fields.Item("U_MAINFG").Value = pDataSource.GetValue("U_MAINFG", i).Trim();

                                //int priceListId = Convert.ToInt32(pDataSource.GetValue("U_PListId", i).Trim());
                                //double priceValue = Convert.ToDouble(pDataSource.GetValue("U_Price", i).Trim() == "0" ? "0.01" : pDataSource.GetValue("U_Price", i).Trim());

                                //for (int j = 0; j < existingItem.PriceList.Count; j++)
                                //{
                                //    existingItem.PriceList.SetCurrentLine(j);
                                //    if (existingItem.PriceList.PriceList == priceListId)
                                //    {
                                //        existingItem.PriceList.Price = priceValue;
                                //        break;
                                //    }
                                //}

                                int lRetCode = existingItem.Update();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(existingItem);

                                if (lRetCode != 0)
                                {
                                    int errCode;
                                    string errMsg;
                                    Global.oComp.GetLastError(out errCode, out errMsg);
                                    Global.G_UI_Application.StatusBar.SetText(errMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                                    if (Global.oComp.InTransaction)
                                        Global.oComp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    break;
                                }
                            }
                        }
                    }

                    string qStr;
                    if (Global.oComp.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        qStr = $"UPDATE \"@FIL_DR_SMPLITEM\" SET \"U_CREATEFLAG\"='Y' WHERE \"DocEntry\"='{oForm.DataSources.DBDataSources.Item("@FIL_DH_SMPLMAST").GetValue("DocEntry", 0).Trim()}'";
                    }
                    else
                    {
                        qStr = $"UPDATE [@FIL_DR_SMPLITEM] SET U_CREATEFLAG='Y' WHERE Code='{oForm.DataSources.DBDataSources.Item("@FIL_DH_SMPLMAST").GetValue("DocEntry", 0).Trim()}'";
                    }

                    SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rSet.DoQuery(qStr);
                    oProgressBar.Value++;
                    oProgressBar.Stop();
                    Global.oComp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    Global.G_UI_Application.StatusBar.SetText("Successfully Created.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                    Global.G_UI_Application.ActivateMenuItem("1304");
                    oForm.Freeze(false);
                }
                catch (Exception ex)
                {
                    if (Global.oComp.InTransaction)
                        Global.oComp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                    Global.G_UI_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                    try { oProgressBar.Stop(); } catch { }
                    oForm.Freeze(false);
                }
            }


        }
        private void BTNITMTX_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {


            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXITEM").Specific;
            SAPbouiCOM.DBDataSource pDS = oForm.DataSources.DBDataSources.Item("@FIL_DR_SMPLITEM");

            string docEntry = ((SAPbouiCOM.EditText)oForm.Items.Item("ETDOCTRY").Specific).Value;
            if (string.IsNullOrEmpty(docEntry))
            {
                return;
            }
            else
            {
                string query = $@"
                            SELECT
                                ROW_NUMBER() OVER (ORDER BY C.""U_COLOCODE"", S.""U_SIZECODE"") AS ""LineId"",
                                'N' AS ""U_CREATEFLAG"",
                                C.""U_COLOCODE"" AS ""U_COLOCODE"",
                                C.""U_COLONAME"" AS ""U_COLONAME"",
                                S.""U_SIZECODE"" AS ""U_SIZECODE"",
                                S.""U_SIZENAME"" AS ""U_SIZENAME"",
                                M.""U_UOM"" AS ""U_UOM"",
                                M.""U_SMPLCODE"" || '-' || C.""U_COLOCODE"" || '-' || S.""U_SIZECODE"" AS ""U_ITEMCODE"",
                                M.""U_SMPLDESC"" || '-' || C.""U_COLONAME"" || '-' || S.""U_SIZENAME"" AS ""U_ITEMNAME""
                            FROM ""FMIL_PROD"".""@FIL_DH_SMPLMAST"" M
                            JOIN ""FMIL_PROD"".""@FIL_DR_SMPLCOLO"" C
                                ON C.""DocEntry"" = M.""DocEntry""
                            JOIN ""FMIL_PROD"".""@FIL_DR_SMPLSIZE"" S
                                ON S.""DocEntry"" = M.""DocEntry""
                            WHERE
                                M.""DocEntry"" = '{docEntry}'
                                AND COALESCE(C.""U_COLOURAPP"", 'N') = 'Y'
                                AND COALESCE(S.""U_SIZEAPP"", 'N') = 'Y'
                            ORDER BY
                                C.""U_COLOCODE"", S.""U_SIZECODE""";

                oForm.Freeze(true);
                try
                {
                    oRS.DoQuery(query);

                    // --- Clear previous data ---
                    pDS.Clear();
                    oMatrix.Clear();
                    oMatrix.LoadFromDataSource();

                    // --- Populate from recordset ---
                    while (!oRS.EoF)
                    {
                        pDS.InsertRecord(pDS.Size);
                        pDS.Offset = pDS.Size - 1;

                        pDS.SetValue("LineId", pDS.Offset, (pDS.Offset + 1).ToString());
                        pDS.SetValue("U_CREATEFLAG", pDS.Offset, oRS.Fields.Item("U_CREATEFLAG").Value.ToString());
                        pDS.SetValue("U_UOM", pDS.Offset, oRS.Fields.Item("U_UoM").Value.ToString());

                        pDS.SetValue("U_COLOCODE", pDS.Offset, oRS.Fields.Item("U_COLOCODE").Value.ToString());
                        pDS.SetValue("U_COLONAME", pDS.Offset, oRS.Fields.Item("U_COLONAME").Value.ToString());

                        pDS.SetValue("U_SIZECODE", pDS.Offset, oRS.Fields.Item("U_SIZECODE").Value.ToString());
                        pDS.SetValue("U_SIZENAME", pDS.Offset, oRS.Fields.Item("U_SIZENAME").Value.ToString());

                        pDS.SetValue("U_ITEMCODE", pDS.Offset, oRS.Fields.Item("U_ITEMCODE").Value.ToString());
                        pDS.SetValue("U_ITEMNAME", pDS.Offset, oRS.Fields.Item("U_ITEMNAME").Value.ToString());


                        oRS.MoveNext();
                    }

                    // --- Refresh matrix ---
                    oMatrix.LoadFromDataSource();
                    oMatrix.AutoResizeColumns();


                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }

                }
                catch (Exception ex)
                {
                    Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                finally
                {
                    oForm.Freeze(false);
                }

            }



        }

        private void ADDButton_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                ValidateForm(ref oForm, ref BubbleEvent);
            }

        }

        private bool ValidateForm(ref SAPbouiCOM.Form oForm, ref bool BubbleEvent)
        {

            sampleCode = oForm.DataSources.DBDataSources.Item("@FIL_DH_SMPLMAST").GetValue("U_SMPLCODE", 0);
            string itemGrp = oForm.DataSources.DBDataSources.Item("@FIL_DH_SMPLMAST").GetValue("U_ITEMGRP", 0);
            string route = oForm.DataSources.DBDataSources.Item("@FIL_DH_SMPLMAST").GetValue("U_ROUTSTAG", 0);

            if (sampleCode == "")
            {
                Global.GFunc.ShowError("Enter Sample Code");
                oForm.ActiveItem = "ETSLCODE";
                return BubbleEvent = false;
            }
            else if (itemGrp == "")
            {
                Global.GFunc.ShowError("Enter SAP Item Group");
                oForm.ActiveItem = "ETITMGRP";
                return BubbleEvent = false;
            }
            else if (route == "")
            {
                Global.GFunc.ShowError("Enter Route Stage");
                oForm.ActiveItem = "ETRUTSTG";
                return BubbleEvent = false;
            }

            PreventEmptyLastRow(oForm, "@FIL_DR_SMPLSIZE", MTXSIZE, "U_SIZECODE");
            PreventEmptyLastRow(oForm, "@FIL_DR_SMPLCOLO", MTXCOLOR, "U_COLOCODE");
            PreventEmptyLastRow(oForm, "@FIL_DR_SMPLITEM", MTXITEM, "U_ITEMCODE");
            PreventEmptyLastRow(oForm, "@FIL_DR_SMPLBUYER", MTXBUYER, "U_CARDCODE");
            PreventEmptyLastRow(oForm, "@FIL_DR_SMPLATACH", MTXATTCH, "U_ATCHMENT");

            return BubbleEvent;
        }

        private void PreventEmptyLastRow(SAPbouiCOM.Form oForm, string dbDatasourceUID, SAPbouiCOM.Matrix matrix, string columnName)
        {
            SAPbouiCOM.DBDataSource oDB = oForm.DataSources.DBDataSources.Item(dbDatasourceUID);
            int rowCount = matrix.VisualRowCount;

            if (rowCount > 0)
            {
                string lastValue = oDB.GetValue(columnName, rowCount - 1).Trim();

                if (string.IsNullOrEmpty(lastValue) || lastValue.Equals("0.0"))
                {
                    matrix.DeleteRow(rowCount);
                    oDB.RemoveRecord(rowCount - 1);
                }
            }
        }

        private void ADDButton_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                oForm.Freeze(true);
                SAPbouiCOM.EditText sampleEdit = (SAPbouiCOM.EditText)oForm.Items.Item("ETSLCODE").Specific;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                oForm.Items.Item("ETSLCODE").Enabled = true;
                sampleEdit.Value = sampleCode;
                oForm.Items.Item("1").Click();
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                SampleEnableButtons(ref oForm);

                // Add new row if last row has data
                AddLineIfLastRowHasValue(oForm, "MTXCOLOR", "@FIL_DR_SMPLCOLO", "U_COLOCODE");
                AddLineIfLastRowHasValue(oForm, "MTXSIZE", "@FIL_DR_SMPLSIZE", "U_SIZECODE");
                AddLineIfLastRowHasValue(oForm, "MTXBUYER", "@FIL_DR_SMPLBUYER", "U_CARDCODE");

                oForm.Freeze(false);
            }
            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {

                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

                // Disable StyleId field
                SAPbouiCOM.Item oSampleCode = oForm.Items.Item("ETSLCODE");
                if (oSampleCode.Enabled)
                    oSampleCode.Enabled = false;

                // Add new row if last row has data
                AddLineIfLastRowHasValue(oForm, "MTXCOLOR", "@FIL_DR_SMPLCOLO", "U_COLOCODE");
                AddLineIfLastRowHasValue(oForm, "MTXSIZE", "@FIL_DR_SMPLSIZE", "U_SIZECODE");
                AddLineIfLastRowHasValue(oForm, "MTXBUYER", "@FIL_DR_SMPLBUYER", "U_CARDCODE");

                // Enable/disable other buttons based on matrix
                SampleEnableButtons(ref oForm);

                // --- Only check for changes if we already had previous states ---
                bool newCheckedInSztType = false;
                bool newCheckedInColor = false;

                if (prevSztTypeCheckboxStates.Count > 0)
                    newCheckedInSztType = HasCheckboxChanges(oForm, "MTXSIZE", "CLSZAPL", prevSztTypeCheckboxStates);

                if (prevColorCheckboxStates.Count > 0)
                    newCheckedInColor = HasCheckboxChanges(oForm, "MTXCOLOR", "CLCLRAPL", prevColorCheckboxStates);

                if (newCheckedInSztType || newCheckedInColor)
                {
                    SAPbouiCOM.Item oBtnItmTx = oForm.Items.Item("BTNITMTX");
                    oBtnItmTx.Enabled = true;

                    Application.SBO_Application.StatusBar.SetText(
                        "Detected new checked rows in matrix — BTNITMTX enabled.",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Success
                    );
                }
                else
                {
                    //Disable if no new checks
                    SAPbouiCOM.Item oBtnItmTx = oForm.Items.Item("BTNITMTX");
                    oBtnItmTx.Enabled = false;
                }
                // --- Always refresh states AFTER processing ---
                CaptureCheckboxStates(oForm, "MTXSIZE", "CLSZAPL", prevSztTypeCheckboxStates);
                CaptureCheckboxStates(oForm, "MTXCOLOR", "CLCLRAPL", prevColorCheckboxStates);
            }

        }

        // --- Generic reusable method to store checkbox states ---
        private void CaptureCheckboxStates(SAPbouiCOM.Form oForm, string matrixId, string columnId, Dictionary<int, bool> targetDict)
        {
            try
            {
                targetDict.Clear();
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixId).Specific;

                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    SAPbouiCOM.CheckBox chk = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item(columnId).Cells.Item(i).Specific;
                    targetDict[i] = chk.Checked;
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText($"Error capturing states for {matrixId}: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private bool HasCheckboxChanges(SAPbouiCOM.Form oForm, string matrixId, string columnId, Dictionary<int, bool> prevStates)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixId).Specific;

                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    SAPbouiCOM.CheckBox chk = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item(columnId).Cells.Item(i).Specific;
                    bool prevState = prevStates.ContainsKey(i) ? prevStates[i] : false;

                    // Detect any change: checked → unchecked OR unchecked → checked
                    if (chk.Checked != prevState)
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Error checking checkbox changes for {matrixId}: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            return false;
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
        private void SampleEnableButtons(ref SAPbouiCOM.Form oForm)
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
        private void DELBTN_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.DBDataSource DBDataSourceLine = oForm.DataSources.DBDataSources.Item("@FIL_DR_SMPLATACH");
            SAPbouiCOM.Matrix MTXATTCH = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXATTCH").Specific;


            MTXATTCH.FlushToDataSource();
            for (int i = 1; i <= MTXATTCH.RowCount; i++)
            {
                if (MTXATTCH.IsRowSelected(i))
                {
                    int rowIndex = i - 1;

                    if (rowIndex >= 0 && rowIndex < DBDataSourceLine.Size)
                    {
                        DBDataSourceLine.RemoveRecord(rowIndex);
                        for (int j = 0; j < DBDataSourceLine.Size; j++)
                        {
                            DBDataSourceLine.Offset = j;
                            DBDataSourceLine.SetValue("LineId", j, (j + 1).ToString());
                        }
                        MTXATTCH.LoadFromDataSource();
                        Application.SBO_Application.MessageBox("Selected row deleted.");
                    }
                    else
                    {
                        Application.SBO_Application.MessageBox("Invalid row index.");
                    }

                    break;
                }
            }

        }

        private void DISPBTN_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.Matrix MTXATTCH = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXATTCH").Specific;

            for (int i = 1; i <= MTXATTCH.RowCount; i++)
            {
                if (MTXATTCH.IsRowSelected(i))
                {
                    string filePath = ((SAPbouiCOM.EditText)MTXATTCH.Columns.Item("CLATTACH").Cells.Item(i).Specific).Value;
                    if (!string.IsNullOrEmpty(filePath) && System.IO.File.Exists(filePath))
                    {
                        System.Diagnostics.Process.Start(filePath);
                    }
                    else
                    {
                        Application.SBO_Application.MessageBox("File does not exist or path is empty.");
                    }
                    break;
                }
            }

        }

        private void BRWSBTN_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.DBDataSource DBDataSourceLine = oForm.DataSources.DBDataSources.Item("@FIL_DR_SMPLATACH");
            SAPbouiCOM.Matrix MTXATTCH = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXATTCH").Specific;


            string filePath = FileDialogHelper.ShowFileDialog();
            if (!string.IsNullOrEmpty(filePath))
            {
                int lastRow = MTXATTCH.VisualRowCount;
                bool needNewRow = (lastRow == 0) ||
                                  !string.IsNullOrEmpty(((SAPbouiCOM.EditText)MTXATTCH.Columns.Item("CLATTACH").Cells.Item(lastRow).Specific).Value);

                if (needNewRow)
                {
                    Global.GFunc.SetNewLine(MTXATTCH, DBDataSourceLine, 1, "");
                    lastRow = MTXATTCH.VisualRowCount;
                }

                ((SAPbouiCOM.EditText)MTXATTCH.Columns.Item("CLATTACH").Cells.Item(lastRow).Specific).Value = filePath;
                MTXATTCH.FlushToDataSource();
            }

        }
        private void MTXSIZE_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_SIZE")
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(cflUID);
                    SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                    SAPbouiCOM.Condition oCon1 = oCons.Add();
                    oCon1.Alias = "U_ACTIVE";
                    oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon1.CondVal = "Y";
                    oCFL.SetConditions(oCons);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error filtering Size CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }

        }

        private void MTXBUYER_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (!pVal.ActionSuccess)
                    return;

                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;

                SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
                if (dt == null || dt.Rows.Count == 0)
                    return;

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXBUYER").Specific;
                SAPbouiCOM.DBDataSource DBDataSourceLine = oForm.DataSources.DBDataSources.Item("@FIL_DR_SMPLBUYER");
                int row = pVal.Row;

                // 🔹 CFL 1 : Business Partner (OCRD)
                if (cflArg.ChooseFromListUID == "CFL_OCRD2")
                {
                    string cardCode = dt.GetValue("CardCode", 0).ToString().Trim();
                    string cardName = dt.GetValue("CardName", 0).ToString().Trim();
                    oMatrix.SetCellWithoutValidation(row, "CLBUYCOD", cardCode);
                    oMatrix.SetCellWithoutValidation(row, "CLBUYNAM", cardName);

                    // Add new row if last row has data
                    int lastRow = oMatrix.RowCount;
                    bool lastRowHasData = !string.IsNullOrWhiteSpace(((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLBUYCOD").Cells.Item(lastRow).Specific).Value);
                    if (pVal.Row == lastRow && lastRowHasData)
                    {
                        Global.GFunc.SetNewLine(oMatrix, DBDataSourceLine, 1, "");
                    }
                }
                // 🔹 CFL 2 : Status CFL
                else if (cflArg.ChooseFromListUID == "CFL_STAT")
                {
                    string statusCode = dt.GetValue("Code", 0).ToString().Trim();
                    oMatrix.SetCellWithoutValidation(row, "CLSMPSTA", statusCode);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar
                    .SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short,
                             SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }


        private void MTXCOLOR_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_CLOR")
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(cflUID);
                    SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                    SAPbouiCOM.Condition oCon1 = oCons.Add();
                    oCon1.Alias = "U_ACTIVE";
                    oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon1.CondVal = "Y";
                    oCFL.SetConditions(oCons);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error filtering Size CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }

        }

        private void MTXCOLOR_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXCOLOR").Specific;
                SAPbouiCOM.DBDataSource DBDataSourceLine = oForm.DataSources.DBDataSources.Item("@FIL_DR_SMPLCOLO");
                SAPbouiCOM.DataTable oDataTable = cflArg.SelectedObjects;

                if (oDataTable.Rows.Count > 0)
                {
                    string code = oDataTable.GetValue("Code", 0).ToString();
                    string name = oDataTable.GetValue("Name", 0).ToString();
                    string pantone = oDataTable.GetValue("U_PANTONE", 0).ToString();
                    int row = pVal.Row;
                    //Set Values
                    oMatrix.SetCellWithoutValidation(row, "CLCLRCOD", code);
                    oMatrix.SetCellWithoutValidation(row, "CLCLRNAM", name);
                    oMatrix.SetCellWithoutValidation(row, "CLPANTON", pantone);
                    oMatrix.FlushToDataSource();

                    // Add new row if last row has data
                    int lastRow = oMatrix.RowCount;
                    bool lastRowHasData = !string.IsNullOrWhiteSpace(((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLCLRCOD").Cells.Item(lastRow).Specific).Value);
                    if (pVal.Row == lastRow && lastRowHasData)
                    {
                        Global.GFunc.SetNewLine(oMatrix, DBDataSourceLine, 1, "");
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Color Matrix CFL Error: " + ex.Message,
                   SAPbouiCOM.BoMessageTime.bmt_Short,
                   SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void MTXSIZE_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXSIZE").Specific;
                SAPbouiCOM.DBDataSource DBDataSourceLine = oForm.DataSources.DBDataSources.Item("@FIL_DR_SMPLSIZE");
                SAPbouiCOM.DataTable oDataTable = cflArg.SelectedObjects;

                if (oDataTable.Rows.Count > 0)
                {
                    string code = oDataTable.GetValue("Code", 0).ToString();
                    string name = oDataTable.GetValue("Name", 0).ToString();

                    int row = pVal.Row; // matrix row where CFL triggered
                    //Set Values
                    oMatrix.SetCellWithoutValidation(row, "CLSZCODE", code);
                    oMatrix.SetCellWithoutValidation(row, "CLSZNAME", name);
                    oMatrix.FlushToDataSource();

                    // Add new row if last row has data
                    int lastRow = oMatrix.RowCount;
                    bool lastRowHasData = !string.IsNullOrWhiteSpace(((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLSZCODE").Cells.Item(lastRow).Specific).Value);
                    if (pVal.Row == lastRow && lastRowHasData)
                    {
                        Global.GFunc.SetNewLine(oMatrix, DBDataSourceLine, 1, "");
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Size Matrix CFL Error: " + ex.Message,
                   SAPbouiCOM.BoMessageTime.bmt_Short,
                   SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        private void ETCRDCOD_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("CardCode", 0).ToString().Trim();
            string Name = dt.GetValue("CardName", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETCRDCOD").Specific;
            ETCD.Value = Code;
            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETCRDNAM").Specific;
            ETNM.Value = Name;

            //Matrix Open 
            EnsureLine(oForm, "MTXSIZE", "@FIL_DR_SMPLSIZE");
            EnsureLine(oForm, "MTXCOLOR", "@FIL_DR_SMPLCOLO");
            //EnsureLine(oForm, "MTXBUYER", "@FIL_DR_SMPLBUYER");


            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXBUYER").Specific;
            SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("@FIL_DR_SMPLBUYER");

            if (oMatrix.RowCount == 0)
            {
                oMatrix.AddRow();
            }
            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("#").Cells.Item(1).Specific).Value = "1";
            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLBUYCOD").Cells.Item(1).Specific).Value = Code;
            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLBUYNAM").Cells.Item(1).Specific).Value = Name;

            oMatrix.FlushToDataSource();
            oMatrix.LoadFromDataSource();
            oMatrix.AutoResizeColumns();

            int lastRow = oMatrix.RowCount;
            bool lastRowHasData = !string.IsNullOrWhiteSpace(
                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLBUYCOD").Cells.Item(lastRow).Specific).Value);

            if (lastRowHasData)
                Global.GFunc.SetNewLine(oMatrix, ds, 1, "");

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

        private void ETMERCOD_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("empID", 0).ToString().Trim();
            string Name = dt.GetValue("U_FNAME", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETMERCOD").Specific;
            ETCD.Value = Code;
            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETMERNAM").Specific;
            ETNM.Value = Name;

        }


        private void ETRUTSTG_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("Code", 0).ToString().Trim();
            string Name = dt.GetValue("Name", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETRUTSTG").Specific;
            ETCD.Value = Code;
            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETRUTSNM").Specific;
            ETNM.Value = Name;

        }



        private void ETITMGRP_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_SITM")
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(cflUID);
                    SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                    SAPbouiCOM.Condition oCon1 = oCons.Add();
                    oCon1.Alias = "U_ITMSGRPTP";
                    oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon1.CondVal = "FG";
                    oCFL.SetConditions(oCons);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error filtering SAP Group CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }

        }

        private void ETITMGRP_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("ItmsGrpCod", 0).ToString().Trim();
            string Name = dt.GetValue("ItmsGrpNam", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETITMGRP").Specific;
            ETCD.Value = Code;
            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETITMGNM").Specific;
            ETNM.Value = Name;

        }
        private void ETSLTYPE_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_SMTY")
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(cflUID);
                    SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                    SAPbouiCOM.Condition oCon1 = oCons.Add();
                    oCon1.Alias = "U_ACTIVE";
                    oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon1.CondVal = "Y";
                    oCFL.SetConditions(oCons);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error filtering Sample Type CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }

        }

        private void ETSLTYPE_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("Code", 0).ToString().Trim();
            string Name = dt.GetValue("Name", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETSLTYPE").Specific;
            ETCD.Value = Code;
            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETSLTYNM").Specific;
            ETNM.Value = Name;

        }
        private void ETGENDER_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("Code", 0).ToString();
            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETGENDER").Specific;
            ETCD.Value = Code;
        }

        private void ETGENDER_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_GEN")
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(cflUID);
                    SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                    SAPbouiCOM.Condition oCon1 = oCons.Add();
                    oCon1.Alias = "U_ACTIVE";
                    oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon1.CondVal = "Y";
                    oCFL.SetConditions(oCons);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error filtering Gender CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }

        }
    }
}
