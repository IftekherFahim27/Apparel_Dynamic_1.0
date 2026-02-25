using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Apparel_Dynamic_1._0.Helper;

namespace Apparel_Dynamic_1._0.Resources.Setup
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Setup.ProductGroup", "Resources/Setup/ProductGroup.b1f")]
    class ProductGroup : UserFormBase
    {
        public ProductGroup()
        {
        }


        private SAPbouiCOM.StaticText STCODE, STNAME, STPDTYPE, STPDLINE;
        private SAPbouiCOM.EditText ETCODE, ETNAME, ETDOCTRY, ETPDTYPE, ETPDLINE;
        private SAPbouiCOM.Button ADDButton, CancelButton;
        private SAPbouiCOM.CheckBox CKACTIVE;
        public override void OnInitializeComponent()
        {
            this.STCODE = ((SAPbouiCOM.StaticText)(this.GetItem("STCODE").Specific));
            this.STNAME = ((SAPbouiCOM.StaticText)(this.GetItem("STNAME").Specific));
            this.ETCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETCODE").Specific));
            this.ETCODE.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.ETCODE_LostFocusAfter);
            this.ETNAME = ((SAPbouiCOM.EditText)(this.GetItem("ETNAME").Specific));
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.ADDButton.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.ADDButton_PressedBefore);
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.CKACTIVE = ((SAPbouiCOM.CheckBox)(this.GetItem("CKACTIVE").Specific));
            this.STPDTYPE = ((SAPbouiCOM.StaticText)(this.GetItem("STPDTYPE").Specific));
            this.STPDLINE = ((SAPbouiCOM.StaticText)(this.GetItem("STPDLINE").Specific));
            this.ETPDTYPE = ((SAPbouiCOM.EditText)(this.GetItem("ETPDTYPE").Specific));
            this.ETPDTYPE.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETPDTYPE_ChooseFromListAfter);
            this.ETPDTYPE.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETPDTYPE_ChooseFromListBefore);
            this.ETPDLINE = ((SAPbouiCOM.EditText)(this.GetItem("ETPDLINE").Specific));
            this.ETPDLINE.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETPDLINE_ChooseFromListAfter);
            this.ETPDLINE.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETPDLINE_ChooseFromListBefore);
            this.OnCustomInitialize();

        }

        public override void OnInitializeFormEvents()
        {
            this.DataLoadAfter += new DataLoadAfterHandler(this.Form_DataLoadAfter);

        }



        private void OnCustomInitialize()
        {

        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            oForm.Freeze(true);
            try
            {
                SetItemsEnabled(oForm, false, "ETCODE");
            }
            finally
            {
                oForm.Freeze(false);
            }

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

        private void ETPDTYPE_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("Code", 0).ToString();
            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETPDTYPE").Specific;
            ETCD.Value = Code;

        }

        private void ETPDLINE_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_LINE")
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
                    "Error filtering Product Line CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }

        }

        private void ETPDLINE_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("Code", 0).ToString();
            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETPDLINE").Specific;
            ETCD.Value = Code;

        }

        private void ETPDTYPE_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_TYPE")
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
                    "Error filtering Product Type CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
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

        private void ETCODE_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {

                    string code = ((SAPbouiCOM.EditText)oForm.Items.Item("ETCODE").Specific).Value.Trim();
                    string UCode = Global.GFunc.ToUpperCase(code);
                    ((SAPbouiCOM.EditText)oForm.Items.Item("ETCODE").Specific).Value = UCode;

                    if (!string.IsNullOrEmpty(UCode))
                    {

                        if (!IsValidCode(UCode, out string err))
                        {
                            Application.SBO_Application.StatusBar.SetText(err,
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                            ((SAPbouiCOM.EditText)oForm.Items.Item("ETCODE").Specific).Value = "";
                            return;
                        }

                        SAPbobsCOM.Recordset oRS =
                            (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        string query = $@"SELECT 1 FROM ""@FIL_MH_OPGM"" WHERE ""Code"" = '{UCode.Replace("'", "''")}'";
                        oRS.DoQuery(query);

                        if (!oRS.EoF)
                        {
                            Application.SBO_Application.StatusBar.SetText("Code already exists!",
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                            ((SAPbouiCOM.EditText)oForm.Items.Item("ETCODE").Specific).Value = "";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }


        private bool IsValidCode(string code, out string errorMessage)
        {
            errorMessage = "";

            if (string.IsNullOrWhiteSpace(code))
            {
                errorMessage = "Code cannot be empty.";
                return false;
            }

            if (code.Contains(" "))
            {
                errorMessage = "Code cannot contain spaces.";
                return false;
            }

            if (!System.Text.RegularExpressions.Regex.IsMatch(code, @"^[A-Za-z0-9]+$"))
            {
                errorMessage = "Code contains invalid characters. Only letters and numbers are allowed.";
                return false;
            }

            return true;
        }


        private bool ValidateForm(ref SAPbouiCOM.Form oForm, ref bool BubbleEvent)
        {
            string code = oForm.DataSources.DBDataSources.Item("@FIL_MH_OPGM").GetValue("Code", 0);
            string name = oForm.DataSources.DBDataSources.Item("@FIL_MH_OPGM").GetValue("Name", 0);
            string type = oForm.DataSources.DBDataSources.Item("@FIL_MH_OPGM").GetValue("U_PRDTYPE", 0);
            string line = oForm.DataSources.DBDataSources.Item("@FIL_MH_OPGM").GetValue("U_PRDLINE", 0);

            if (code == "")
            {
                Global.GFunc.ShowError("Enter Product Group Mater Code");
                oForm.ActiveItem = "ETCODE";
                return BubbleEvent = false;
            }
            else if (name == "")
            {
                Global.GFunc.ShowError("Enter Product Group Mater Name");
                oForm.ActiveItem = "ETNAME";
                return BubbleEvent = false;
            }
            else if (type == "")
            {
                Global.GFunc.ShowError("Enter Product Type Code");
                oForm.ActiveItem = "ETPDTYPE";
                return BubbleEvent = false;
            }
            else if (line == "")
            {
                Global.GFunc.ShowError("Enter Product Line Code");
                oForm.ActiveItem = "ETPDLINE";
                return BubbleEvent = false;
            }
            return BubbleEvent;
        }

    }
}
