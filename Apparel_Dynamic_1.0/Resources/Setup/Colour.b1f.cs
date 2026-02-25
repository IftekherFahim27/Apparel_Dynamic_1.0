using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Apparel_Dynamic_1._0.Helper;

namespace Apparel_Dynamic_1._0.Resources.Setup
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Setup.Colour", "Resources/Setup/Colour.b1f")]
    class Colour : UserFormBase
    {
        public Colour()
        {
        }

        private SAPbouiCOM.StaticText STCODE, STNAME, STACTIVE, STPANTON;
        private SAPbouiCOM.EditText ETCODE, ETNAME, ETDOCTRY, ETPANTON;
        private SAPbouiCOM.Button ADDButton, CancelButton;
        private SAPbouiCOM.CheckBox CKACTIVE;



        public override void OnInitializeComponent()
        {
            this.STCODE = ((SAPbouiCOM.StaticText)(this.GetItem("STCODE").Specific));
            this.STNAME = ((SAPbouiCOM.StaticText)(this.GetItem("STNAME").Specific));
            this.ETCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETCODE").Specific));
            this.ETCODE.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.ETCODE_LostFocusAfter);
            this.ETNAME = ((SAPbouiCOM.EditText)(this.GetItem("ETNAME").Specific));
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.ADDButton.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.ADDButton_PressedBefore);
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            this.STACTIVE = ((SAPbouiCOM.StaticText)(this.GetItem("STACTIVE").Specific));
            this.CKACTIVE = ((SAPbouiCOM.CheckBox)(this.GetItem("CKACTIVE").Specific));
            this.STPANTON = ((SAPbouiCOM.StaticText)(this.GetItem("STPANTON").Specific));
            this.ETPANTON = ((SAPbouiCOM.EditText)(this.GetItem("ETPANTON").Specific));
            this.ETPANTON.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.ETPANTON_LostFocusAfter);
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

        private void ADDButton_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                ValidateForm(ref oForm, ref BubbleEvent);
            }

        }
        private void ETPANTON_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    // Get the entered Code
                    string pcode = ((SAPbouiCOM.EditText)oForm.Items.Item("ETPANTON").Specific).Value.Trim();
                    string UpCode = Global.GFunc.ToUpperCase(pcode);
                    ((SAPbouiCOM.EditText)oForm.Items.Item("ETPANTON").Specific).Value = UpCode;

                    if (!string.IsNullOrEmpty(UpCode))
                    {
                        // 🔍 Validate the code
                        if (!IsValidCode(UpCode, out string err))
                        {
                            Application.SBO_Application.StatusBar.SetText(err,
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                            ((SAPbouiCOM.EditText)oForm.Items.Item("ETPANTON").Specific).Value = "";
                            return;
                        }

                        // 🔍 Check duplicate
                        SAPbobsCOM.Recordset oRS =
                            (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        string query = $@"SELECT 1 FROM ""@FIL_MH_OCOM"" WHERE ""U_PANTONE"" = '{UpCode.Replace("'", "''")}'";
                        oRS.DoQuery(query);

                        if (!oRS.EoF)
                        {
                            Application.SBO_Application.StatusBar.SetText("Pantone Code already exists!",
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                            ((SAPbouiCOM.EditText)oForm.Items.Item("ETPANTON").Specific).Value = "";
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
        private void ETCODE_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    // Get the entered Code
                    string code = ((SAPbouiCOM.EditText)oForm.Items.Item("ETCODE").Specific).Value.Trim();
                    string UCode = Global.GFunc.ToUpperCase(code);
                    ((SAPbouiCOM.EditText)oForm.Items.Item("ETCODE").Specific).Value = UCode;

                    if (!string.IsNullOrEmpty(UCode))
                    {
                        // 🔍 Validate the code
                        if (!IsValidCode(UCode, out string err))
                        {
                            Application.SBO_Application.StatusBar.SetText(err,
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                            ((SAPbouiCOM.EditText)oForm.Items.Item("ETCODE").Specific).Value = "";
                            return;
                        }

                        // 🔍 Check duplicate
                        SAPbobsCOM.Recordset oRS =
                            (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        string query = $@"SELECT 1 FROM ""@FIL_MH_OCOM"" WHERE ""Code"" = '{UCode.Replace("'", "''")}'";
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
            string code = oForm.DataSources.DBDataSources.Item("@FIL_MH_OCOM").GetValue("Code", 0);
            string name = oForm.DataSources.DBDataSources.Item("@FIL_MH_OCOM").GetValue("Name", 0);
            string pantone = oForm.DataSources.DBDataSources.Item("@FIL_MH_OCOM").GetValue("U_PANTONE", 0);
            if (code == "")
            {
                Global.GFunc.ShowError("Enter Colour Master Code");
                oForm.ActiveItem = "ETCODE";
                return BubbleEvent = false;
            }
            else if (pantone == "")
            {
                Global.GFunc.ShowError("Enter Panotne Name");
                oForm.ActiveItem = "ETPANTON";
                return BubbleEvent = false;
            }
            else if (name == "")
            {
                Global.GFunc.ShowError("Enter Colour Master Name");
                oForm.ActiveItem = "ETNAME";
                return BubbleEvent = false;
            }
            return BubbleEvent;
        }
    }
}
