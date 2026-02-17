using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Apparel_Dynamic_1._0.Helper;

namespace Apparel_Dynamic_1._0.Resources.Setup
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Setup.SampleStatus", "Resources/Setup/SampleStatus.b1f")]
    class SampleStatus : UserFormBase
    {
        public SampleStatus()
        {
        }

        private SAPbouiCOM.StaticText STCODE, STACTIVE;
        private SAPbouiCOM.EditText ETCODE, ETDOCTRY;
        private SAPbouiCOM.Button ADDButton, CancelButton;
        private SAPbouiCOM.CheckBox CKACTIVE;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.STCODE = ((SAPbouiCOM.StaticText)(this.GetItem("STCODE").Specific));
            this.ETCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETCODE").Specific));
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.ADDButton.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.ADDButton_PressedBefore);
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            this.STACTIVE = ((SAPbouiCOM.StaticText)(this.GetItem("STACTIVE").Specific));
            this.CKACTIVE = ((SAPbouiCOM.CheckBox)(this.GetItem("CKACTIVE").Specific));
            this.OnCustomInitialize();
        }

        public override void OnInitializeFormEvents()
        {
        }


        private void OnCustomInitialize()
        {

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
            try
            {
                SAPbouiCOM.EditText oCode =
                    (SAPbouiCOM.EditText)oForm.Items.Item("ETCODE").Specific;

                string code = oCode.Value.Trim().ToUpper();
                oCode.Value = code;

                // 🔍 Empty check
                if (string.IsNullOrEmpty(code))
                {
                    Global.GFunc.ShowError("Enter Size Code");
                    oForm.ActiveItem = "ETCODE";
                    return BubbleEvent = false;
                }

                // 🔍 Format validation (no spaces + no special characters)
                if (!IsValidCode(code, out string err))
                {
                    Global.GFunc.ShowError(err);
                    oCode.Value = "";
                    oForm.ActiveItem = "ETCODE";
                    return BubbleEvent = false;
                }

                // 🔍 Duplicate check
                SAPbobsCOM.Recordset oRS =
                    (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query;

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    query = $@"
                    SELECT 1 
                    FROM ""@FIL_MH_SMPLSTAT""
                    WHERE ""Code"" = '{code.Replace("'", "''")}'";
                }
                else // UPDATE MODE
                {
                    string docEntry =
                        oForm.DataSources.DBDataSources.Item("@FIL_MH_SMPLSTAT")
                        .GetValue("DocEntry", 0).Trim();

                    query = $@"
                    SELECT 1 
                    FROM ""@FIL_MH_SMPLSTAT""
                    WHERE ""Code"" = '{code.Replace("'", "''")}'
                    AND ""DocEntry"" <> '{docEntry}'";
                }

                oRS.DoQuery(query);

                if (!oRS.EoF)
                {
                    Global.GFunc.ShowError("Code already exists!");
                    oCode.Value = "";
                    oForm.ActiveItem = "ETCODE";
                    return BubbleEvent = false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Global.GFunc.ShowError(ex.Message);
                return BubbleEvent = false;
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
    }
}
