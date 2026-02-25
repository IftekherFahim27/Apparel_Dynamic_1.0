using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Apparel_Dynamic_1._0.Helper;

namespace Apparel_Dynamic_1._0.Resources.Setup
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Setup.ComponentStages", "Resources/Setup/ComponentStages.b1f")]
    class ComponentStages : UserFormBase
    {
        public ComponentStages()
        {
        }

        private SAPbouiCOM.StaticText STCODE, STNAME, STACTIVE, STUOMAPL, STUOM;
        private SAPbouiCOM.EditText ETCODE, ETNAME, ETDOCTRY, ETUOM;
        private SAPbouiCOM.Button ADDButton, CancelButton;
        private SAPbouiCOM.ComboBox CBUOMAPL;



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
            this.STUOMAPL = ((SAPbouiCOM.StaticText)(this.GetItem("STUOMAPL").Specific));
            this.CBUOMAPL = ((SAPbouiCOM.ComboBox)(this.GetItem("CBUOMAPL").Specific));
            this.CBUOMAPL.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.CBUOMAPL_ComboSelectAfter);
            this.STUOM = ((SAPbouiCOM.StaticText)(this.GetItem("STUOM").Specific));
            this.ETUOM = ((SAPbouiCOM.EditText)(this.GetItem("ETUOM").Specific));
            this.CKACTIVE = ((SAPbouiCOM.CheckBox)(this.GetItem("CKACTIVE").Specific));
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
                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBUOMAPL").Specific;

                SAPbouiCOM.Item oUomItem = oForm.Items.Item("ETUOM");

                string selectedValue = oCombo.Selected.Value;

                if (selectedValue == "Y")
                {
                    oUomItem.Enabled = true;
                    oForm.ActiveItem = "ETUOM";

                    SAPbouiCOM.StaticText oLabel = (SAPbouiCOM.StaticText)oForm.Items.Item("STUOM").Specific;
                    oLabel.Caption = "UoM*";
                }
                else
                {
                    oUomItem.Enabled = false;
                    ((SAPbouiCOM.EditText)oUomItem.Specific).Value = "";
                    SAPbouiCOM.StaticText oLabel = (SAPbouiCOM.StaticText)oForm.Items.Item("STUOM").Specific;
                    oLabel.Caption = "UoM";
                }
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

        private void CBUOMAPL_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBUOMAPL").Specific;

            SAPbouiCOM.Item oUomItem = oForm.Items.Item("ETUOM");

            string selectedValue = oCombo.Selected.Value;

            if (selectedValue == "Y")
            {
                oUomItem.Enabled = true;
                oForm.ActiveItem = "ETUOM";

                SAPbouiCOM.StaticText oLabel =(SAPbouiCOM.StaticText)oForm.Items.Item("STUOM").Specific;
                oLabel.Caption = "UoM*";
            }
            else
            {
                oUomItem.Enabled = false;
                ((SAPbouiCOM.EditText)oUomItem.Specific).Value = "";
                SAPbouiCOM.StaticText oLabel = (SAPbouiCOM.StaticText)oForm.Items.Item("STUOM").Specific;
                oLabel.Caption = "UoM";
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
                        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string query = $@"SELECT 1 FROM ""@FIL_MH_CMPNSTAG"" WHERE ""Code"" = '{UCode.Replace("'", "''")}'";
                        oRS.DoQuery(query);
                        if (!oRS.EoF)
                        {
                            Application.SBO_Application.StatusBar.SetText("Code already exists!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            ((SAPbouiCOM.EditText)oForm.Items.Item("ETCODE").Specific).Value = "";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private bool ValidateForm(ref SAPbouiCOM.Form oForm, ref bool BubbleEvent)
        {
            string code = oForm.DataSources.DBDataSources.Item("@FIL_MH_CMPNSTAG").GetValue("Code", 0);
            string name = oForm.DataSources.DBDataSources.Item("@FIL_MH_CMPNSTAG").GetValue("Name", 0);
            string uomApl = oForm.DataSources.DBDataSources.Item("@FIL_MH_CMPNSTAG").GetValue("U_UOMAPPLY", 0);

            if (code == "")
            {
                Global.GFunc.ShowError("Enter Stage Code");
                oForm.ActiveItem = "ETCODE";
                return BubbleEvent = false;
            }
            else if (name == "")
            {
                Global.GFunc.ShowError("Enter Stage Name");
                oForm.ActiveItem = "ETNAME";
                return BubbleEvent = false;
            }

            if (uomApl == "Y")
            {
                string uom = oForm.DataSources.DBDataSources.Item("@FIL_MH_CMPNSTAG").GetValue("U_UOM", 0);
                if (uom == "")
                {
                    Global.GFunc.ShowError("Enter Uom");
                    oForm.ActiveItem = "ETUOM";
                    return BubbleEvent = false;
                }
            }

            return BubbleEvent;
        }
    }
}
