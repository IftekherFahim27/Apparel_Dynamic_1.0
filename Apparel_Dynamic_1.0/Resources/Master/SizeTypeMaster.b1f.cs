using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Apparel_Dynamic_1._0.Helper;

namespace Apparel_Dynamic_1._0.Resources.Master
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Master.SizeTypeMaster", "Resources/Master/SizeTypeMaster.b1f")]
    class SizeTypeMaster : UserFormBase
    {
        public SizeTypeMaster()
        {
        }

        public override void OnInitializeComponent()
        {
            this.STCODE = ((SAPbouiCOM.StaticText)(this.GetItem("STCODE").Specific));
            this.STNAME = ((SAPbouiCOM.StaticText)(this.GetItem("STNAME").Specific));
            this.ETCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETCODE").Specific));
            this.ETCODE.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.ETCODE_LostFocusAfter);
            this.ETNAME = ((SAPbouiCOM.EditText)(this.GetItem("ETNAME").Specific));
            this.CKACTIVE = ((SAPbouiCOM.CheckBox)(this.GetItem("CKACTIVE").Specific));
            this.MTXSIZE = ((SAPbouiCOM.Matrix)(this.GetItem("MTXSIZE").Specific));
            this.MTXSIZE.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.MTXSIZE_ChooseFromListAfter);
            this.MTXSIZE.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.MTXSIZE_ChooseFromListBefore);
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.ADDButton.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.ADDButton_PressedBefore);
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            this.OnCustomInitialize();

        }

        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.StaticText STCODE, STNAME;
        private SAPbouiCOM.EditText ETCODE, ETNAME, ETDOCTRY;
        private SAPbouiCOM.CheckBox CKACTIVE;
        private SAPbouiCOM.Matrix MTXSIZE;
        private SAPbouiCOM.Button ADDButton, CancelButton;



        private void OnCustomInitialize()
        {

        }

        private void MTXSIZE_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {

                if (pVal.ColUID != "CLSZCODE") return;
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Matrix oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXSIZE").Specific;
                string colCode = "CLSZCODE";

                HashSet<string> usedCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                for (int i = 1; i <= oMtx.RowCount; i++)
                {
                    var cell = (SAPbouiCOM.EditText)oMtx.Columns.Item(colCode).Cells.Item(i).Specific;
                    string code = (cell.Value ?? "").Trim();

                    if (!string.IsNullOrEmpty(code))
                        usedCodes.Add(code);
                }


                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item("CFL_SIZE");
                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                if (usedCodes.Count == 0)
                {
                    oCFL.SetConditions(oCons);
                    return;
                }

                int idx = 0;
                foreach (string c in usedCodes)
                {
                    SAPbouiCOM.Condition cond = oCons.Add();
                    cond.Alias = "Code";
                    cond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                    cond.CondVal = c;

                    idx++;
                    if (idx < usedCodes.Count)
                        cond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                }
                oCFL.SetConditions(oCons);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Size CFL Error: " + ex.Message,
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
                SAPbouiCOM.DBDataSource DBDataSourceLine = oForm.DataSources.DBDataSources.Item("@FIL_MR_STM1");
                SAPbouiCOM.DataTable oDataTable = cflArg.SelectedObjects;

                if (oDataTable == null || oDataTable.Rows.Count == 0)
                    return;
                
                    string code = oDataTable.GetValue("Code", 0).ToString();
                    string name = oDataTable.GetValue("Name", 0).ToString();
                    int row = pVal.Row;
                    oMatrix.SetCellWithoutValidation(row, "CLSZCODE", code);
                    oMatrix.SetCellWithoutValidation(row, "CLSZNAME", name);
                    oMatrix.FlushToDataSource();

                    int lastRow = oMatrix.RowCount;
                    bool lastRowHasData = !string.IsNullOrWhiteSpace(((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLSZCODE").Cells.Item(lastRow).Specific).Value);
                    if (pVal.Row == lastRow && lastRowHasData)
                    {
                        Global.GFunc.SetNewLine(oMatrix, DBDataSourceLine, 1, "");
                    }

                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }
                
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Size CFL Error: " + ex.Message,
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

                        string query = $@"SELECT 1 FROM ""@FIL_MH_OSTM"" WHERE ""Code"" = '{UCode.Replace("'", "''")}'";
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

        private void ADDButton_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                ValidateForm(ref oForm, ref BubbleEvent);
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
            string code = oForm.DataSources.DBDataSources.Item("@FIL_MH_OSTM").GetValue("Code", 0);
            string name = oForm.DataSources.DBDataSources.Item("@FIL_MH_OSTM").GetValue("Name", 0);
            if (code == "")
            {
                Global.GFunc.ShowError("Enter Size Type Master Code");
                oForm.ActiveItem = "ETCODE";
                return BubbleEvent = false;
            }
            else if (name == "")
            {
                Global.GFunc.ShowError("Enter Size Type Master Description");
                oForm.ActiveItem = "ETNAME";
                return BubbleEvent = false;
            }

            PreventEmptyLastRow(oForm, "@FIL_MR_STM1", MTXSIZE, "U_SIZECODE");

            if (!RequireAtLeastOneFilledRow(oForm,"@FIL_MR_STM1",MTXSIZE,"U_SIZECODE",ref BubbleEvent))
            {
                EnsureLine(oForm, "MTXSIZE", "@FIL_MR_STM1");
                return BubbleEvent;
            }

            return BubbleEvent;
        }

        private bool RequireAtLeastOneFilledRow(SAPbouiCOM.Form oForm,string dbDatasourceUID,SAPbouiCOM.Matrix matrix,string columnName,ref bool BubbleEvent)
        {
            try
            {
                SAPbouiCOM.DBDataSource oDB =oForm.DataSources.DBDataSources.Item(dbDatasourceUID);
                int rowCount = matrix.VisualRowCount;
                if (rowCount == 0)
                {
                    Global.GFunc.ShowError("At least one Size Code is required.");
                    oForm.ActiveItem = matrix.Item.UniqueID;
                    BubbleEvent = false;
                    return false;
                }

                for (int i = 0; i < rowCount; i++)
                {
                    string val = oDB.GetValue(columnName, i)
                                    .Replace("\0", "")
                                    .Trim();

                    if (!string.IsNullOrEmpty(val) && val != "0.0")
                        return true;
                }

                Global.GFunc.ShowError("At least one Size Code is required.");
                oForm.ActiveItem = matrix.Item.UniqueID;

                BubbleEvent = false;
                return false;
            }
            catch (Exception ex)
            {
                Global.GFunc.ShowError("Validation error: " + ex.Message);
                BubbleEvent = false;
                return false;
            }
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

        public static void EnsureLine(SAPbouiCOM.Form oForm, string matrixID, string dbTable)
        {
            SAPbouiCOM.Matrix matrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixID).Specific;
            SAPbouiCOM.DBDataSource db = oForm.DataSources.DBDataSources.Item(dbTable);
            if (matrix.RowCount == 0)
            {
                Global.GFunc.SetNewLine(matrix, db, 1, "");
            }
        }
    }
}
