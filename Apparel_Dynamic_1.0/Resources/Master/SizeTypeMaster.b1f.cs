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
            this.MTXSIZE.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.MTXSIZE_ValidateAfter);
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
            this.DataLoadAfter += new DataLoadAfterHandler(this.Form_DataLoadAfter);

        }

        private SAPbouiCOM.StaticText STCODE, STNAME;
        private SAPbouiCOM.EditText ETCODE, ETNAME, ETDOCTRY;
        private SAPbouiCOM.CheckBox CKACTIVE;
        private SAPbouiCOM.Matrix MTXSIZE;
        private SAPbouiCOM.Button ADDButton, CancelButton;



        private void OnCustomInitialize()
        {

        }

        private void MTXSIZE_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID != "CLSZCODE") return;

                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Matrix oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXSIZE").Specific;

                int row = pVal.Row;
                if (row <= 0) return;

                oForm.Freeze(true);
                try
                {
                    string code = ((SAPbouiCOM.EditText)oMtx.Columns.Item("CLSZCODE").Cells.Item(row).Specific).Value
                                    .Replace("\0", "").Trim();

                    if (string.IsNullOrEmpty(code))
                    {
                        // 1) clear name (no flicker because Freeze)
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("CLSZNAME").Cells.Item(row).Specific).Value = "";

                        // 2) remove this row + resequence LineId
                        RemoveRowIfCodeEmptyAndResequence(oForm, oMtx, "@FIL_MR_STM1", "U_SIZECODE");
                        AddLineIfLastRowHasValue(oForm, "MTXSIZE", "@FIL_MR_STM1", "U_SIZECODE");
                    }

                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                finally
                {
                    oForm.Freeze(false);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Validation Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }
        private void RemoveRowIfCodeEmptyAndResequence(SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix matrix, string dbDatasourceUID, string codeFieldName)
        {
    
            matrix.FlushToDataSource();

            SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item(dbDatasourceUID);

            // Remove rows where code is empty (iterate from bottom to avoid index shifting)
            for (int i = ds.Size - 1; i >= 0; i--)
            {
                string code = (ds.GetValue(codeFieldName, i) ?? "").Replace("\0", "").Trim();

                if (string.IsNullOrEmpty(code))
                {
                    ds.RemoveRecord(i);
                }
            }

            // Resequence LineId to 1..n 
            for (int i = 0; i < ds.Size; i++)
            {
                ds.SetValue("LineId", i, (i + 1).ToString());
            }

            matrix.LoadFromDataSource();
        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

            oForm.Freeze(true);
            try
            {
                SetItemsEnabled(oForm, false, "ETCODE");
                ApplyStageCodeEditability(oForm);
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
        private void ApplyStageCodeEditability(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXSIZE").Specific;
            oMatrix.Columns.Item("CLSZCODE").Editable = true;
            string sizeTypeCode = ((SAPbouiCOM.EditText)oForm.Items.Item("ETCODE").Specific).Value.Trim();

            if (string.IsNullOrEmpty(sizeTypeCode))
                return;

            if (IsSizeTypeCodeUsed(sizeTypeCode))
            {
                oMatrix.Columns.Item("CLSZCODE").Editable = false;
            }
            else
            {
                oMatrix.Columns.Item("CLSZCODE").Editable = true;
                AddLineIfLastRowHasValue(oForm, "MTXSIZE", "@FIL_MR_STM1", "U_SIZECODE");
            }
            oMatrix.LoadFromDataSource();
        }
        private bool IsSizeTypeCodeUsed(string sizeTypeCode)
        {
            try
            {
                SAPbobsCOM.Recordset oRS =
                    (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = $@"
                                SELECT 1 FROM ""@FIL_DR_PSMST""
                                WHERE ""U_SIZETYPE"" = '{sizeTypeCode}'
                                LIMIT 1";

                oRS.DoQuery(query);

                return oRS.RecordCount > 0;
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Size Type Used Check Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                return false;
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

            if (!RequireAtLeastOneFilledRow(oForm,MTXSIZE, "CLSZCODE", ref BubbleEvent))
            {
                EnsureLine(oForm, "MTXSIZE", "@FIL_MR_STM1");
                return BubbleEvent;
            }

            return BubbleEvent;
        }


        private bool RequireAtLeastOneFilledRow(SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix matrix, string matrixColumnId, ref bool BubbleEvent)
        {
            try
            {
                int rowCount = matrix.VisualRowCount;
                if (rowCount == 0)
                {
                    Global.GFunc.ShowError("At least one Route Stage is required.");
                    oForm.ActiveItem = matrix.Item.UniqueID;
                    BubbleEvent = false;
                    return false;
                }

                for (int i = 1; i <= rowCount; i++) // Matrix row index is 1-based
                {
                    string val = ((SAPbouiCOM.EditText)matrix.Columns.Item(matrixColumnId)
                                    .Cells.Item(i).Specific).Value
                                    .Replace("\0", "")
                                    .Trim();

                    if (!string.IsNullOrEmpty(val) && val != "0.0")
                        return true;
                }

                Global.GFunc.ShowError("At least one Route Stage is required.");
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
