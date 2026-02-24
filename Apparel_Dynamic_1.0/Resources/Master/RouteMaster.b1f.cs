using Apparel_Dynamic_1._0.Helper;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;

namespace Apparel_Dynamic_1._0.Resources.Master
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Master.RouteMaster", "Resources/Master/RouteMaster.b1f")]
    class RouteMaster : UserFormBase
    {
        public RouteMaster()
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
            this.MTXSTAGE = ((SAPbouiCOM.Matrix)(this.GetItem("MTXSTAGE").Specific));
            this.MTXSTAGE.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.MTXSTAGE_ChooseFromListAfter);
            this.MTXSTAGE.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.MTXSTAGE_ChooseFromListBefore);
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.ADDButton.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.ADDButton_PressedAfter);
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
        private SAPbouiCOM.Matrix MTXSTAGE;
        private SAPbouiCOM.Button ADDButton, CancelButton;

       

        private void OnCustomInitialize()
        {

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

                        string query = $@"SELECT 1 FROM ""@FIL_MH_ORSM"" WHERE ""Code"" = '{UCode.Replace("'", "''")}'";
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
        private void ADDButton_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            if(oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXSTAGE").Specific;
                oMatrix.Columns.Item("CLSTGCOD").Editable = true;
                EnsureLine(oForm, "MTXSTAGE", "@FIL_MR_RSM1");             
            }
            else if(oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
            {
                oForm.Freeze(true);
                try
                {
                    ApplyStageCodeEditability(oForm);
                }
                finally
                {
                    oForm.Freeze(false);
                }
            }

        }
        private void MTXSTAGE_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {

                if (pVal.ColUID != "CLSTGCOD") return;
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Matrix oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXSTAGE").Specific;
                string colCode = "CLSTGCOD";

                HashSet<string> usedCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                for (int i = 1; i <= oMtx.RowCount; i++)
                {
                    var cell = (SAPbouiCOM.EditText)oMtx.Columns.Item(colCode).Cells.Item(i).Specific;
                    string code = (cell.Value ?? "").Trim();

                    if (!string.IsNullOrEmpty(code))
                        usedCodes.Add(code);
                }

              
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item("CFL_ORST");
                // Always clear previous conditions
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
                Application.SBO_Application.StatusBar.SetText("Route Stage CFL Error: " + ex.Message,
                   SAPbouiCOM.BoMessageTime.bmt_Short,
                   SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void MTXSTAGE_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXSTAGE").Specific;
                SAPbouiCOM.DBDataSource DBDataSourceLine = oForm.DataSources.DBDataSources.Item("@FIL_MR_RSM1");
                SAPbouiCOM.DataTable oDataTable = cflArg.SelectedObjects;

                if (oDataTable == null || oDataTable.Rows.Count == 0)
                    return;

                string code = oDataTable.GetValue("Code", 0).ToString();
                string name = oDataTable.GetValue("Desc", 0).ToString();

                int row = pVal.Row; // matrix row where CFL triggered
                oMatrix.SetCellWithoutValidation(row, "CLSTGCOD", code);
                oMatrix.SetCellWithoutValidation(row, "CLSTGNAM", name);
                oMatrix.FlushToDataSource();

                // Add new row if last row has data
                int lastRow = oMatrix.RowCount;
                bool lastRowHasData = !string.IsNullOrWhiteSpace(((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLSTGCOD").Cells.Item(lastRow).Specific).Value);
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
                Application.SBO_Application.StatusBar.SetText("Route Stage CFL Error: " + ex.Message,
                   SAPbouiCOM.BoMessageTime.bmt_Short,
                   SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void ApplyStageCodeEditability(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXSTAGE").Specific;
            oMatrix.Columns.Item("CLSTGCOD").Editable = true;
            string routeCode = ((SAPbouiCOM.EditText)oForm.Items.Item("ETCODE").Specific).Value.Trim();

            if (string.IsNullOrEmpty(routeCode))
                return;

            if (IsRouteCodeUsed(routeCode))
            {
                oMatrix.Columns.Item("CLSTGCOD").Editable = false;
            }
            else
            {
                oMatrix.Columns.Item("CLSTGCOD").Editable = true;
                AddLineIfLastRowHasValue(oForm, "MTXSTAGE", "@FIL_MR_RSM1", "U_STAGECODE");
            }
            oMatrix.LoadFromDataSource();
        }
        private bool IsRouteCodeUsed(string routeCode)
        {
            try
            {
                SAPbobsCOM.Recordset oRS =
                    (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = $@"
                                SELECT 1 FROM ""@FIL_DH_SMPLMAST""
                                WHERE ""U_ROUTSTAG"" = '{routeCode}'
                                UNION ALL
                                SELECT 1 FROM ""@FIL_DH_OPSM""
                                WHERE ""U_ROUTESTAGE"" = '{routeCode}'
                                LIMIT 1";

                oRS.DoQuery(query);

                return oRS.RecordCount > 0;
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Route Check Error: " + ex.Message,
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
            string code = oForm.DataSources.DBDataSources.Item("@FIL_MH_ORSM").GetValue("Code", 0);
            string name = oForm.DataSources.DBDataSources.Item("@FIL_MH_ORSM").GetValue("Name", 0);
            if (code == "")
            {
                Global.GFunc.ShowError("Enter Route Master Code");
                oForm.ActiveItem = "ETCODE";
                return BubbleEvent = false;
            }
            else if (name == "")
            {
                Global.GFunc.ShowError("Enter Route Master Description");
                oForm.ActiveItem = "ETNAME";
                return BubbleEvent = false;
            }
            
            PreventEmptyLastRow(oForm, "@FIL_MR_RSM1", MTXSTAGE, "U_STAGECODE");

            if (!RequireAtLeastOneFilledRow(oForm, "@FIL_MR_RSM1", MTXSTAGE, "U_STAGECODE", ref BubbleEvent))
            {
                EnsureLine(oForm, "MTXSTAGE", "@FIL_MR_RSM1");
                return BubbleEvent;
            }

            return BubbleEvent;
        }

        private bool RequireAtLeastOneFilledRow(SAPbouiCOM.Form oForm, string dbDatasourceUID, SAPbouiCOM.Matrix matrix, string columnName, ref bool BubbleEvent)
        {
            try
            {
                SAPbouiCOM.DBDataSource oDB = oForm.DataSources.DBDataSources.Item(dbDatasourceUID);
                int rowCount = matrix.VisualRowCount;
                if (rowCount == 0)
                {
                    Global.GFunc.ShowError("At least one Route Stage is required.");
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

        public static void EnsureLine(SAPbouiCOM.Form oForm, string matrixID, string dbTable)
        {
            SAPbouiCOM.Matrix matrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixID).Specific;
            SAPbouiCOM.DBDataSource db = oForm.DataSources.DBDataSources.Item(dbTable);
            if (matrix.RowCount == 0)
            {
                Global.GFunc.SetNewLine(matrix, db, 1, "");
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

    }
}
