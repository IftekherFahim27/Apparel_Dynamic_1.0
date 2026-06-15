using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Apparel_Dynamic_1._0.Helper;

namespace Apparel_Dynamic_1._0.Resources.Setup
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Setup.OrderType", "Resources/Setup/OrderType.b1f")]
    class OrderType : UserFormBase
    {
        public OrderType()
        {
        }

        private SAPbouiCOM.StaticText STPRDCOD, STPRDNAM;
        private SAPbouiCOM.EditText ETPRDCOD, ETPRDNAM, ETDOCTRY, ETCODE, ETNAME;

        private SAPbouiCOM.Matrix MTXORDR;
        private SAPbouiCOM.Button ADDButton, CancelButton, BTNEWLN;

        private bool _isAddButtonPressed = false;


        public override void OnInitializeComponent()
        {
            this.STPRDCOD = ((SAPbouiCOM.StaticText)(this.GetItem("STPRDCOD").Specific));
            this.STPRDNAM = ((SAPbouiCOM.StaticText)(this.GetItem("STPRDNAM").Specific));
            this.ETPRDCOD = ((SAPbouiCOM.EditText)(this.GetItem("ETPRDCOD").Specific));
            this.ETPRDCOD.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETPRDCOD_ChooseFromListAfter);
            this.ETPRDCOD.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETPRDCOD_ChooseFromListBefore);
            this.ETPRDNAM = ((SAPbouiCOM.EditText)(this.GetItem("ETPRDNAM").Specific));
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            this.ETCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETCODE").Specific));
            this.ETNAME = ((SAPbouiCOM.EditText)(this.GetItem("ETNAME").Specific));
            this.MTXORDR = ((SAPbouiCOM.Matrix)(this.GetItem("MTXORDR").Specific));
            this.MTXORDR.ClickBefore += new SAPbouiCOM._IMatrixEvents_ClickBeforeEventHandler(this.MTXORDR_ClickBefore);
            this.MTXORDR.LostFocusAfter += new SAPbouiCOM._IMatrixEvents_LostFocusAfterEventHandler(this.MTXORDR_LostFocusAfter);
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.ADDButton.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.ADDButton_PressedAfter);
            this.ADDButton.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.ADDButton_PressedBefore);
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.BTNEWLN = ((SAPbouiCOM.Button)(this.GetItem("BTNEWLN").Specific));
            this.BTNEWLN.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.BTNEWLN_PressedBefore);
            this.BTNEWLN.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.BTNEWLN_PressedAfter);
            this.OnCustomInitialize();

        }
        public override void OnInitializeFormEvents()
        {
            this.DataLoadAfter += new SAPbouiCOM.Framework.FormBase.DataLoadAfterHandler(this.Form_DataLoadAfter);
            this.RightClickBefore += new RightClickBeforeHandler(this.Form_RightClickBefore);

        }


        private void OnCustomInitialize()
        {
  
        }



        private void ETPRDCOD_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_PRD")
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
                    "Error filtering Product Group CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }


        }

        private void ETPRDCOD_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {

                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
                if (dt == null || dt.Rows.Count == 0)
                    return;

                string Code = dt.GetValue("Code", 0).ToString();
                SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETPRDCOD").Specific;
                ETCD.Value = Code;

                string Name = dt.GetValue("Name", 0).ToString();
                SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETPRDNAM").Specific;
                ETNM.Value = Name;

                //assign same value for Code and Name 
                ETCODE.Value = Code;
                ETNAME.Value = Name;

                // adding new line on the matrix and assign Code column a default value 
                EnsureLine(oForm, "MTXORDR", "@FIL_MR_ORDRTYPE");
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXORDR").Specific;

                // Only first line, only first time
                if (oMatrix.RowCount > 0)
                {
                    SAPbouiCOM.EditText txtCode = (SAPbouiCOM.EditText)oMatrix.Columns.Item("CLCODE").Cells.Item(1).Specific;
                    if (string.IsNullOrWhiteSpace(txtCode.Value))
                    {
                        txtCode.Value = "Code 1";
                    }
                }
                int minQtyColNo = GetMatrixColumnNumber(oMatrix, "CLMINQTY");

                if (minQtyColNo > 0)
                {
                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {
                        oMatrix.CommonSetting.SetCellEditable(i, minQtyColNo, i == 1);
                    }
                }

            }
            catch (Exception ex)
            {

            }


        }

        private void MTXORDR_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (_isAddButtonPressed)
                {
                    _isAddButtonPressed = false;
                    return;
                }

                if (pVal.Row <= 0)
                    return;

                if (pVal.ColUID != "CLMINQTY" && pVal.ColUID != "CLMAXQTY")
                    return;

                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXORDR").Specific;
                SAPbouiCOM.DBDataSource db = oForm.DataSources.DBDataSources.Item("@FIL_MR_ORDRTYPE");

                oForm.Freeze(true);

                if (pVal.ColUID == "CLMINQTY" && pVal.Row == 1)
                {
                    double minQty = GetMatrixDoubleValue(oMatrix, "CLMINQTY", 1);
                    if (minQty < 0)
                    {
                        SetMatrixValue(oMatrix, "CLMINQTY", 1, "0");
                        Application.SBO_Application.StatusBar.SetText(
                            "Minimum Quantity cannot be negative.",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error
                        );
                        return;
                    }
                }

                if (pVal.ColUID == "CLMAXQTY")
                {
                    double minQty = GetMatrixDoubleValue(oMatrix, "CLMINQTY", pVal.Row);
                    double maxQty = GetMatrixDoubleValue(oMatrix, "CLMAXQTY", pVal.Row);

                    if (maxQty < 0)
                    {
                        SetMatrixValue(oMatrix, "CLMAXQTY", pVal.Row, "0");

                        Application.SBO_Application.StatusBar.SetText(
                            "Maximum Quantity cannot be negative.",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error
                        );

                        return;
                    }

                    if (maxQty != 0 && maxQty <= minQty)
                    {
                        SetMatrixValue(oMatrix, "CLMAXQTY", pVal.Row, "");

                        Application.SBO_Application.StatusBar.SetText(
                            "Maximum Quantity must be greater than Minimum Quantity.",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error
                        );

                        return;
                    }

                    // If CLMAXQTY is non-zero, create next row
                    if (maxQty != 0 && oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        int nextRow = pVal.Row + 1;

                        // Add row only if current row is the last row
                        if (pVal.Row == oMatrix.RowCount)
                        {
                            oMatrix.FlushToDataSource();
                            Global.GFunc.SetNewLine(oMatrix, db, nextRow, "");
                            oMatrix.LoadFromDataSource();
                            SetMatrixValue(oMatrix, "#", nextRow, nextRow.ToString());
                        }

                        SetMatrixValue(oMatrix, "CLCODE", nextRow, "Code " + nextRow);
                        SetMatrixValue(oMatrix, "CLMINQTY", nextRow, (maxQty + 1).ToString("0"));
                    }
                }

                //SetMinQtyEditable(oMatrix);
                SetOrderTypeMatrixEditableAfterLoad(oMatrix);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
            finally
            {
                try
                {
                    SAPbouiCOM.Form oForm =
                        Application.SBO_Application.Forms.Item(pVal.FormUID);

                    oForm.Freeze(false);
                }
                catch { }
            }
        }


        private void ADDButton_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            _isAddButtonPressed = true;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

            // Do not validate in OK mode
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                return;

            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                ValidateForm(ref oForm, ref BubbleEvent);
            }

        }

        private void ADDButton_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                SAPbouiCOM.Matrix MTXORDR = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXORDR").Specific;
                MTXORDR.AutoResizeColumns();
            }

        }

        private void BTNEWLN_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Matrix matrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXORDR").Specific;

                if (matrix.VisualRowCount == 0)
                    return;

                int lastRow = matrix.VisualRowCount;

                string lastMaxText = GetMatrixStringValue(matrix, "CLMAXQTY", lastRow);
                double lastMinQty = GetMatrixDoubleValue(matrix, "CLMINQTY", lastRow);
                double lastMaxQty = GetMatrixDoubleValue(matrix, "CLMAXQTY", lastRow);

                if (string.IsNullOrWhiteSpace(lastMaxText) || lastMaxQty <= 0)
                {
                    Global.GFunc.ShowError("Please enter Maximum Quantity in the last row before adding new line.");
                    matrix.Columns.Item("CLMAXQTY").Cells.Item(lastRow).Click();
                    BubbleEvent = false;
                    return;
                }

                if (lastMaxQty <= lastMinQty)
                {
                    Global.GFunc.ShowError("Maximum Quantity must be greater than Minimum Quantity in the last row.");
                    matrix.Columns.Item("CLMAXQTY").Cells.Item(lastRow).Click();
                    BubbleEvent = false;
                    return;
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Add New Line Validation Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );

                BubbleEvent = false;
            }
        }

        private void BTNEWLN_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = null;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                oForm.Freeze(true);

                SAPbouiCOM.Matrix matrix =
                    (SAPbouiCOM.Matrix)oForm.Items.Item("MTXORDR").Specific;

                SAPbouiCOM.DBDataSource db =
                    oForm.DataSources.DBDataSources.Item("@FIL_MR_ORDRTYPE");

                matrix.FlushToDataSource();

                // remove deleted / empty ghost rows
                RemoveEmptyLastRowsFromDataSource(db);

                matrix.LoadFromDataSource();

                int dbRowCount = db.Size;

                if (dbRowCount == 0)
                {
                    Global.GFunc.SetNewLine(matrix, db, 1, "");

                    matrix.LoadFromDataSource();

                    SetMatrixValue(matrix, "#", 1, "1");
                    SetMatrixValue(matrix, "CLCODE", 1, "Code 1");
                    SetMatrixValue(matrix, "CLMINQTY", 1, "1");
                    SetMatrixValue(matrix, "CLMAXQTY", 1, "");

                    SetOrderTypeMatrixEditableAfterLoad(matrix);
                    return;
                }

                int lastRow = matrix.VisualRowCount;

                if (lastRow <= 0)
                    return;

                double lastMaxQty = GetMatrixDoubleValue(matrix, "CLMAXQTY", lastRow);

                int newLineNo = lastRow + 1;

                if (lastMaxQty > 0.0)
                {
                    Global.GFunc.SetNewLine(matrix, db, newLineNo, "");

                    matrix.LoadFromDataSource();

                    SetMatrixValue(matrix, "#", newLineNo, newLineNo.ToString());
                    SetMatrixValue(matrix, "CLCODE", newLineNo, "Code " + newLineNo);
                    SetMatrixValue(matrix, "CLMINQTY", newLineNo, (lastMaxQty + 1).ToString("0"));
                    SetMatrixValue(matrix, "CLMAXQTY", newLineNo, "");

                    SetOrderTypeMatrixEditableAfterLoad(matrix);

                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

                    matrix.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Add New Line Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
            finally
            {
                if (oForm != null)
                    oForm.Freeze(false);
            }
        }

        private void RemoveEmptyLastRowsFromDataSource(SAPbouiCOM.DBDataSource db)
        {
            for (int i = db.Size - 1; i >= 0; i--)
            {
                string minQtyText = db.GetValue("U_MINQTY", i).Trim();
                string maxQtyText = db.GetValue("U_MAXQTY", i).Trim();

                double minQty = 0;
                double maxQty = 0;

                double.TryParse(minQtyText, out minQty);
                double.TryParse(maxQtyText, out maxQty);

                // remove ghost row even if U_OTYPECODE has "Code 3"
                if (minQty <= 0 && maxQty <= 0)
                {
                    db.RemoveRecord(i);
                }
                else
                {
                    break;
                }
            }
        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            SAPbouiCOM.Form oForm = null;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                oForm.Freeze(true);

                // Enable Add New Line button only after existing data loaded
                oForm.Items.Item("BTNEWLN").Enabled = true;

                SAPbouiCOM.Matrix oMatrix =
                    (SAPbouiCOM.Matrix)oForm.Items.Item("MTXORDR").Specific;

                SetOrderTypeMatrixEditableAfterLoad(oMatrix);

                oMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "DataLoadAfter Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
            finally
            {
                if (oForm != null)
                {
                    try
                    {
                        oForm.Freeze(false);
                    }
                    catch { }
                }
            }
        }

        private void SetOrderTypeMatrixEditableAfterLoad(SAPbouiCOM.Matrix oMatrix)
        {
            int minQtyColNo = GetMatrixColumnNumber(oMatrix, "CLMINQTY");
            int maxQtyColNo = GetMatrixColumnNumber(oMatrix, "CLMAXQTY");

            if (minQtyColNo <= 0 || maxQtyColNo <= 0)
                return;

            int rowCount = oMatrix.VisualRowCount;

            if (rowCount <= 0)
                return;

            // Case 1: only one row
            if (rowCount == 1)
            {
                oMatrix.CommonSetting.SetCellEditable(1, minQtyColNo, true);
                oMatrix.CommonSetting.SetCellEditable(1, maxQtyColNo, true);
                return;
            }

            // Case 2: more than one row
            for (int i = 1; i <= rowCount; i++)
            {
                // CLMINQTY disabled for all rows
                oMatrix.CommonSetting.SetCellEditable(i, minQtyColNo, false);

                // CLMAXQTY enabled only for last row
                oMatrix.CommonSetting.SetCellEditable(i, maxQtyColNo, i == rowCount);
            }
        }


        private string GetMatrixStringValue(SAPbouiCOM.Matrix oMatrix, string colUID, int row)
        {
            SAPbouiCOM.EditText txt = (SAPbouiCOM.EditText)oMatrix.Columns.Item(colUID).Cells.Item(row).Specific;
            return txt.Value.Trim();
        }

        private bool ValidateForm(ref SAPbouiCOM.Form oForm, ref bool BubbleEvent)
        {
            string code = oForm.DataSources.DBDataSources.Item("@FIL_MH_ORDRTYPE").GetValue("U_PRDGRP", 0).Trim();

            if (string.IsNullOrWhiteSpace(code))
            {
                Global.GFunc.ShowError("Enter Product Group Master Code");
                oForm.ActiveItem = "ETPRDCOD";
                return BubbleEvent = false;
            }

            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                if (IsProductGroupAlreadyExists(code))
                {
                    Global.GFunc.ShowError("Product Group Master Code already exists.");
                    oForm.ActiveItem = "ETPRDCOD";
                    return BubbleEvent = false;
                }
            }

            SAPbouiCOM.Matrix oMatrix =
                (SAPbouiCOM.Matrix)oForm.Items.Item("MTXORDR").Specific;

            for (int i = 1; i <= oMatrix.VisualRowCount; i++)
            {
                string minQtyText = GetMatrixStringValue(oMatrix, "CLMINQTY", i);
                string maxQtyText = GetMatrixStringValue(oMatrix, "CLMAXQTY", i);

                // Skip only last auto-added empty row
                //if (i == oMatrix.RowCount && string.IsNullOrWhiteSpace(maxQtyText))
                //    continue;

                double minQty = GetMatrixDoubleValue(oMatrix, "CLMINQTY", i);
                double maxQty = GetMatrixDoubleValue(oMatrix, "CLMAXQTY", i);

                if (string.IsNullOrWhiteSpace(minQtyText) || minQty <= 0)
                {
                    Global.GFunc.ShowError("Minimum Quantity must have value in row " + i);
                    oMatrix.Columns.Item("CLMINQTY").Cells.Item(i).Click();
                    return BubbleEvent = false;
                }

                if (string.IsNullOrWhiteSpace(maxQtyText) || maxQty <= 0)
                {
                    Global.GFunc.ShowError("Maximum Quantity must have value in row " + i);
                    oMatrix.Columns.Item("CLMAXQTY").Cells.Item(i).Click();
                    return BubbleEvent = false;
                }

                if (maxQty <= minQty)
                {
                    Global.GFunc.ShowError("Maximum Quantity must be greater than Minimum Quantity in row " + i);
                    oMatrix.Columns.Item("CLMAXQTY").Cells.Item(i).Click();
                    return BubbleEvent = false;
                }
            }

            oMatrix.FlushToDataSource();

            return BubbleEvent;
        }

        private void MTXORDR_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.Row <= 0)
                    return;

                SAPbouiCOM.Form oForm =
                    Application.SBO_Application.Forms.Item(pVal.FormUID);

                SAPbouiCOM.Matrix matrix =
                    (SAPbouiCOM.Matrix)oForm.Items.Item("MTXORDR").Specific;

            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Matrix Click Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
        }

        private void Form_RightClickBefore(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item(eventInfo.FormUID);
            try
            {
                if (eventInfo.ItemUID != "MTXORDR" || eventInfo.Row <= 0)
                    return;

                oForm.EnableMenu("1293", true);
            }
            catch { }

        }


        private bool IsProductGroupAlreadyExists(string code)
        {
            SAPbobsCOM.Recordset oRs = null;

            try
            {
                oRs = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                code = code.Replace("'", "''");
                string query = $@"
                                SELECT COUNT(*) AS ""Cnt""
                                FROM ""@FIL_MH_ORDRTYPE""
                                WHERE ""U_PRDGRP"" = '{code}'";

                oRs.DoQuery(query);
                int count = Convert.ToInt32(oRs.Fields.Item("Cnt").Value);
                return count > 0;
            }
            finally
            {
                if (oRs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
                    oRs = null;
                }
            }
        }

        private double GetMatrixDoubleValue(SAPbouiCOM.Matrix oMatrix, string colUID, int row)
        {
            SAPbouiCOM.EditText txt =
                (SAPbouiCOM.EditText)oMatrix.Columns.Item(colUID).Cells.Item(row).Specific;

            double value = 0;
            double.TryParse(txt.Value, out value);

            return value;
        }

        private void SetMatrixValue(SAPbouiCOM.Matrix oMatrix, string colUID, int row, string value)
        {
            SAPbouiCOM.EditText txt =
                (SAPbouiCOM.EditText)oMatrix.Columns.Item(colUID).Cells.Item(row).Specific;

            txt.Value = value;
        }

        private void SetMinQtyEditable(SAPbouiCOM.Matrix oMatrix)
        {
            int minQtyColNo = GetMatrixColumnNumber(oMatrix, "CLMINQTY");

            if (minQtyColNo <= 0)
                return;

            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                oMatrix.CommonSetting.SetCellEditable(i, minQtyColNo, i == 1);
            }
        }

        private int GetMatrixColumnNumber(SAPbouiCOM.Matrix oMatrix, string colUID)
        {
            for (int i = 1; i <= oMatrix.Columns.Count; i++)
            {
                if (oMatrix.Columns.Item(i).UniqueID == colUID)
                    return i;
            }

            return -1;
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
