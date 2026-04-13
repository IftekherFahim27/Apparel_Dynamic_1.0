using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Apparel_Dynamic_1._0.Helper;

namespace Apparel_Dynamic_1._0.Resources.Transaction
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Transaction.OTT", "Resources/Transaction/OTT.b1f")]
    class OTT : UserFormBase
    {
        public OTT()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.STOTDATE = ((SAPbouiCOM.StaticText)(this.GetItem("STOTDATE").Specific));
            this.STSTATUS = ((SAPbouiCOM.StaticText)(this.GetItem("STSTATUS").Specific));
            this.STDOCNUM = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCNUM").Specific));
            this.STMERCHN = ((SAPbouiCOM.StaticText)(this.GetItem("STMERCHN").Specific));
            this.FOLOTDTL = ((SAPbouiCOM.Folder)(this.GetItem("FOLOTDTL").Specific));
            this.FOLDODTL = ((SAPbouiCOM.Folder)(this.GetItem("FOLDODTL").Specific));
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.ADDButton.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.ADDButton_PressedAfter);
            this.ADDButton.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.ADDButton_PressedBefore);
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.ETOTDATE = ((SAPbouiCOM.EditText)(this.GetItem("ETOTDATE").Specific));
            this.CBSTATUS = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSTATUS").Specific));
            this.ETDOCNUM = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCNUM").Specific));
            this.CBSERIES = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSERIES").Specific));
            this.ETMERCNM = ((SAPbouiCOM.EditText)(this.GetItem("ETMERCNM").Specific));
            this.ETMERCHN = ((SAPbouiCOM.EditText)(this.GetItem("ETMERCHN").Specific));
            this.ETMERCHN.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETMERCHN_ChooseFromListAfter);
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            this.MTXOTDTL = ((SAPbouiCOM.Matrix)(this.GetItem("MTXOTDTL").Specific));
            this.MTXOTDTL.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.MTXOTDTL_ValidateAfter);
            this.MTXOTDTL.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.MTXOTDTL_ChooseFromListAfter);
            this.GRDDODTL = ((SAPbouiCOM.Grid)(this.GetItem("GRDDODTL").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.DataLoadAfter += new DataLoadAfterHandler(this.Form_DataLoadAfter);

        }

        private void OnCustomInitialize()
        {

        }

        private SAPbouiCOM.StaticText STOTDATE, STSTATUS, STDOCNUM, STMERCHN;
        private SAPbouiCOM.Folder FOLOTDTL, FOLDODTL;
        private SAPbouiCOM.Button ADDButton, CancelButton;
        private SAPbouiCOM.EditText ETOTDATE, ETDOCNUM, ETMERCNM, ETMERCHN, ETDOCTRY;
        private SAPbouiCOM.ComboBox CBSTATUS, CBSERIES;



        private SAPbouiCOM.Matrix MTXOTDTL;
        private SAPbouiCOM.Grid GRDDODTL;



        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            //throw new System.NotImplementedException();

        }

        private void ETMERCHN_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("empID", 0).ToString().Trim();
            string Name = dt.GetValue("U_FNAME", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETMERCHN").Specific;
            ETCD.Value = Code;
            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETMERCNM").Specific;
            ETNM.Value = Name;

        }

        private void MTXOTDTL_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXOTDTL").Specific;
                SAPbouiCOM.DBDataSource DBDataSourceLine = oForm.DataSources.DBDataSources.Item("@FIL_DR_TT1");
                SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;

                if (dt == null || dt.Rows.Count == 0)
                    return;

                string styleCode = dt.GetValue("U_STYLECODE", 0).ToString();
                string styleName = dt.GetValue("U_STYLENM", 0).ToString();
                string styleEntry = dt.GetValue("DocEntry", 0).ToString();
                int row = pVal.Row;
                //Set Values
                oMatrix.SetCellWithoutValidation(row, "CLSTYLNO", styleCode);
                oMatrix.SetCellWithoutValidation(row, "CLSTYLNM", styleName);
                oMatrix.SetCellWithoutValidation(row, "CLSLNTRY", styleEntry);
                oMatrix.FlushToDataSource();

                AddLineIfLastRowHasValue(oForm, "MTXOTDTL", "@FIL_DR_TT1", "U_STYLECODE");

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }

            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Color Matrix CFL Error: " + ex.Message,
                   SAPbouiCOM.BoMessageTime.bmt_Short,
                   SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }
        private void MTXOTDTL_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID != "CLSTYLNO") return;

                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Matrix m = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXOTDTL").Specific;

                int row = pVal.Row;
                if (row <= 0) return;

                oForm.Freeze(true);
                try
                {
                    string code = ((SAPbouiCOM.EditText)m.Columns.Item("CLSTYLNO").Cells.Item(row).Specific).Value
                                    .Replace("\0", "").Trim();

                    if (string.IsNullOrEmpty(code))
                    {
                        // 1) clear data
                        ((SAPbouiCOM.EditText)m.Columns.Item("CLSTYLNM").Cells.Item(row).Specific).Value = "";
                        ((SAPbouiCOM.EditText)m.Columns.Item("CLSLNTRY").Cells.Item(row).Specific).Value = "";
                        SAPbouiCOM.DBDataSource dbds = oForm.DataSources.DBDataSources.Item("@FIL_DR_TT1");
                        dbds.SetValue("U_QUANTITY", row - 1, "0.000000");
                        //m.LoadFromDataSource();


                        // 2) remove this row + resequence LineId
                        RemoveRowIfCodeEmptyAndResequence(oForm, m, "@FIL_DR_TT1", "U_STYLECODE");
                        EnsureLine(oForm, "MTXOTDTL", "@FIL_DR_TT1");
                        AddLineIfLastRowHasValue(oForm, "MTXOTDTL", "@FIL_DR_TT1", "U_STYLECODE");
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
        private void ADDButton_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                PreventEmptyLastRow(oForm, "@FIL_DR_TT1", MTXOTDTL, "U_STYLECODE");
            }

        }

        private void ADDButton_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.DBDataSource oDBH = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@FIL_DH_OTT");   //DEFINE  DATASOURCES.
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                SAPbouiCOM.ComboBox ocmb = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBSERIES").Specific;
                Global.GFunc.LoadComboBoxSeries(ocmb, "FIL_D_OTT");  //Object Type
                string ocmbvalue = ocmb.Selected.Value;
                long docno = oForm.BusinessObject.GetNextSerialNumber(ocmbvalue, "FIL_D_OTT");

                oDBH.SetValue("DocNum", 0, docno.ToString()); // only set the value in string.

                //Date
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETOTDATE").Specific).Value = DateTime.Now.ToString("yyyyMMdd");

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

        private void RemoveRowIfCodeEmptyAndResequence(SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix matrix, string dbDatasourceUID, string codeFieldName)
        {
            matrix.FlushToDataSource();
            SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item(dbDatasourceUID);
            for (int i = ds.Size - 1; i >= 0; i--)
            {
                string code = (ds.GetValue(codeFieldName, i) ?? "").Replace("\0", "").Trim();

                if (string.IsNullOrEmpty(code))
                {
                    ds.RemoveRecord(i);
                }
            }

            for (int i = 0; i < ds.Size; i++)
            {
                ds.SetValue("LineId", i, (i + 1).ToString());
            }
            matrix.LoadFromDataSource();
        }

    }
}
