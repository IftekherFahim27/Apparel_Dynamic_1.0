using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Apparel_Dynamic_1._0.Helper;


namespace Apparel_Dynamic_1._0.Resources.Setup
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Setup.LeadTime", "Resources/Setup/LeadTime.b1f")]
    class LeadTime : UserFormBase
    {
        public LeadTime()
        {
        }

        private SAPbouiCOM.StaticText STEFRMDT, STEFTODT, STDOCNUM;
        private SAPbouiCOM.EditText ETEFRMDT, ETEFTODT, ETDOCTRY, ETDOCNUM;
        private SAPbouiCOM.ComboBox CBSERIES;
        private SAPbouiCOM.Matrix MTXLEDTM;
        private SAPbouiCOM.Button ADDButton,CancelButton;

        private bool _isAddButtonPressed = false;

        public override void OnInitializeComponent()
        {
            this.STEFRMDT = ((SAPbouiCOM.StaticText)(this.GetItem("STEFRMDT").Specific));
            this.STEFTODT = ((SAPbouiCOM.StaticText)(this.GetItem("STEFTODT").Specific));
            this.ETEFRMDT = ((SAPbouiCOM.EditText)(this.GetItem("ETEFRMDT").Specific));
            this.ETEFTODT = ((SAPbouiCOM.EditText)(this.GetItem("ETEFTODT").Specific));
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            this.ETDOCNUM = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCNUM").Specific));
            this.CBSERIES = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSERIES").Specific));
            this.MTXLEDTM = ((SAPbouiCOM.Matrix)(this.GetItem("MTXLEDTM").Specific));
            this.MTXLEDTM.LostFocusAfter += new SAPbouiCOM._IMatrixEvents_LostFocusAfterEventHandler(this.MTXLEDTM_LostFocusAfter);
            this.MTXLEDTM.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.MTXLEDTM_ChooseFromListAfter);
            this.MTXLEDTM.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.MTXLEDTM_ChooseFromListBefore);
            this.STDOCNUM = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCNUM").Specific));
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.ADDButton.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.ADDButton_PressedAfter);
            this.ADDButton.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.ADDButton_PressedBefore);
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.OnCustomInitialize();

        }


        public override void OnInitializeFormEvents()
        {
            this.RightClickBefore += new SAPbouiCOM.Framework.FormBase.RightClickBeforeHandler(this.Form_RightClickBefore);
            this.DataLoadAfter += new DataLoadAfterHandler(this.Form_DataLoadAfter);

        }



        private void OnCustomInitialize()
        {

        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXLEDTM").Specific;
            SetItemsEnabled(oForm, false, "ETDOCNUM");
            LoadMatrixCombos(oForm);
            EnsureLine(oForm, "MTXLEDTM", "@FIL_DR_LEADTMST");
            AddLineIfLastRowHasValue(oForm, "MTXLEDTM", "@FIL_DR_LEADTMST", "U_CARDCODE");
            oMatrix.AutoResizeColumns();
        }

        private void ADDButton_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            _isAddButtonPressed = true;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
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
                //Series Initialization
                SAPbouiCOM.DBDataSource oDBH = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@FIL_DH_LEADTMST");
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    SAPbouiCOM.ComboBox ocmb = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBSERIES").Specific;
                    Global.GFunc.LoadComboBoxSeries(ocmb, "FIL_D_LEADTMST");  //Object Type
                    string ocmbvalue = ocmb.Selected.Value;
                    long docno = oForm.BusinessObject.GetNextSerialNumber(ocmbvalue, "FIL_D_LEADTMST");
                    oDBH.SetValue("DocNum", 0, docno.ToString()); // only set the value in string.

                    //LoadShippingModeComboInMatrix(oForm);
                    LoadMatrixCombos(oForm);
                    EnsureLine(oForm, "MTXLEDTM", "@FIL_DR_LEADTMST");
                }
            }

        }


        private bool ValidateForm(ref SAPbouiCOM.Form oForm, ref bool BubbleEvent)
        {
            SAPbouiCOM.DBDataSource db =oForm.DataSources.DBDataSources.Item("@FIL_DH_LEADTMST");

            string efFrmDate = db.GetValue("U_EFROMDATE", 0).Trim();
            string efToDate = db.GetValue("U_ETODATE", 0).Trim();

            if (string.IsNullOrWhiteSpace(efFrmDate))
            {
                Global.GFunc.ShowError("Enter Effective From Date");
                oForm.ActiveItem = "ETEFRMDT";
                return BubbleEvent = false;
            }

            if (string.IsNullOrWhiteSpace(efToDate))
            {
                Global.GFunc.ShowError("Enter Effective To Date");
                oForm.ActiveItem = "ETEFTODT";
                return BubbleEvent = false;
            }

            DateTime fromDate = DateTime.ParseExact(efFrmDate, "yyyyMMdd", null);
            DateTime toDate = DateTime.ParseExact(efToDate, "yyyyMMdd", null);

            if (toDate < fromDate)
            {
                Global.GFunc.ShowError("Effective To Date cannot be before Effective From Date");
                oForm.ActiveItem = "ETEFTODT";
                return BubbleEvent = false;
            }

            string docEntry = db.GetValue("DocEntry", 0).Trim();

            if (IsLeadTimeDateRangeExists(fromDate, toDate, docEntry))
            {
                Global.GFunc.ShowError("This effective date period already exists or overlaps with another period.");
                oForm.ActiveItem = "ETEFRMDT";
                return BubbleEvent = false;
            }

            PreventEmptyLastRow(oForm, "@FIL_DR_LEADTMST", MTXLEDTM, "U_CARDCODE");

            return BubbleEvent;
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

        private void MTXLEDTM_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (_isAddButtonPressed)
            {
                _isAddButtonPressed = false;
                return;
            }

            SAPbouiCOM.Form oForm = null;

            try
            {
                if (pVal.Row <= 0)
                    return;

                if (pVal.ColUID != "CLLEADTM")
                    return;

                oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Matrix matrix =(SAPbouiCOM.Matrix)oForm.Items.Item("MTXLEDTM").Specific;
                SAPbouiCOM.DBDataSource db =oForm.DataSources.DBDataSources.Item("@FIL_DR_LEADTMST");
                matrix.FlushToDataSource();

                int currentRow = pVal.Row;
                int currentIndex = currentRow - 1;

                string vendorCode = db.GetValue("U_CARDCODE", currentIndex).Trim();
                string shipFromCountry = db.GetValue("U_CONTRYCODE", currentIndex).Trim();
                string shippingMode = db.GetValue("U_SHIPMODE", currentIndex).Trim();
                string incoTerm = db.GetValue("U_INCOTRMS", currentIndex).Trim();
                string itemGroup = db.GetValue("U_ITMGRPCD", currentIndex).Trim();

                if (string.IsNullOrWhiteSpace(vendorCode) ||
                    string.IsNullOrWhiteSpace(shipFromCountry) ||
                    string.IsNullOrWhiteSpace(shippingMode) ||
                    string.IsNullOrWhiteSpace(incoTerm) ||
                    string.IsNullOrWhiteSpace(itemGroup))
                {
                    return;
                }

                for (int i = 0; i < db.Size; i++)
                {
                    if (i == currentIndex)
                        continue;

                    string oldVendorCode = db.GetValue("U_CARDCODE", i).Trim();
                    string oldShipFromCountry = db.GetValue("U_CONTRYCODE", i).Trim();
                    string oldShippingMode = db.GetValue("U_SHIPMODE", i).Trim();
                    string oldIncoTerm = db.GetValue("U_INCOTRMS", i).Trim();
                    string oldItemGroup = db.GetValue("U_ITMGRPCD", i).Trim();

                    if (vendorCode == oldVendorCode &&
                        shipFromCountry == oldShipFromCountry &&
                        shippingMode == oldShippingMode &&
                        incoTerm == oldIncoTerm &&
                        itemGroup == oldItemGroup)
                    {
                        Global.GFunc.ShowError(
                            "Duplicate combination found with row no. " + (i + 1)
                        );

                        ClearLeadTimeRow(db, currentIndex);
                        matrix.LoadFromDataSource();

                        return;
                    }
                }

                //AddLineIfLastRowHasValue(oForm,"MTXLEDTM","@FIL_DR_LEADTMST", "U_CARDCODE");

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Lead Time duplicate validation error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
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

        private void ClearLeadTimeRow(SAPbouiCOM.DBDataSource db, int rowIndex)
        {
            db.SetValue("U_CARDCODE", rowIndex, "");
            db.SetValue("U_CARDNAME", rowIndex, "");
            db.SetValue("U_CONTRYCODE", rowIndex, "");
            db.SetValue("U_SHIPMODE", rowIndex, "");
            db.SetValue("U_INCOTRMS", rowIndex, "");
            db.SetValue("U_ITMGRPCD", rowIndex, "");
            db.SetValue("U_LEADDAYS", rowIndex, "");
        }

        private bool IsLeadTimeDateRangeExists(DateTime fromDate, DateTime toDate, string currentDocEntry)
        {
            SAPbobsCOM.Recordset rs = null;

            try
            {
                SAPbobsCOM.Company oCompany =(SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string fromDateStr = fromDate.ToString("yyyyMMdd");
                string toDateStr = toDate.ToString("yyyyMMdd");

                string query = $@"
                                SELECT TOP 1 ""DocEntry""
                                FROM ""@FIL_DH_LEADTMST""
                                WHERE 
                                    ""U_EFROMDATE"" <= '{toDateStr}'
                                    AND ""U_ETODATE"" >= '{fromDateStr}'";

                if (!string.IsNullOrWhiteSpace(currentDocEntry))
                {
                    query += $@" AND ""DocEntry"" <> '{currentDocEntry}'";
                }

                rs.DoQuery(query);

                return !rs.EoF;
            }
            catch (Exception ex)
            {
                Global.GFunc.ShowError("Date validation error: " + ex.Message);
                return true;
            }
            finally
            {
                if (rs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                    rs = null;
                }

                GC.Collect();
            }
        }



        private void Form_RightClickBefore(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item(eventInfo.FormUID);
            try
            {
                if (eventInfo.ItemUID != "MTXLEDTM" || eventInfo.Row <= 0)
                    return;

                oForm.EnableMenu("1293", true);
            }
            catch { }

        }

        private void MTXLEDTM_ChooseFromListBefore(object sboObject,SAPbouiCOM.SBOItemEventArg pVal,out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.ColUID != "CLVNCOD")
                    return;

                SAPbouiCOM.Form oForm =Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.ChooseFromList oCFL =oForm.ChooseFromLists.Item("CFL_OCRD");

                SAPbouiCOM.Conditions oCons =new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition oCon =oCons.Add();

                oCon.Alias = "CardType";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "S"; 

                oCFL.SetConditions(oCons);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Vendor CFL condition error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );

                BubbleEvent = false;
            }
        }

        private void MTXLEDTM_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg =(SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                SAPbouiCOM.Form oForm =Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Matrix oMatrix =(SAPbouiCOM.Matrix)oForm.Items.Item("MTXLEDTM").Specific;
                SAPbouiCOM.DBDataSource db =oForm.DataSources.DBDataSources.Item("@FIL_DR_LEADTMST");
                SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;

                if (dt == null || dt.Rows.Count == 0)
                    return;

                int row = pVal.Row;

                if (pVal.ColUID == "CLVNCOD") // Currency OCRN
                {
                    string cardCode = dt.GetValue("CardCode", 0).ToString();
                    string cardName = dt.GetValue("CardName", 0).ToString();

                    oMatrix.SetCellWithoutValidation(row, "CLVNCOD", cardCode);
                    oMatrix.SetCellWithoutValidation(row, "CLVNNAM", cardName);
                    oMatrix.AutoResizeColumns();
                   
                }
                else if (pVal.ColUID == "CLSHFCN") // Country OCRY
                {
                    string countryName = dt.GetValue("Name", 0).ToString();
                    oMatrix.SetCellWithoutValidation(row, "CLSHFCN", countryName);
                }
                else if (pVal.ColUID == "CLITMGRP") // Item Group OITB
                {
                    string itemGroupName = dt.GetValue("ItmsGrpNam", 0).ToString();
                    oMatrix.SetCellWithoutValidation(row, "CLITMGRP", itemGroupName);
                    AddLineIfLastRowHasValue(oForm, "MTXLEDTM", "@FIL_DR_LEADTMST", "U_CARDCODE");
                }
                else
                {
                    return;
                }

                oMatrix.FlushToDataSource();
                //AddLineIfLastRowHasValue(oForm, "MTXLEDTM", "@FIL_DR_LEADTMST", "U_CARDCODE");
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "MTXLEDTM CFL Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
        }


       

        private void LoadMatrixCombos(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix =
                    (SAPbouiCOM.Matrix)oForm.Items.Item("MTXLEDTM").Specific;

                // Shipping Mode
                Global.GFunc.SetMatrixComboValue(
                    oMatrix,
                    "CLSHPMOD",
                    @"SELECT
                ""TrnspCode"",
                ""TrnspName""
              FROM ""OSHP""
              ORDER BY ""TrnspName"""
                );

                // Incoterms
                Global.GFunc.SetMatrixComboValue(
                    oMatrix,
                    "CLINTRMS",
                    @"SELECT
                ""Code"",
                ""Name""
              FROM ""@FIL_MH_INCOTRMS""
              WHERE ""U_ACTIVE"" = 'Y'
              ORDER BY ""Name"""
                );
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Matrix Combo Load Error : " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                try
                {
                    oForm.Freeze(false);
                }
                catch { }
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
    }
}
