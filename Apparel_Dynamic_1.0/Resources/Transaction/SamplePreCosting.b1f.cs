using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Apparel_Dynamic_1._0.Helper;
using Apparel_Dynamic_1._0.Resources.Master;

namespace Apparel_Dynamic_1._0.Resources.Transaction
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Transaction.SamplePreCosting", "Resources/Transaction/SamplePreCosting.b1f")]
    class SamplePreCosting : UserFormBase
    {
        public SamplePreCosting()
        {
        }

        private SAPbouiCOM.StaticText STSMPLCD, STCURR, STTCNAMT, STNO, STDOCNUM, STDATE, STVERSON, STBUYER;

        private SAPbouiCOM.ComboBox CBNO;

        private SAPbouiCOM.EditText ETCURR, ETTCNAMT, ETNO, ETBYRNM, ETSMPLNM, ETDOCNUM, ETDATE, ETVERSON, ETBUYER, ETDOCTRY, ETSMPLCD;

        private SAPbouiCOM.Folder FOLCMPNT, FOLOTCST, FOLVERSN;


        private SAPbouiCOM.Button ADDButton, CancelButton, BTNVRNUP, BTNLCSTH;

        private SAPbouiCOM.Matrix MTXCMPNT, MTXOTCST;

        private SAPbouiCOM.LinkedButton LKBUYER, LKSMPLCD;



        private SAPbouiCOM.Grid GRDVERSN;
        private string SelDocEntry = "";
        private string SelDocNum = "";
        private string SelVersion = "";
        private string SelLogInst = "";


        public override void OnInitializeComponent()
        {
            this.STSMPLCD = ((SAPbouiCOM.StaticText)(this.GetItem("STSMPLCD").Specific));
            this.STCURR = ((SAPbouiCOM.StaticText)(this.GetItem("STCURR").Specific));
            this.STTCNAMT = ((SAPbouiCOM.StaticText)(this.GetItem("STTCNAMT").Specific));
            this.STNO = ((SAPbouiCOM.StaticText)(this.GetItem("STNO").Specific));
            this.STDOCNUM = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCNUM").Specific));
            this.STDATE = ((SAPbouiCOM.StaticText)(this.GetItem("STDATE").Specific));
            this.STVERSON = ((SAPbouiCOM.StaticText)(this.GetItem("STVERSON").Specific));
            this.STBUYER = ((SAPbouiCOM.StaticText)(this.GetItem("STBUYER").Specific));
            this.CBNO = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSERIES").Specific));
            this.ETCURR = ((SAPbouiCOM.EditText)(this.GetItem("ETCURR").Specific));
            this.ETCURR.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETCURR_ChooseFromListAfter);
            this.ETTCNAMT = ((SAPbouiCOM.EditText)(this.GetItem("ETTCNAMT").Specific));
            this.ETDOCNUM = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCNUM").Specific));
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            this.ETDATE = ((SAPbouiCOM.EditText)(this.GetItem("ETDATE").Specific));
            this.ETVERSON = ((SAPbouiCOM.EditText)(this.GetItem("ETVERSON").Specific));
            this.ETBUYER = ((SAPbouiCOM.EditText)(this.GetItem("ETBUYER").Specific));
            this.ETBUYER.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETBUYER_ChooseFromListAfter);
            this.FOLCMPNT = ((SAPbouiCOM.Folder)(this.GetItem("FOLCMPNT").Specific));
            this.FOLOTCST = ((SAPbouiCOM.Folder)(this.GetItem("FOLOTCST").Specific));
            this.FOLOTCST.ClickAfter += new SAPbouiCOM._IFolderEvents_ClickAfterEventHandler(this.FOLOTCST_ClickAfter);
            this.FOLVERSN = ((SAPbouiCOM.Folder)(this.GetItem("FOLVERSN").Specific));
            this.FOLVERSN.ClickAfter += new SAPbouiCOM._IFolderEvents_ClickAfterEventHandler(this.FOLVERSN_ClickAfter);
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.ADDButton.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.ADDButton_PressedAfter);
            this.ADDButton.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.ADDButton_PressedBefore);
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.BTNVRNUP = ((SAPbouiCOM.Button)(this.GetItem("BTNVRNUP").Specific));
            this.BTNVRNUP.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.BTNVRNUP_PressedAfter);
            this.MTXCMPNT = ((SAPbouiCOM.Matrix)(this.GetItem("MTXCMPNT").Specific));
            this.MTXCMPNT.LostFocusAfter += new SAPbouiCOM._IMatrixEvents_LostFocusAfterEventHandler(this.MTXCMPNT_LostFocusAfter);
            this.MTXCMPNT.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.MTXCMPNT_ChooseFromListAfter);
            this.MTXCMPNT.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.MTXCMPNT_ChooseFromListBefore);
            this.MTXCMPNT.ComboSelectAfter += new SAPbouiCOM._IMatrixEvents_ComboSelectAfterEventHandler(this.MTXCMPNT_ComboSelectAfter);
            this.MTXOTCST = ((SAPbouiCOM.Matrix)(this.GetItem("MTXOTCST").Specific));
            this.MTXOTCST.LostFocusAfter += new SAPbouiCOM._IMatrixEvents_LostFocusAfterEventHandler(this.MTXOTCST_LostFocusAfter);
            this.LKBUYER = ((SAPbouiCOM.LinkedButton)(this.GetItem("LKBUYER").Specific));
            this.LKSMPLCD = ((SAPbouiCOM.LinkedButton)(this.GetItem("LKSMPLCD").Specific));
            this.LKSMPLCD.PressedAfter += new SAPbouiCOM._ILinkedButtonEvents_PressedAfterEventHandler(this.LKSMPLCD_PressedAfter);
            this.ETSMPLNM = ((SAPbouiCOM.EditText)(this.GetItem("ETSMPLNM").Specific));
            this.ETBYRNM = ((SAPbouiCOM.EditText)(this.GetItem("ETBYRNM").Specific));
            this.GRDVERSN = ((SAPbouiCOM.Grid)(this.GetItem("GRDVERSN").Specific));
            this.GRDVERSN.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.GRDVERSN_DoubleClickAfter);
            this.BTNLCSTH = ((SAPbouiCOM.Button)(this.GetItem("BTNLCSTH").Specific));
            this.BTNLCSTH.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.BTNLCSTH_PressedAfter);
            this.ETSMPLCD = ((SAPbouiCOM.EditText)(this.GetItem("ETSMPLCD").Specific));
            this.ETSMPLCD.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETSMPLCD_ChooseFromListAfter);
            this.OnCustomInitialize();

        }



        public override void OnInitializeFormEvents()
        {
        }


        private void OnCustomInitialize()
        {

        }

        private void LKSMPLCD_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.EditText ETSMPLCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETSMPLCD").Specific;
            string sampleCode = ETSMPLCD.Value.Trim();
            SampleMaster sampleMaster = new SampleMaster();
            sampleMaster.Show();
            //styleMaster. = Global.G_UI_Application.Forms.ActiveForm;
            SAPbouiCOM.Form cForm = Application.SBO_Application.Forms.Item("FIL_FRM_SMPLMSTR");
            try
            {
                cForm.Freeze(true);
                cForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                cForm.Items.Item("ETSLCODE").Enabled = true;
                SAPbouiCOM.EditText cETSLCODE = (SAPbouiCOM.EditText)cForm.Items.Item("ETSLCODE").Specific;
                cETSLCODE.Value = sampleCode;
                cForm.Items.Item("1").Click();
                cForm.Items.Item("FOLSIZE").Click();
                cForm.Freeze(false);
            }
            catch (Exception ex)
            {
                cForm.Freeze(false);
            }

        }

        private void GRDVERSN_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            try
            {
                if (pVal.Row < 0)
                    return;

                int ret = Application.SBO_Application.MessageBox(
                    "Are you sure you want see the Version Details?",
                    1, "OK", "Cancel");

                if (ret != 1)
                    return;

                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("GRDVERSN").Specific;
                SAPbouiCOM.DataTable oDT = oGrid.DataTable;
                int row = pVal.Row;

                SelDocEntry = oDT.GetValue("DocEntry", row).ToString().Trim();
                SelDocNum = oDT.GetValue("DocNum", row).ToString().Trim();
                SelVersion = oDT.GetValue("Version", row).ToString().Trim();
                SelLogInst = oDT.GetValue("MAX_LOGINST", row).ToString().Trim();

                string check = $@"
                                SELECT 
                                    CASE 
                                        WHEN MAX(""LogInst"") = '{SelLogInst}' THEN 1 
                                        ELSE 0
                                    END AS ""IsEqual""
                                FROM ""@AFIL_DH_PRECOSTING""
                                WHERE ""DocEntry"" = '{SelDocEntry}';";

                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(check);

                if (!rs.EoF)
                {
                    int isEqual = Convert.ToInt32(rs.Fields.Item("IsEqual").Value);

                    if (isEqual == 1)
                    {
                        Application.SBO_Application.MessageBox(
                            "You are currently in this version.");

                        return;
                    }
                    else
                    {
                        Application.SBO_Application.StatusBar.SetText(
                           $"Loading Version {SelVersion}",
                           SAPbouiCOM.BoMessageTime.bmt_Short,
                           SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                        oForm.Freeze(true);
                        ClearForm(oForm);
                        oForm.PaneLevel = 1;
                        LoadHeaderVersion(oForm, SelDocEntry, SelDocNum, SelVersion, SelLogInst);
                        LoadComponentMatrix(oForm, SelDocEntry, SelVersion, SelLogInst);
                        LoadOtherCostMatrix(oForm, SelDocEntry, SelVersion, SelLogInst);
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE;
                    }
                }


            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "GRDVERSN_DoubleClickAfter Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Long,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private void ClearForm(SAPbouiCOM.Form oForm)
        {
            // Clear EditTexts
            ((SAPbouiCOM.EditText)oForm.Items.Item("ETDOCTRY").Specific).Value = "";
            ((SAPbouiCOM.EditText)oForm.Items.Item("ETDOCNUM").Specific).Value = "";
            ((SAPbouiCOM.EditText)oForm.Items.Item("ETVERSON").Specific).Value = "";
            ((SAPbouiCOM.EditText)oForm.Items.Item("ETSMPLCD").Specific).Value = "";
            ((SAPbouiCOM.EditText)oForm.Items.Item("ETSMPLNM").Specific).Value = "";
            ((SAPbouiCOM.EditText)oForm.Items.Item("ETBUYER").Specific).Value = "";
            ((SAPbouiCOM.EditText)oForm.Items.Item("ETBYRNM").Specific).Value = "";
            ((SAPbouiCOM.EditText)oForm.Items.Item("ETCURR").Specific).Value = "";
            ((SAPbouiCOM.EditText)oForm.Items.Item("ETTCNAMT").Specific).Value = "";
            ((SAPbouiCOM.EditText)oForm.Items.Item("ETDOCNUM").Specific).Value = "";
            ((SAPbouiCOM.EditText)oForm.Items.Item("ETDATE").Specific).Value = "";

            // Clear Matrices
            SAPbouiCOM.Matrix mtxComp = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXCMPNT").Specific;
            SAPbouiCOM.Matrix mtxOth = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXOTCST").Specific;
            SAPbouiCOM.DataTable oDT = oForm.DataSources.DataTables.Item("DT_VERSN");
            mtxComp.Clear();
            mtxOth.Clear();
            oDT.Clear();
        }
        private void LoadHeaderVersion(SAPbouiCOM.Form oForm, string docEntry, string docNum, string version,string log)
        {
            string sql = $@"
                            SELECT ""DocEntry"", ""DocNum"" ,""U_VERSION"", ""U_SMPLCODE"" ,""U_SMPLDESC"",
                                    ""U_CARDCODE"" ,""U_CARDNAME"", ""U_CURRENCY"" ,""U_TOTCONAMT"", ""U_DOCDATE""
                            FROM ""@AFIL_DH_PRECOSTING"" 
                            WHERE ""DocEntry""='{docEntry}' 
                              AND ""DocNum""='{docNum}'
                              AND ""U_VERSION""='{version}'
                              AND ""LogInst""='{log}'";

            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(sql);

            if (!rs.EoF)
            {
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETDOCTRY").Specific).Value = rs.Fields.Item("DocEntry").Value.ToString();
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETDOCNUM").Specific).Value = rs.Fields.Item("DocNum").Value.ToString();
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETVERSON").Specific).Value = rs.Fields.Item("U_VERSION").Value.ToString();
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETSMPLCD").Specific).Value = rs.Fields.Item("U_SMPLCODE").Value.ToString();
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETSMPLNM").Specific).Value = rs.Fields.Item("U_SMPLDESC").Value.ToString();
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETBUYER").Specific).Value = rs.Fields.Item("U_CARDCODE").Value.ToString();
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETBYRNM").Specific).Value = rs.Fields.Item("U_CARDNAME").Value.ToString();
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETCURR").Specific).Value = rs.Fields.Item("U_CURRENCY").Value.ToString();
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETTCNAMT").Specific).Value = rs.Fields.Item("U_TOTCONAMT").Value.ToString();
                DateTime docDate = (DateTime)rs.Fields.Item("U_DOCDATE").Value;
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETDATE").Specific).Value = docDate.ToString("yyyyMMdd");

            }
        }
        private void LoadComponentMatrix(SAPbouiCOM.Form oForm, string docEntry, string version, string log)
        {
            SAPbouiCOM.Matrix mtx =
                (SAPbouiCOM.Matrix)oForm.Items.Item("MTXCMPNT").Specific;

            // Clear matrix
            mtx.Clear();

            string sql = $@"
                            SELECT
                                ""LineId"",
                                ""U_ROUTSTAG"",
                                ""U_ROUTSTAGN"",
                                ""U_COMPSTAG"",
                                ""U_COMPSTAGN"",
                                ""U_UOM"",
                                ""U_QTY"",
                                ""U_CONAMT"",
                                ""U_VERSION""
                            FROM ""@AFIL_DR_PRECOSTCOMP""
                            WHERE ""DocEntry""='{docEntry}'
                              AND ""U_VERSION""='{version}'
                              AND ""LogInst""='{log}'";

            SAPbobsCOM.Recordset rs =
                (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            rs.DoQuery(sql);

            int row = 1;

            while (!rs.EoF)
            {
                mtx.AddRow();

                ((SAPbouiCOM.EditText)mtx.Columns.Item("#").Cells.Item(row).Specific).Value = rs.Fields.Item("LineId").Value.ToString();
                SAPbouiCOM.ComboBox cb = (SAPbouiCOM.ComboBox)mtx.Columns.Item("CLRSTGCD").Cells.Item(row).Specific;

                string val = rs.Fields.Item("U_ROUTSTAG").Value.ToString().Trim();

                if (!string.IsNullOrEmpty(val))
                {
                    cb.Select(val, SAPbouiCOM.BoSearchKey.psk_ByValue);
                }

                ((SAPbouiCOM.EditText)mtx.Columns.Item("CLRSTGDS").Cells.Item(row).Specific).Value = rs.Fields.Item("U_ROUTSTAGN").Value.ToString();
                ((SAPbouiCOM.EditText)mtx.Columns.Item("CLCSTGCD").Cells.Item(row).Specific).Value = rs.Fields.Item("U_COMPSTAG").Value.ToString();
                ((SAPbouiCOM.EditText)mtx.Columns.Item("CLCSTGNM").Cells.Item(row).Specific).Value = rs.Fields.Item("U_COMPSTAGN").Value.ToString();
                ((SAPbouiCOM.EditText)mtx.Columns.Item("CLUOM").Cells.Item(row).Specific).Value = rs.Fields.Item("U_UOM").Value.ToString();
                ((SAPbouiCOM.EditText)mtx.Columns.Item("CLQTY").Cells.Item(row).Specific).Value = rs.Fields.Item("U_QTY").Value.ToString();
                ((SAPbouiCOM.EditText)mtx.Columns.Item("CLAMT").Cells.Item(row).Specific).Value = rs.Fields.Item("U_CONAMT").Value.ToString();
                ((SAPbouiCOM.EditText)mtx.Columns.Item("CLVERSN").Cells.Item(row).Specific).Value = rs.Fields.Item("U_VERSION").Value.ToString();

                row++;
                rs.MoveNext();
            }
        }


        private void LoadOtherCostMatrix(SAPbouiCOM.Form oForm, string docEntry, string version, string log)
        {
            SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXOTCST").Specific;
            mtx.Clear();

            string sql = $@"
                            SELECT 
                                    ""LineId"",
                                    ""U_ALCCODE"",
                                    ""U_ALCNAME"",
                                    ""U_OHCONAMT"",
                                    ""U_VERSION""
                            FROM ""@AFIL_DR_PRECOSTOTHR"" 
                            WHERE ""DocEntry""='{docEntry}' 
                                AND ""U_VERSION""='{version}'
                                AND ""LogInst""='{log}'";

            SAPbobsCOM.Recordset rs =
                (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            rs.DoQuery(sql);

            int row = 1;

            while (!rs.EoF)
            {
                mtx.AddRow();

                ((SAPbouiCOM.EditText)mtx.Columns.Item("#").Cells.Item(row).Specific).Value = rs.Fields.Item("LineId").Value.ToString();
                ((SAPbouiCOM.EditText)mtx.Columns.Item("CLCSTHCD").Cells.Item(row).Specific).Value = rs.Fields.Item("U_ALCCODE").Value.ToString();
                ((SAPbouiCOM.EditText)mtx.Columns.Item("CLCSTHNM").Cells.Item(row).Specific).Value = rs.Fields.Item("U_ALCNAME").Value.ToString();
                ((SAPbouiCOM.EditText)mtx.Columns.Item("CLAMT").Cells.Item(row).Specific).Value = rs.Fields.Item("U_OHCONAMT").Value.ToString();
                ((SAPbouiCOM.EditText)mtx.Columns.Item("CLVERSN").Cells.Item(row).Specific).Value = rs.Fields.Item("U_VERSION").Value.ToString();
                row++;
                rs.MoveNext();
            }
        }

        private void ADDButton_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            //Series Initialization
            SAPbouiCOM.DBDataSource oDBH = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@FIL_DH_PRECOSTING");
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                SAPbouiCOM.ComboBox ocmb = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBSERIES").Specific;
                Global.GFunc.LoadComboBoxSeries(ocmb, "FIL_D_PRECOSTING");
                string ocmbvalue = ocmb.Selected.Value;
                long docno = oForm.BusinessObject.GetNextSerialNumber(ocmbvalue, "FIL_D_PRECOSTING");
                oDBH.SetValue("DocNum", 0, docno.ToString());
            }
            //Date
            ((SAPbouiCOM.EditText)oForm.Items.Item("ETDATE").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
            ((SAPbouiCOM.EditText)oForm.Items.Item("ETVERSON").Specific).Value = "1"; //Default version 
            //Enable off
            SetItemsEnabled(oForm, false, "ETSMPLNM", "ETBUYER", "ETBYRNM", "ETDOCNUM", "ETDATE", "ETVERSON");
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

        private void FOLVERSN_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.EditText etDocEntry, etDocNum, etVersion;
            SAPbouiCOM.Grid oGrid;
            SAPbouiCOM.DataTable oDT;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                etDocEntry = (SAPbouiCOM.EditText)oForm.Items.Item("ETDOCTRY").Specific;
                etDocNum = (SAPbouiCOM.EditText)oForm.Items.Item("ETDOCNUM").Specific;
                etVersion = (SAPbouiCOM.EditText)oForm.Items.Item("ETVERSON").Specific;

                // Parse values
                string docEntry = etDocEntry.Value.Trim();
                string docNum = etDocNum.Value.Trim();
                int version = Convert.ToInt32(string.IsNullOrEmpty(etVersion.Value) ? "0" : etVersion.Value);

                if (version <= 1)
                {
                    Application.SBO_Application.StatusBar
                        .SetText("Version must be greater than 1",
                                 SAPbouiCOM.BoMessageTime.bmt_Short,
                                 SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                    oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("GRDVERSN").Specific;
                    oDT = oForm.DataSources.DataTables.Item("DT_VERSN");
                    oDT.Clear();
                    return;
                }

                oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("GRDVERSN").Specific;
                oDT = oForm.DataSources.DataTables.Item("DT_VERSN");
                oDT.Clear();
                string sQuery = $@"
                                    SELECT
                                        T.""DocEntry"",
                                        T.""DocNum"",
                                        T.""Creator"",
                                        T.""CreateDate"",
                                        T.""UpdateDate"",
                                        T.""U_VERSION"" AS ""Version"",
                                        T.""LogInst"" AS ""MAX_LOGINST""
                                    FROM ""@AFIL_DH_PRECOSTING"" T
                                    JOIN (
                                        SELECT
                                            ""DocEntry"",
                                            ""DocNum"",
                                            ""U_VERSION"",
                                            MAX(""LogInst"") AS ""MAX_LOGINST""
                                        FROM ""@AFIL_DH_PRECOSTING""
                                        WHERE ""DocEntry"" = '{docEntry}'
                                          AND ""DocNum""   = '{docNum}'
                                        GROUP BY
                                            ""DocEntry"",
                                            ""DocNum"",
                                            ""U_VERSION""
                                    ) M
                                    ON  T.""DocEntry""  = M.""DocEntry""
                                    AND T.""DocNum""    = M.""DocNum""
                                    AND T.""U_VERSION"" = M.""U_VERSION""
                                    AND T.""LogInst""   = M.""MAX_LOGINST""
                                    WHERE T.""DocEntry"" = '{docEntry}'
                                      AND T.""DocNum""   = '{docNum}'
                                    ORDER BY T.""U_VERSION""";


                oDT.ExecuteQuery(sQuery);
                oGrid.DataTable = oDT;
                oGrid.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar
                    .SetText(ex.Message,
                             SAPbouiCOM.BoMessageTime.bmt_Long,
                             SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }




        private void BTNVRNUP_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

            try
            {
                //Confirmation
                int ret = Application.SBO_Application.MessageBox(
                    "Are you sure you want to increase the version?",
                    1, "OK", "Cancel");

                if (ret != 1)
                    return;

                //Read values from edit texts
                string docEntry = ((SAPbouiCOM.EditText)oForm.Items.Item("ETDOCTRY").Specific).Value.Trim();
                string docNum = ((SAPbouiCOM.EditText)oForm.Items.Item("ETDOCNUM").Specific).Value.Trim();
                string version = ((SAPbouiCOM.EditText)oForm.Items.Item("ETVERSON").Specific).Value.Trim();

                if (string.IsNullOrWhiteSpace(docEntry) || string.IsNullOrWhiteSpace(docNum) || string.IsNullOrWhiteSpace(version))
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "DocEntry, DocNum, and Version are required.",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    return;
                }

                //Check if current version exists in DB
                string sql = $@"
                                Select 1 from ""@FIL_DH_PRECOSTING"" 
                                    where ""DocEntry"" = '{docEntry}'
                                    and ""DocNum""   = '{docNum}'
                                    and ""U_VERSION""= '{version}'";

                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                //If exists
                if (rs.RecordCount > 0)
                {
                    if (!int.TryParse(version, out int v))
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            "Version is not numeric, cannot increment.",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return;
                    }

                    v++;
                    ((SAPbouiCOM.EditText)oForm.Items.Item("ETVERSON").Specific).Value = v.ToString();

                    // Matrix Other Cost
                    SAPbouiCOM.Matrix mtxOtherCost = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXOTCST").Specific;
                    UpdateMatrixVersion(mtxOtherCost, "CLVERSN", v);

                    // Matrix Component
                    SAPbouiCOM.Matrix mtxComponent = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXCMPNT").Specific;
                    UpdateMatrixVersion(mtxComponent, "CLVERSN", v);

                    // =================================================


                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }

                    Application.SBO_Application.StatusBar.SetText(
                        "Version increased successfully.",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                else
                {
                    //Otherwise message
                    Application.SBO_Application.MessageBox(
                        "You have to update/save this version to DB first, then you can increase the version.",
                        1, "OK");
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "BTNVRNUP_PressedAfter Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void UpdateMatrixVersion(SAPbouiCOM.Matrix oMatrix, string colId, int version)
        {
            if (oMatrix == null || oMatrix.RowCount == 0)
                return;

            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                SAPbouiCOM.EditText et =
                    (SAPbouiCOM.EditText)oMatrix.Columns.Item(colId).Cells.Item(i).Specific;

                et.Value = version.ToString();
            }

            // Flush if matrix is bound
            oMatrix.FlushToDataSource();
        }



        private void BTNLCSTH_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

            try
            {
                // Confirmation
                int ret = Application.SBO_Application.MessageBox(
                    "Are you sure you want to refresh Other Cost Head list?",
                    1, "OK", "Cancel");

                if (ret != 1) // 1 = OK
                    return;

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXOTCST").Specific;

                //Backup existing user inputs 
                Dictionary<string, string> amtByCode = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    string code = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLCSTHCD").Cells.Item(i).Specific).Value.Trim();
                    if (string.IsNullOrWhiteSpace(code)) continue;

                    string amt = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLAMT").Cells.Item(i).Specific).Value.Trim();

                    //duplicate filter
                    if (!amtByCode.ContainsKey(code))
                        amtByCode.Add(code, amt);
                    else if (!string.IsNullOrWhiteSpace(amt))
                        amtByCode[code] = amt;
                }

                string sql = @"Select ""AlcCode"",""AlcName"" from ""OALC""";
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                if (rs.RecordCount == 0)
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "No active cost head found.",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    return;
                }

                oForm.Freeze(true);

                // 4) Clear matrix rows (we will rebuild from query) BUT restore CLAMT from backup
                // Safer clear:
                while (oMatrix.RowCount > 0)
                    oMatrix.DeleteRow(1);

                int row = 1;
                rs.MoveFirst();
                while (!rs.EoF)
                {
                    oMatrix.AddRow();

                    string code = Convert.ToString(rs.Fields.Item("Code").Value).Trim();
                    string name = Convert.ToString(rs.Fields.Item("Name").Value).Trim();

                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("#").Cells.Item(row).Specific).Value = row.ToString();
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLCSTHCD").Cells.Item(row).Specific).Value = code;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLCSTHNM").Cells.Item(row).Specific).Value = name;

                    // 5) Restore amount if user already entered earlier
                    if (amtByCode.TryGetValue(code, out string oldAmt) && !string.IsNullOrWhiteSpace(oldAmt))
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLAMT").Cells.Item(row).Specific).Value = oldAmt;

                    row++;
                    rs.MoveNext();
                }

                oMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "BTNLCSTH_PressedAfter Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                try { oForm.Freeze(false); } catch { }
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

        private bool ValidateForm(ref SAPbouiCOM.Form oForm, ref bool BubbleEvent)
        {
            string sampleCode = oForm.DataSources.DBDataSources.Item("@FIL_DH_PRECOSTING").GetValue("U_SMPLCODE", 0);
            string buyer = oForm.DataSources.DBDataSources.Item("@FIL_DH_PRECOSTING").GetValue("U_CARDCODE", 0);
            string curr = oForm.DataSources.DBDataSources.Item("@FIL_DH_PRECOSTING").GetValue("U_CURRENCY", 0);

            if (sampleCode == "")
            {
                Global.GFunc.ShowError("Enter Sample Code");
                oForm.ActiveItem = "ETSMPLCD";
                return BubbleEvent = false;
            }
            else if (buyer == "")
            {
                Global.GFunc.ShowError("Enter Buyer");
                oForm.ActiveItem = "ETBUYER";
                return BubbleEvent = false;
            }
            else if (curr == "")
            {
                Global.GFunc.ShowError("Enter Currency");
                oForm.ActiveItem = "ETCURR";
                return BubbleEvent = false;
            }

            PreventEmptyLastRow(oForm, "@FIL_DR_PRECOSTCOMP", MTXCMPNT, "U_ROUTSTAG");
            ValidateRouteCompStage(oForm, ref BubbleEvent);
            if (!BubbleEvent)
            {
                Global.GFunc.ShowError("Route Stage or Component Stage Missing");
                //AddLineIfLastRowHasValue(oForm, "MTXCMPNT", "@FIL_DR_PRECOSTCOMP", "U_ROUTSTAG");
                EnsureLine(oForm, "MTXCMPNT", "@FIL_DR_PRECOSTCOMP");
                return BubbleEvent;
            }

            return BubbleEvent;
        }

        private bool ValidateRouteCompStage(SAPbouiCOM.Form oForm, ref bool BubbleEvent)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXCMPNT").Specific;

                oMatrix.FlushToDataSource();
                SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("@FIL_DR_PRECOSTCOMP");


                if (ds.Size == 0 ||
                    (ds.Size == 1 &&
                     string.IsNullOrWhiteSpace((ds.GetValue("U_ROUTSTAG", 0) ?? "").Trim()) &&
                     string.IsNullOrWhiteSpace((ds.GetValue("U_COMPSTAG", 0) ?? "").Trim())))
                {
                    Global.GFunc.ShowError("Component matrix is empty. Please add at least one row.");
                    oForm.Items.Item("MTXCMPNT").Click();
                    BubbleEvent = false;
                    return BubbleEvent;
                }

                for (int i = 0; i < ds.Size; i++)
                {
                    string routeStag = (ds.GetValue("U_ROUTSTAG", i) ?? "").Trim();
                    string compStag = (ds.GetValue("U_COMPSTAG", i) ?? "").Trim();

                    if (string.IsNullOrWhiteSpace(routeStag))
                    {
                        Global.GFunc.ShowError($"Route Stage is mandatory. (Row: {i + 1})");
                        oForm.Items.Item("MTXCMPNT").Click();
                        oMatrix.Columns.Item("U_ROUTSTAG").Cells.Item(i + 1).Click();
                        BubbleEvent = false;
                        return BubbleEvent;
                    }

                    if (string.IsNullOrWhiteSpace(compStag))
                    {
                        Global.GFunc.ShowError($"Component Stage is mandatory. (Row: {i + 1})");
                        oForm.Items.Item("MTXCMPNT").Click();
                        oMatrix.Columns.Item("U_COMPSTAG").Cells.Item(i + 1).Click();
                        BubbleEvent = false;
                        return BubbleEvent;
                    }
                }

                return BubbleEvent;
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "ValidateRouteCompStage Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                BubbleEvent = false;
                return BubbleEvent;
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

        private void FOLOTCST_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            try
            {
                //if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                //    return;

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXOTCST").Specific;
                bool isEmpty = (oMatrix.RowCount == 0);
                if (!isEmpty)
                {
                    var firstCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLCSTHCD").Cells.Item(1).Specific).Value;
                    isEmpty = string.IsNullOrWhiteSpace(firstCode);
                }
                if (!isEmpty)
                    return;
                string version = ((SAPbouiCOM.EditText)oForm.Items.Item("ETVERSON").Specific).Value.Trim();
                string sql = @"Select ""AlcCode"",""AlcName"" from ""OALC""";
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                if (rs.RecordCount == 0)
                    return;

                oForm.Freeze(true);
                int row = 1;
                rs.MoveFirst();
                while (!rs.EoF)
                {
                    if (row > oMatrix.RowCount)
                        oMatrix.AddRow();

                    string code = Convert.ToString(rs.Fields.Item("AlcCode").Value).Trim();
                    string name = Convert.ToString(rs.Fields.Item("AlcName").Value).Trim();
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("#").Cells.Item(row).Specific).Value = row.ToString();
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLCSTHCD").Cells.Item(row).Specific).Value = code;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLCSTHNM").Cells.Item(row).Specific).Value = name;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLVERSN").Cells.Item(row).Specific).Value = version;

                    row++;
                    rs.MoveNext();
                }
                oForm.Freeze(false);
                oMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                Application.SBO_Application.StatusBar.SetText(
                    "FOLOTCST_ClickAfter Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void MTXOTCST_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID != "CLAMT") return;
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                UpdateTotalAmountBothMatrices(oForm);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error on Amount Calculation: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }


        private void MTXCMPNT_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID != "CLAMT") return;
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                UpdateTotalAmountBothMatrices(oForm);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error on Amount Calculation: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }


        private void UpdateTotalAmountBothMatrices(SAPbouiCOM.Form oForm)
        {
            double total = 0;

            total += GetMatrixColumnSum(oForm, "MTXCMPNT", "CLAMT");
            total += GetMatrixColumnSum(oForm, "MTXOTCST", "CLAMT");

            ((SAPbouiCOM.EditText)oForm.Items.Item("ETTCNAMT").Specific).Value =
                total.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);
        }

        private double GetMatrixColumnSum(SAPbouiCOM.Form oForm, string matrixId, string colId)
        {
            double sum = 0;

            if (!oForm.Items.Item(matrixId).Visible) { /* optional */ }

            SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixId).Specific;

            for (int i = 1; i <= mtx.RowCount; i++)
            {
                string valStr = ((SAPbouiCOM.EditText)mtx.Columns.Item(colId).Cells.Item(i).Specific).Value;
                if (string.IsNullOrWhiteSpace(valStr)) continue;

                valStr = valStr.Trim().Replace(",", "");

                if (double.TryParse(valStr, System.Globalization.NumberStyles.Any,
                                    System.Globalization.CultureInfo.InvariantCulture, out double amt))
                    sum += amt;
                else if (double.TryParse(valStr, out amt))
                    sum += amt;
            }

            return sum;
        }


        private void MTXCMPNT_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_CMPN")
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
                    "Error filtering Size CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }


        }

        private void MTXCMPNT_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXCMPNT").Specific;
            SAPbouiCOM.DBDataSource DBDataSourceLine = oForm.DataSources.DBDataSources.Item("@FIL_DR_PRECOSTCOMP");
            int row = pVal.Row;

            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0) return;

            string Code = dt.GetValue("Code", 0).ToString().Trim();
            string Name = dt.GetValue("Name", 0).ToString().Trim();
            string upmApl = dt.GetValue("U_UOMAPPLY", 0).ToString().Trim();
            if (upmApl.Equals("Y"))
            {
                string uom = dt.GetValue("U_UOM", 0).ToString().Trim();
                oMatrix.SetCellWithoutValidation(row, "CLUOM", Code);
            }

            oMatrix.SetCellWithoutValidation(row, "CLCSTGCD", Code);
            oMatrix.SetCellWithoutValidation(row, "CLCSTGNM", Name);

            AddLineIfLastRowHasValue(oForm, "MTXCMPNT", "@FIL_DR_PRECOSTCOMP", "U_COMPSTAG");

        }

        private void ETSMPLCD_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0) return;

            try
            {
                if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    string Code = dt.GetValue("U_SMPLCODE", 0).ToString().Trim();
                    string Name = dt.GetValue("U_SMPLDESC", 0).ToString().Trim();
                    string buyCode = dt.GetValue("U_CARDCODE", 0).ToString().Trim();
                    string buyName = dt.GetValue("U_CARDNAME", 0).ToString().Trim();
                    string route = dt.GetValue("U_ROUTSTAG", 0).ToString().Trim();

                    ETSMPLCD.Value = Code;
                    ETSMPLNM.Value = Name;
                    ETBUYER.Value = buyCode;
                    ETBYRNM.Value = buyName;

                    EnsureLine(oForm, "MTXCMPNT", "@FIL_DR_PRECOSTCOMP");
                    EnsureLine(oForm, "MTXOTCST", "@FIL_DR_PRECOSTOTHR");
                    LoadRouteWiseComboToMatrixColumn(oForm, "MTXCMPNT", "CLRSTGCD", route);

                    SAPbouiCOM.Item ETCUSCOD = oForm.Items.Item("ETBUYER");
                    ETCUSCOD.Enabled = true;
                    SAPbouiCOM.EditText ECUSCOD = (SAPbouiCOM.EditText)ETCUSCOD.Specific;
                    ECUSCOD.ChooseFromListUID = "CFL_OCRD";
                    ECUSCOD.ChooseFromListAlias = "CardCode";
                }

            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox("Error: " + ex.Message);
            }

        }

        private void MTXCMPNT_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXCMPNT").Specific;
                string version = ((SAPbouiCOM.EditText)oForm.Items.Item("ETVERSON").Specific).Value.Trim();
                if (pVal.Row <= 0) return;

                if (pVal.ColUID == "CLRSTGCD")
                {
                    int row = pVal.Row;
                    SAPbouiCOM.ComboBox cb = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("CLRSTGCD").Cells.Item(row).Specific;
                    string desc = cb.Selected.Description.Trim();
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLRSTGDS").Cells.Item(row).Specific).Value = desc;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLVERSN").Cells.Item(row).Specific).Value = version;
                    oMatrix.FlushToDataSource();

                    AddLineIfLastRowHasValue(oForm, "MTXCMPNT", "@FIL_DR_PRECOSTCOMP", "U_ROUTSTAG");
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox("Error: " + ex.Message);
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

        private void ETCURR_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("CurrCode", 0).ToString().Trim();
            SAPbouiCOM.EditText ETBCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETCURR").Specific;
            ETBCD.Value = Code;

        }

        private void ETBUYER_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;


            string buyCode = dt.GetValue("CardCode", 0).ToString().Trim();
            string buyName = dt.GetValue("CardName", 0).ToString().Trim();

            SAPbouiCOM.EditText ETBCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETBUYER").Specific;
            ETBCD.Value = buyCode;
            SAPbouiCOM.EditText ETBNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETBYRNM").Specific;
            ETBNM.Value = buyName;

        }


        private void LoadRouteWiseComboToMatrixColumn(SAPbouiCOM.Form oForm, string matrixId, string comboColId, string routeCode)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixId).Specific;
            SAPbouiCOM.Column oCol = oMatrix.Columns.Item(comboColId);

            if (oCol.Type != SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                throw new Exception($"Column {comboColId} is not a ComboBox column.");

            //Remove Prev Values
            while (oCol.ValidValues.Count > 0)
                oCol.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);


            string query = $@"Select Distinct ""U_Code"", ""U_Name""
                                from ""@FIL_MR_RSM1""
                                where ""Code"" = '{routeCode}'";

            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(query);
            oCol.ValidValues.Add("", "");

            while (!rs.EoF)
            {
                string v = rs.Fields.Item("U_Code").Value.ToString().Trim();
                string d = rs.Fields.Item("U_Name").Value.ToString().Trim();

                if (!ValidValueExists(oCol, v))
                    oCol.ValidValues.Add(v, d);

                rs.MoveNext();
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
        }

        private bool ValidValueExists(SAPbouiCOM.Column col, string value)
        {
            for (int i = 0; i < col.ValidValues.Count; i++)
            {
                if (col.ValidValues.Item(i).Value == value)
                    return true;
            }
            return false;
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
