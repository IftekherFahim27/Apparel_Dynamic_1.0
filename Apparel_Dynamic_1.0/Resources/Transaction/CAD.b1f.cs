using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Apparel_Dynamic_1._0.Helper;

namespace Apparel_Dynamic_1._0.Resources.Transaction
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Transaction.CAD", "Resources/Transaction/CAD.b1f")]
    class CAD : UserFormBase
    {
        public CAD()
        {
        }

        private SAPbouiCOM.StaticText STSTATUS, STSTYLCD, STSTYLDS, STMERCHN, STDOCNUM, STDOCDAT, STDRFTNO, STSLCLR, STCDCLR;
        private SAPbouiCOM.EditText ETSLCLR, ETSTYLCD, ETMERHN,ETSTYLDS, ETDOCNUM, ETMERCNM, ETDONTRY, ETDOCDAT, ETDRFTNO, ETCDCLR, ETDOCTRY, ETSTNTRY;


        private SAPbouiCOM.ComboBox CBDOCNUM, CBSTATUS;
        private SAPbouiCOM.Folder FOLMERCON, FOLCANCON, FOLTEMP;
        private SAPbouiCOM.Matrix MTXMRCON, MTXCDCLR, MTXCDCON;
        private SAPbouiCOM.Grid GRDCDCON, GRDSIZE, GRDSCLR;
        private SAPbouiCOM.Button ADDButton, CancelButton, BTNFETCH, BTNLDCAD, BTNSAVE;


        public override void OnInitializeComponent()
        {
            //         Static text
            this.STSTATUS = ((SAPbouiCOM.StaticText)(this.GetItem("STSTATUS").Specific));
            this.STSTYLCD = ((SAPbouiCOM.StaticText)(this.GetItem("STSTYLCD").Specific));
            this.STSTYLDS = ((SAPbouiCOM.StaticText)(this.GetItem("STSTYLDS").Specific));
            this.STMERCHN = ((SAPbouiCOM.StaticText)(this.GetItem("STMERHN").Specific));
            this.STDOCNUM = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCNUM").Specific));
            this.STDOCDAT = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCDAT").Specific));
            this.STDRFTNO = ((SAPbouiCOM.StaticText)(this.GetItem("STDRFTNO").Specific));
            this.STSLCLR = ((SAPbouiCOM.StaticText)(this.GetItem("STSLCLR").Specific));
            this.STCDCLR = ((SAPbouiCOM.StaticText)(this.GetItem("STCDCLR").Specific));
            //         Edittext
            this.ETSLCLR = ((SAPbouiCOM.EditText)(this.GetItem("ETSLCLR").Specific));
            this.ETSTYLCD = ((SAPbouiCOM.EditText)(this.GetItem("ETSTYLCD").Specific));
            this.ETSTYLCD.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETSTYLCD_ChooseFromListAfter);
            this.ETSTYLDS = ((SAPbouiCOM.EditText)(this.GetItem("ETSTYLDS").Specific));
            this.ETDOCNUM = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCNUM").Specific));
            this.ETDOCDAT = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCDAT").Specific));
            this.ETDRFTNO = ((SAPbouiCOM.EditText)(this.GetItem("ETDRFTNO").Specific));
            this.ETDRFTNO.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETDRFTNO_ChooseFromListAfter);
            this.ETDRFTNO.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETDRFTNO_ChooseFromListBefore);
            this.ETCDCLR = ((SAPbouiCOM.EditText)(this.GetItem("ETCDCLR").Specific));
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            //         Combo box
            this.CBDOCNUM = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSERIES").Specific));
            this.CBSTATUS = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSTATUS").Specific));
            //         Folder
            this.FOLMERCON = ((SAPbouiCOM.Folder)(this.GetItem("FOLMERCON").Specific));
            this.FOLCANCON = ((SAPbouiCOM.Folder)(this.GetItem("FOLCANCN").Specific));
            this.FOLTEMP = ((SAPbouiCOM.Folder)(this.GetItem("FOLTEMP").Specific));
            //         Matrix
            this.MTXMRCON = ((SAPbouiCOM.Matrix)(this.GetItem("MTXMRCON").Specific));
            this.MTXCDCLR = ((SAPbouiCOM.Matrix)(this.GetItem("MTXCDCLR").Specific));
            this.MTXCDCON = ((SAPbouiCOM.Matrix)(this.GetItem("MTXCDCON").Specific));
            //         Grid
            this.GRDCDCON = ((SAPbouiCOM.Grid)(this.GetItem("GRDCDCON").Specific));
            this.GRDSIZE = ((SAPbouiCOM.Grid)(this.GetItem("GRDSIZE").Specific));
            //         Button
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.BTNFETCH = ((SAPbouiCOM.Button)(this.GetItem("BTNFETCH").Specific));
            this.BTNFETCH.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.BTNFETCH_PressedAfter);
            this.BTNLDCAD = ((SAPbouiCOM.Button)(this.GetItem("BTNLDCAD").Specific));
            this.BTNSAVE = ((SAPbouiCOM.Button)(this.GetItem("BTNSAVE").Specific));
            this.GRDSCLR = ((SAPbouiCOM.Grid)(this.GetItem("GRDSCLR").Specific));
            this.GRDSCLR.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.GRDSCLR_DoubleClickAfter);
            this.ETSTNTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETSTNTRY").Specific));
            this.ETMERCNM = ((SAPbouiCOM.EditText)(this.GetItem("ETMERCNM").Specific));
            this.ETDONTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDONTRY").Specific));
            this.ETMERHN = ((SAPbouiCOM.EditText)(this.GetItem("ETMERHN").Specific));
            this.ETMERHN.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETMERHN_ChooseFromListAfter);
            this.OnCustomInitialize();

        }

        public override void OnInitializeFormEvents()
        {
            this.ResizeAfter += new ResizeAfterHandler(this.Form_ResizeAfter);

        }


        private void OnCustomInitialize()
        {

        }


        private void ETMERHN_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("empID", 0).ToString().Trim();
            string Name = dt.GetValue("U_FNAME", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETMERHN").Specific;
            ETCD.Value = Code;
            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETMERCNM").Specific;
            ETNM.Value = Name;

        }



        private void GRDSCLR_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = null;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                oForm.Freeze(true);

                if (pVal.Row < 0)
                    return;

                string docEntry = ((SAPbouiCOM.EditText)oForm.Items.Item("ETDONTRY").Specific).Value.Trim();

                if (string.IsNullOrWhiteSpace(docEntry))
                {
                    Application.SBO_Application.MessageBox("Please select Draft Order first.");
                    return;
                }

                SAPbouiCOM.Grid grdClr = (SAPbouiCOM.Grid)oForm.Items.Item("GRDSCLR").Specific;
                int dtRow = grdClr.GetDataTableRowIndex(pVal.Row);

                if (dtRow < 0)
                    return;

                SAPbouiCOM.DataTable dtClr = oForm.DataSources.DataTables.Item("DT_CLR");

                string colourCode = dtClr.GetValue("Colour Code", dtRow).ToString().Trim();

                if (string.IsNullOrWhiteSpace(colourCode))
                {
                    Application.SBO_Application.MessageBox("Colour Code not found.");
                    return;
                }

                // Set selected colour code
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETSLCLR").Specific).Value = colourCode;

                string query = $@"
                                SELECT 
                                    T0.""U_FGSIZE"" AS ""Size Code"",
                                    T1.""Name""     AS ""Size Name"",
                                    SUM(T0.""Quantity"") AS ""Quantity""
                                FROM ""QUT1"" T0
                                LEFT JOIN ""@FIL_MH_SIZEMSTR"" T1
                                    ON T0.""U_FGSIZE"" = T1.""Code""
                                WHERE T0.""DocEntry"" = '{docEntry}'
                                  AND IFNULL(T0.""U_FGCOLOUR"", '') = '{colourCode.Replace("'", "''")}'
                                  AND IFNULL(T0.""U_FGSIZE"", '') <> ''
                                GROUP BY 
                                    T0.""U_FGSIZE"",
                                    T1.""Name""
                                ORDER BY 
                                    T0.""U_FGSIZE"" ";

                SAPbouiCOM.DataTable dtSize = oForm.DataSources.DataTables.Item("DT_SIZE");
                dtSize.ExecuteQuery(query);

                SAPbouiCOM.Grid grdSize = (SAPbouiCOM.Grid)oForm.Items.Item("GRDSIZE").Specific;
                grdSize.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox("Colour selection failed: " + ex.Message);
            }
            finally
            {
                if (oForm != null)
                    oForm.Freeze(false);
            }
        }

        private void BTNFETCH_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = null;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                oForm.Freeze(true);

                string docEntry = ((SAPbouiCOM.EditText)oForm.Items.Item("ETDONTRY").Specific).Value.Trim();

                if (string.IsNullOrWhiteSpace(docEntry))
                {
                    Application.SBO_Application.MessageBox("Please select Draft Order first.");
                    return;
                }

                string query = $@"
                                SELECT DISTINCT
                                    T0.""U_FGCOLOUR""  AS ""Colour Code"",
                                    T0.""U_FGCOLRNM""  AS ""Colour Name"",
                                    T0.""U_FGSIZE""    AS ""Size Code"",
                                    T1.""Name""        AS ""Size Name""
                                FROM ""QUT1"" T0
                                LEFT JOIN ""@FIL_MH_SIZEMSTR"" T1
                                    ON T0.""U_FGSIZE"" = T1.""Code""
                                WHERE T0.""DocEntry"" = '{docEntry}'
                                  AND IFNULL(T0.""U_FGCOLOUR"", '') <> ''
                                  AND IFNULL(T0.""U_FGSIZE"", '') <> ''
                                ORDER BY 
                                    T0.""U_FGCOLOUR"",
                                    T0.""U_FGSIZE"" ";

                SAPbobsCOM.Recordset rs =
                    (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                rs.DoQuery(query);

                // =========================
                // Clear Grid DataTables
                // =========================
                SAPbouiCOM.DataTable dtClr = oForm.DataSources.DataTables.Item("DT_CLR");
                SAPbouiCOM.DataTable dtSize = oForm.DataSources.DataTables.Item("DT_SIZE");

                dtClr.Rows.Clear();
                dtSize.Rows.Clear();

                // =========================
                // Clear Matrix DBDataSource
                // =========================
                SAPbouiCOM.DBDataSource dbClr =
                    oForm.DataSources.DBDataSources.Item("@FIL_DR_CADFABCL");

                dbClr.Clear();

                HashSet<string> colourSet = new HashSet<string>();
                HashSet<string> sizeSet = new HashSet<string>();

                int colourRow = 0;
                int sizeRow = 0;
                int matrixRow = 0;

                while (!rs.EoF)
                {
                    string colourCode = rs.Fields.Item("Colour Code").Value.ToString().Trim();
                    string colourName = rs.Fields.Item("Colour Name").Value.ToString().Trim();
                    string sizeCode = rs.Fields.Item("Size Code").Value.ToString().Trim();
                    string sizeName = rs.Fields.Item("Size Name").Value.ToString().Trim();

                    // =========================
                    // GRDSCLR + MTXCDCLR
                    // Distinct Colour Code
                    // =========================
                    if (!string.IsNullOrWhiteSpace(colourCode) && !colourSet.Contains(colourCode))
                    {
                        colourSet.Add(colourCode);

                        // Fill GRDSCLR DataTable
                        dtClr.Rows.Add();
                        dtClr.SetValue("Colour Code", colourRow, colourCode);
                        dtClr.SetValue("Colour Name", colourRow, colourName);
                        colourRow++;

                        // Fill MTXCDCLR DBDataSource
                        dbClr.InsertRecord(matrixRow);
                        dbClr.SetValue("LineId", matrixRow, (matrixRow + 1).ToString());
                        dbClr.SetValue("U_COLORCODE", matrixRow, colourCode);
                        dbClr.SetValue("U_COLORNAME", matrixRow, colourName);
                        matrixRow++;
                    }

                    // =========================
                    // GRDSIZE
                    // Distinct Size Code
                    // Quantity = 0
                    // =========================
                    if (!string.IsNullOrWhiteSpace(sizeCode) && !sizeSet.Contains(sizeCode))
                    {
                        sizeSet.Add(sizeCode);

                        dtSize.Rows.Add();
                        dtSize.SetValue("Size Code", sizeRow, sizeCode);
                        dtSize.SetValue("Size Name", sizeRow, sizeName);
                        dtSize.SetValue("Quantity", sizeRow, 0);
                        sizeRow++;
                    }

                    rs.MoveNext();
                }

                // =========================
                // Reload UI
                // =========================
                SAPbouiCOM.Grid grdClr = (SAPbouiCOM.Grid)oForm.Items.Item("GRDSCLR").Specific;
                SAPbouiCOM.Grid grdSize = (SAPbouiCOM.Grid)oForm.Items.Item("GRDSIZE").Specific;
                SAPbouiCOM.Matrix mtxClr = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXCDCLR").Specific;

                grdClr.AutoResizeColumns();
                grdSize.AutoResizeColumns();

                mtxClr.LoadFromDataSource();
                mtxClr.AutoResizeColumns();

                Application.SBO_Application.StatusBar.SetText(
                    "Colour and Size fetched successfully.",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success
                );
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox("Fetch failed: " + ex.Message);
            }
            finally
            {
                if (oForm != null)
                    oForm.Freeze(false);
            }
        }

        private void ETDRFTNO_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string docNum = dt.GetValue("DocNum", 0).ToString().Trim();
            string docEntry = dt.GetValue("DocEntry", 0).ToString().Trim();
            //ETDONTRY ETDRFTNO
            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETDRFTNO").Specific;
            ETCD.Value = docNum;
            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETDONTRY").Specific;
            ETNM.Value = docEntry;

        }

        private void ETDRFTNO_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

                string styleCode = ((SAPbouiCOM.EditText)oForm.Items.Item("ETSTYLCD").Specific).Value.Trim();

                if (string.IsNullOrEmpty(styleCode))
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Please Select Style first.",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Error
                    );
                    BubbleEvent = false;
                    return;
                }

                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_DO")
                {
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(cflUID);

                    SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                    SAPbouiCOM.Condition oCon = oCons.Add();

                    oCon.Alias = "U_STYLECODE";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = styleCode;

                    oCFL.SetConditions(oCons);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error filtering Style Code: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }
        }



        private void ETSTYLCD_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string styleCode = dt.GetValue("U_STYLECODE", 0).ToString().Trim();
            string styleDesc = dt.GetValue("U_STYLENM", 0).ToString().Trim();
            string styleEntry = dt.GetValue("DocEntry", 0).ToString().Trim();
            string merCode = dt.GetValue("U_MARCHEN", 0).ToString().Trim();
            string merName = dt.GetValue("U_MARCHENN", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETSTYLCD").Specific;
            ETCD.Value = styleCode;

            SAPbouiCOM.EditText ETDS = (SAPbouiCOM.EditText)oForm.Items.Item("ETSTYLDS").Specific;
            ETDS.Value = styleDesc;

            SAPbouiCOM.EditText ETENTRY = (SAPbouiCOM.EditText)oForm.Items.Item("ETSTNTRY").Specific;
            ETENTRY.Value = styleEntry;

            SAPbouiCOM.EditText merCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETMERHN").Specific;
            merCD.Value = merCode;

            SAPbouiCOM.EditText merNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETMERCNM").Specific;
            merNM.Value = merName;
        }

        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = null;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                oForm.Freeze(true);

                int formWidth = oForm.ClientWidth;
                int formHeight = oForm.ClientHeight;

                int margin = 10;
                int gap = 10;

                // =========================
                // Top Grids: GRDSCLR + GRDSIZE
                // This part is working fine
                // =========================
                SAPbouiCOM.Item grdClr = oForm.Items.Item("GRDSCLR");
                SAPbouiCOM.Item grdSize = oForm.Items.Item("GRDSIZE");
                SAPbouiCOM.Item etStntry = oForm.Items.Item("ETSTNTRY");

                grdClr.Top = 31;
                grdClr.Left = etStntry.Left + etStntry.Width + 50;
                grdClr.Width = 178;

                grdSize.Top = 28;
                grdSize.Left = grdClr.Left + grdClr.Width + 34;
                grdSize.Width = formWidth - grdSize.Left - 220;

                int maxTopGridHeight = 130;
                int tabTopLimit = 170;

                grdClr.Height = Math.Min(maxTopGridHeight, tabTopLimit - grdClr.Top - gap);
                grdSize.Height = Math.Min(maxTopGridHeight, tabTopLimit - grdSize.Top - gap);


                // =========================
                // Tab Container: Item_8
                // Important:
                // Do NOT change Top/Height too much.
                // Otherwise folder headers become unclickable.
                // =========================
                SAPbouiCOM.Item tab = oForm.Items.Item("Item_8");

                tab.Left = margin;
                tab.Top = 170;
                tab.Width = formWidth - (margin * 2);

                // Keep enough bottom space for Add/Cancel button
                int bottomButtonSpace = 55;

                // Only update tab height safely
                tab.Height = Math.Max(190, formHeight - tab.Top - bottomButtonSpace);


                // =========================
                // Common Tab Content Area
                // Folder header needs free space
                // =========================
                int folderHeaderHeight = 55;   // more safe area for folder header
                int insideMargin = 10;         // gap after folder header

                int contentTop = tab.Top + folderHeaderHeight + insideMargin;
                int contentBottom = tab.Top + tab.Height - 20;
                int contentHeight = Math.Max(80, contentBottom - contentTop);


                // =========================
                // Pane 1: FOLMERCON
                // Matrix: MTXMRCON
                // =========================
                SAPbouiCOM.Item mtxMrCon = oForm.Items.Item("MTXMRCON");

                mtxMrCon.Top = contentTop;
                mtxMrCon.Left = tab.Left + 18;
                mtxMrCon.Width = tab.Width - 36;
                mtxMrCon.Height = contentHeight;


                // =========================
                // Pane 2: FOLCANCON
                // Matrix: MTXCDCLR
                // StaticText: STCDCLR
                // EditText: ETCDCLR
                // Grid: GRDCDCON
                // Buttons: BTNLDCAD, BTNSAVE
                // =========================
                SAPbouiCOM.Item mtxCdClr = oForm.Items.Item("MTXCDCLR");
                SAPbouiCOM.Item stCdClr = oForm.Items.Item("STCDCLR");
                SAPbouiCOM.Item etCdClr = oForm.Items.Item("ETCDCLR");
                SAPbouiCOM.Item grdCdCon = oForm.Items.Item("GRDCDCON");
                SAPbouiCOM.Item btnLoadCad = oForm.Items.Item("BTNLDCAD");
                SAPbouiCOM.Item btnSave = oForm.Items.Item("BTNSAVE");

                int labelTop = contentTop;
                int gridTop = labelTop + 30;

                int leftMatrixWidth = 220;
                int buttonHeightSpace = 25;

                mtxCdClr.Top = gridTop;
                mtxCdClr.Left = tab.Left + 10;
                mtxCdClr.Width = leftMatrixWidth;
                mtxCdClr.Height = Math.Max(70, contentBottom - gridTop - buttonHeightSpace);

                stCdClr.Top = labelTop;
                stCdClr.Left = mtxCdClr.Left + mtxCdClr.Width + gap;
                stCdClr.Width = 90;

                etCdClr.Top = labelTop;
                etCdClr.Left = stCdClr.Left + stCdClr.Width + 5;
                etCdClr.Width = 100;

                grdCdCon.Top = gridTop;
                grdCdCon.Left = mtxCdClr.Left + mtxCdClr.Width + gap;
                grdCdCon.Width = tab.Left + tab.Width - grdCdCon.Left - 20;
                grdCdCon.Height = Math.Max(70, contentBottom - gridTop - buttonHeightSpace);

                btnLoadCad.Top = contentBottom - 18;
                btnLoadCad.Left = grdCdCon.Left;
                btnLoadCad.Width = 65;

                btnSave.Top = btnLoadCad.Top;
                btnSave.Left = btnLoadCad.Left + btnLoadCad.Width + 5;
                btnSave.Width = 65;


                // =========================
                // Pane 3: FOLTEMP
                // Matrix: MTXCDCON
                // =========================
                SAPbouiCOM.Item mtxCdCon = oForm.Items.Item("MTXCDCON");

                mtxCdCon.Top = contentTop;
                mtxCdCon.Left = tab.Left + 18;
                mtxCdCon.Width = tab.Width - 36;
                mtxCdCon.Height = contentHeight;


                // =========================
                // Bottom Add / Cancel Buttons
                // =========================
                SAPbouiCOM.Item btnAdd = oForm.Items.Item("1");
                SAPbouiCOM.Item btnCancel = oForm.Items.Item("2");

                btnAdd.Top = formHeight - 35;
                btnCancel.Top = formHeight - 35;

                btnAdd.Left = 13;
                btnCancel.Left = btnAdd.Left + btnAdd.Width + 3;
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Resize error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
            finally
            {
                if (oForm != null)
                {
                    oForm.Freeze(false);
                }
            }
        }
    }
}
