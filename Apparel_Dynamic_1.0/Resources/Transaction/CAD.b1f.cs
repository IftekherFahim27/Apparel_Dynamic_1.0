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
            //                   Static text
            this.STSTATUS = ((SAPbouiCOM.StaticText)(this.GetItem("STSTATUS").Specific));
            this.STSTYLCD = ((SAPbouiCOM.StaticText)(this.GetItem("STSTYLCD").Specific));
            this.STSTYLDS = ((SAPbouiCOM.StaticText)(this.GetItem("STSTYLDS").Specific));
            this.STMERCHN = ((SAPbouiCOM.StaticText)(this.GetItem("STMERHN").Specific));
            this.STDOCNUM = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCNUM").Specific));
            this.STDOCDAT = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCDAT").Specific));
            this.STDRFTNO = ((SAPbouiCOM.StaticText)(this.GetItem("STDRFTNO").Specific));
            this.STSLCLR = ((SAPbouiCOM.StaticText)(this.GetItem("STSLCLR").Specific));
            this.STCDCLR = ((SAPbouiCOM.StaticText)(this.GetItem("STCDCLR").Specific));
            //                   Edittext
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
            //                   Combo box
            this.CBDOCNUM = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSERIES").Specific));
            this.CBSTATUS = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSTATUS").Specific));
            //                   Folder
            this.FOLMERCON = ((SAPbouiCOM.Folder)(this.GetItem("FOLMERCON").Specific));
            this.FOLCANCON = ((SAPbouiCOM.Folder)(this.GetItem("FOLCANCN").Specific));
            this.FOLTEMP = ((SAPbouiCOM.Folder)(this.GetItem("FOLTEMP").Specific));
            //                   Matrix
            this.MTXMRCON = ((SAPbouiCOM.Matrix)(this.GetItem("MTXMRCON").Specific));
            this.MTXMRCON.LostFocusAfter += new SAPbouiCOM._IMatrixEvents_LostFocusAfterEventHandler(this.MTXMRCON_LostFocusAfter);
            this.MTXMRCON.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.MTXMRCON_ChooseFromListAfter);
            this.MTXMRCON.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.MTXMRCON_ChooseFromListBefore);
            this.MTXCDCLR = ((SAPbouiCOM.Matrix)(this.GetItem("MTXCDCLR").Specific));
            this.MTXCDCLR.DoubleClickAfter += new SAPbouiCOM._IMatrixEvents_DoubleClickAfterEventHandler(this.MTXCDCLR_DoubleClickAfter);
            this.MTXCDCON = ((SAPbouiCOM.Matrix)(this.GetItem("MTXCDCON").Specific));
            //                   Grid
            this.GRDCDCON = ((SAPbouiCOM.Grid)(this.GetItem("GRDCDCON").Specific));
            this.GRDCDCON.LostFocusAfter += new SAPbouiCOM._IGridEvents_LostFocusAfterEventHandler(this.GRDCDCON_LostFocusAfter);
            this.GRDCDCON.ChooseFromListBefore += new SAPbouiCOM._IGridEvents_ChooseFromListBeforeEventHandler(this.GRDCDCON_ChooseFromListBefore);
            this.GRDCDCON.ChooseFromListAfter += new SAPbouiCOM._IGridEvents_ChooseFromListAfterEventHandler(this.GRDCDCON_ChooseFromListAfter);
            this.GRDSIZE = ((SAPbouiCOM.Grid)(this.GetItem("GRDSIZE").Specific));
            //                   Button
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.ADDButton.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.ADDButton_PressedAfter);
            this.ADDButton.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.ADDButton_PressedBefore);
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.BTNFETCH = ((SAPbouiCOM.Button)(this.GetItem("BTNFETCH").Specific));
            this.BTNFETCH.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.BTNFETCH_PressedAfter);
            this.BTNLDCAD = ((SAPbouiCOM.Button)(this.GetItem("BTNLDCAD").Specific));
            this.BTNLDCAD.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.BTNLDCAD_PressedAfter);
            this.BTNSAVE = ((SAPbouiCOM.Button)(this.GetItem("BTNSAVE").Specific));
            this.BTNSAVE.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.BTNSAVE_PressedAfter);
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
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);
            this.DataLoadAfter += new DataLoadAfterHandler(this.Form_DataLoadAfter);

        }


        private void OnCustomInitialize()
        {

        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            LoadHeaderColourSizeInfo(pVal.FormUID);
            LoadCADConsumptionDetails(pVal.FormUID);
        }


        private void LoadCADConsumptionDetails(string formUID)
        {
            SAPbouiCOM.Form oForm = null;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(formUID);
                oForm.Freeze(true);

                string docEntry =
                    ((SAPbouiCOM.EditText)oForm.Items.Item("ETDOCTRY").Specific).Value.Trim();

                if (string.IsNullOrEmpty(docEntry))
                    return;

                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("GRDCDCON").Specific;
                SAPbouiCOM.DataTable dtCAD = oForm.DataSources.DataTables.Item("DT_CADCN");

                string query = $@"
                                SELECT
                                    T0.""U_USETYPE""     AS ""Use Type"",
                                    T0.""U_ITEMCODE""    AS ""Item Code"",
                                    T0.""U_ITEMNAME""    AS ""Item Desc"",
                                    T0.""U_UOM""         AS ""UOM"",
                                    T0.""U_SHRNLPRC""    AS ""Shrinkage"",
                                    T0.""U_CBLWDIN""     AS ""Cut WD Inch"",
                                    T0.""U_CBLWDCM""     AS ""Cut WD CM"",
                                    T0.""U_POSITION""    AS ""Position"",
                                    T0.""U_ACTCONSM""    AS ""Consumption"",
                                    T0.""U_CUTWSTG""     AS ""Wastage"",
                                    T0.""U_TTLCONSM""    AS ""Total Consumption"",
                                    T0.""U_REMARKS""     AS ""Remarks"",
                                    T0.""U_BASECLR""    AS ""BSCLR""

                                FROM ""@FIL_DR_CADFABCN"" T0
                                WHERE T0.""DocEntry"" = '{docEntry}'
                                ORDER BY
                                    T0.""LineId""";

                dtCAD.ExecuteQuery(query);
                oGrid.DataTable = dtCAD;
                DisableAllGridColumns(oGrid);
                oGrid.AutoResizeColumns();

                // Set selected colour text
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETCDCLR").Specific).Value = "ALL Colour";
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "CAD consumption load error: " + ex.Message,
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

        private void DisableAllGridColumns(SAPbouiCOM.Grid oGrid)
        {
            for (int i = 0; i < oGrid.Columns.Count; i++)
            {
                oGrid.Columns.Item(i).Editable = false;
            }
        }

        private void LoadHeaderColourSizeInfo(string formUID)
        {
            SAPbouiCOM.Form oForm = null;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(formUID);
                oForm.Freeze(true);

                string docEntry =
                    ((SAPbouiCOM.EditText)oForm.Items.Item("ETDONTRY").Specific).Value.Trim();

                if (string.IsNullOrEmpty(docEntry))
                    return;

                // Set selected colour text
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETSLCLR").Specific).Value = "ALL Colour";

                SAPbouiCOM.Grid grdColour = (SAPbouiCOM.Grid)oForm.Items.Item("GRDSCLR").Specific;
                SAPbouiCOM.Grid grdSize = (SAPbouiCOM.Grid)oForm.Items.Item("GRDSIZE").Specific;
                SAPbouiCOM.DataTable dtColour = oForm.DataSources.DataTables.Item("DT_CLR");
                SAPbouiCOM.DataTable dtSize =oForm.DataSources.DataTables.Item("DT_SIZE");

                string colourQuery = $@"
                                        SELECT DISTINCT
                                            T0.""U_FGCOLOUR"" AS ""Colour Code"",
                                            T0.""U_FGCOLRNM"" AS ""Colour Name""
                                        FROM ""QUT1"" T0
                                        WHERE T0.""DocEntry"" = '{docEntry}'
                                          AND IFNULL(T0.""U_FGCOLOUR"", '') <> ''
                                        ORDER BY
                                            T0.""U_FGCOLOUR""";

                string sizeQuery = $@"
                                        SELECT
                                            T0.""U_FGSIZE""        AS ""Size Code"",
                                            T1.""Name""            AS ""Size Name"",
                                            SUM(T0.""Quantity"")   AS ""Quantity""
                                        FROM ""QUT1"" T0
                                        LEFT JOIN ""@FIL_MH_SIZEMSTR"" T1
                                            ON T0.""U_FGSIZE"" = T1.""Code""
                                        WHERE T0.""DocEntry"" = '{docEntry}'
                                          AND IFNULL(T0.""U_FGSIZE"", '') <> ''
                                        GROUP BY
                                            T0.""U_FGSIZE"",
                                            T1.""Name""
                                        ORDER BY
                                            T0.""U_FGSIZE""";

                dtColour.ExecuteQuery(colourQuery);
                grdColour.DataTable = dtColour;

                dtSize.ExecuteQuery(sizeQuery);
                grdSize.DataTable = dtSize;

                grdColour.AutoResizeColumns();
                grdSize.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Header colour/size load error: " + ex.Message,
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

        private void BTNSAVE_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = null;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                oForm.Freeze(true);

                SAPbouiCOM.DataTable dtCAD =
                    oForm.DataSources.DataTables.Item("DT_CADCN");

                SAPbouiCOM.Matrix oMatrix =
                    (SAPbouiCOM.Matrix)oForm.Items.Item("MTXCDCON").Specific;

                string currentBSCLR =
                    ((SAPbouiCOM.EditText)oForm.Items.Item("ETCDCLR").Specific).Value.Trim();

                if (string.IsNullOrEmpty(currentBSCLR))
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Please select CAD colour first.",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning
                    );
                    return;
                }

                // Remove empty matrix rows first
                for (int i = oMatrix.RowCount; i >= 1; i--)
                {
                    string itemCode =
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLITMCOD")
                            .Cells.Item(i).Specific).Value.Trim();

                    if (string.IsNullOrEmpty(itemCode))
                    {
                        oMatrix.DeleteRow(i);
                    }
                }

                // Remove only previous rows of selected BSCLR
                for (int i = oMatrix.RowCount; i >= 1; i--)
                {
                    string matrixBSCLR =
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLBSCLR")
                            .Cells.Item(i).Specific).Value.Trim();

                    if (matrixBSCLR == currentBSCLR)
                    {
                        oMatrix.DeleteRow(i);
                    }
                }

                // Add grid rows again
                for (int i = 0; i < dtCAD.Rows.Count; i++)
                {
                    string itemCode =
                        dtCAD.GetValue("ItemCode", i).ToString().Trim();

                    if (string.IsNullOrEmpty(itemCode))
                        continue;

                    oMatrix.AddRow(1);

                    int row = oMatrix.RowCount;

                    SetMatrixValue(oMatrix, "#", row, row.ToString());

                    SetMatrixValue(oMatrix, "CLUSETYP", row, dtCAD.GetValue("UseType", i).ToString());
                    SetMatrixValue(oMatrix, "CLPOSTN", row, dtCAD.GetValue("Position", i).ToString());
                    SetMatrixValue(oMatrix, "CLITMCOD", row, dtCAD.GetValue("ItemCode", i).ToString());
                    SetMatrixValue(oMatrix, "CLITMNAM", row, dtCAD.GetValue("Item Desc", i).ToString());
                    SetMatrixValue(oMatrix, "CLUOM", row, dtCAD.GetValue("UoM", i).ToString());

                    SetMatrixValue(oMatrix, "CLCNSMPN", row, dtCAD.GetValue("Consumption", i).ToString());
                    SetMatrixValue(oMatrix, "CLSHRKPR", row, dtCAD.GetValue("Shrinkage", i).ToString());
                    SetMatrixValue(oMatrix, "CLWSTGP", row, dtCAD.GetValue("Wastage", i).ToString());

                    SetMatrixValue(oMatrix, "CLCTWDIN", row, dtCAD.GetValue("Cut WD Inch", i).ToString());
                    SetMatrixValue(oMatrix, "CLCTWDCM", row, dtCAD.GetValue("Cut WD CM", i).ToString());

                    SetMatrixValue(oMatrix, "CLTTLCON", row, dtCAD.GetValue("Total Consumption", i).ToString());
                    SetMatrixValue(oMatrix, "CLRMRKS", row, dtCAD.GetValue("Remarks", i).ToString());

                    SetMatrixValue(oMatrix, "CLBSCLR", row, currentBSCLR);
                }

                // Re-arrange LineId / # after all delete and add
                ReArrangeMatrixLineId(oMatrix);

                oMatrix.FlushToDataSource();

                Application.SBO_Application.StatusBar.SetText(
                    "CAD consumption rows updated successfully.",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success
                );
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "CAD save error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
            finally
            {
                try
                {
                    if (oForm != null)
                        oForm.Freeze(false);
                }
                catch { }
            }
        }

        private void SetMatrixValue(SAPbouiCOM.Matrix matrix,string colUID,int row,string value)
        {
            ((SAPbouiCOM.EditText)matrix.Columns.Item(colUID)
                .Cells.Item(row).Specific).Value = value;
        }

        private void ReArrangeMatrixLineId(SAPbouiCOM.Matrix matrix)
        {
            for (int i = 1; i <= matrix.RowCount; i++)
            {
                SetMatrixValue(matrix, "#", i, i.ToString());
            }
        }

        private void GRDCDCON_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = null;

            try
            {
                if (pVal.Row < 0)
                    return;

                if (pVal.ColUID != "Cut WD Inch" &&
                    pVal.ColUID != "Cut WD CM" &&
                    pVal.ColUID != "Consumption" &&
                    pVal.ColUID != "Shrinkage" &&
                    pVal.ColUID != "Wastage")
                    return;

                oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                oForm.Freeze(true);

                SAPbouiCOM.Grid oGrid =
                    (SAPbouiCOM.Grid)oForm.Items.Item("GRDCDCON").Specific;

                SAPbouiCOM.DataTable dtCAD =
                    oForm.DataSources.DataTables.Item("DT_CADCN");

                int dtRow = oGrid.GetDataTableRowIndex(pVal.Row);

                if (dtRow < 0)
                    return;

                // =========================
                // Negative value protection
                // =========================
                double enteredValue;

                if (double.TryParse(
                    dtCAD.GetValue(pVal.ColUID, dtRow).ToString().Trim(),
                    out enteredValue))
                {
                    if (enteredValue < 0)
                    {
                        dtCAD.SetValue(pVal.ColUID, dtRow, "0");
                    }
                }

                // =========================
                // Inch / CM conversion
                // =========================
                if (pVal.ColUID == "Cut WD Inch")
                {
                    double inch;

                    if (double.TryParse(
                        dtCAD.GetValue("Cut WD Inch", dtRow).ToString().Trim(),
                        out inch))
                    {
                        if (inch < 0)
                        {
                            inch = 0;
                            dtCAD.SetValue("Cut WD Inch", dtRow, "0");
                        }

                        dtCAD.SetValue(
                            "Cut WD CM",
                            dtRow,
                            (inch * 2.54).ToString("0.##")
                        );
                    }
                    else
                    {
                        dtCAD.SetValue("Cut WD CM", dtRow, "");
                    }
                }
                else if (pVal.ColUID == "Cut WD CM")
                {
                    double cm;

                    if (double.TryParse(
                        dtCAD.GetValue("Cut WD CM", dtRow).ToString().Trim(),
                        out cm))
                    {
                        if (cm < 0)
                        {
                            cm = 0;
                            dtCAD.SetValue("Cut WD CM", dtRow, "0");
                        }

                        dtCAD.SetValue(
                            "Cut WD Inch",
                            dtRow,
                            (cm / 2.54).ToString("0.##")
                        );
                    }
                    else
                    {
                        dtCAD.SetValue("Cut WD Inch", dtRow, "");
                    }
                }

                // =========================
                // Total Consumption
                // =========================
                double consumption = 0;
                double shrinkage = 0;
                double wastage = 0;

                double.TryParse(
                    dtCAD.GetValue("Consumption", dtRow).ToString().Trim(),
                    out consumption);

                double.TryParse(
                    dtCAD.GetValue("Shrinkage", dtRow).ToString().Trim(),
                    out shrinkage);

                double.TryParse(
                    dtCAD.GetValue("Wastage", dtRow).ToString().Trim(),
                    out wastage);

                if (consumption < 0)
                {
                    consumption = 0;
                    dtCAD.SetValue("Consumption", dtRow, "0");
                }

                if (shrinkage < 0)
                {
                    shrinkage = 0;
                    dtCAD.SetValue("Shrinkage", dtRow, "0");
                }

                if (wastage < 0)
                {
                    wastage = 0;
                    dtCAD.SetValue("Wastage", dtRow, "0");
                }

                double totalConsumption =
                    consumption +
                    (consumption * shrinkage) +
                    (consumption * wastage);

                dtCAD.SetValue(
                    "Total Consumption",
                    dtRow,
                    totalConsumption.ToString("0.######")
                );
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Grid calculation error: " + ex.Message,
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


        private void GRDCDCON_ChooseFromListBefore(object sboObject,SAPbouiCOM.SBOItemEventArg pVal,out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                SAPbouiCOM.Form oForm =
                    Application.SBO_Application.Forms.Item(pVal.FormUID);

                string cflUID = "";

                if (pVal.ColUID == "UseType")
                {
                    cflUID = "CFL_USE2";
                }
                else if (pVal.ColUID == "Position")
                {
                    cflUID = "CFL_POS2";
                }
                else if (pVal.ColUID == "ItemCode")
                {
                    cflUID = "CFL_ITM2";
                }
                else
                {
                    return;
                }

                SAPbouiCOM.ChooseFromList oCFL =
                    oForm.ChooseFromLists.Item(cflUID);

                SAPbouiCOM.Conditions oCons =
                    new SAPbouiCOM.Conditions();

                SAPbouiCOM.Condition oCon1 = oCons.Add();

                if (pVal.ColUID == "ItemCode")
                {
                    oCon1.Alias = "InvntItem";
                    oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon1.CondVal = "Y";
                }
                else
                {
                    oCon1.Alias = "U_ACTIVE";
                    oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon1.CondVal = "Y";
                }

                oCFL.SetConditions(oCons);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Grid CFL condition error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );

                BubbleEvent = false;
            }
        }

        private void GRDCDCON_ChooseFromListAfter(object sboObject,SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg =
                    (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;

                SAPbouiCOM.DataTable selectedDt = cflArg.SelectedObjects;

                if (selectedDt == null || selectedDt.Rows.Count == 0)
                    return;

                SAPbouiCOM.Form oForm =
                    Application.SBO_Application.Forms.Item(pVal.FormUID);

                SAPbouiCOM.Grid oGrid =
                    (SAPbouiCOM.Grid)oForm.Items.Item("GRDCDCON").Specific;

                SAPbouiCOM.DataTable dtCAD =
                    oForm.DataSources.DataTables.Item("DT_CADCN");

                int dtRow = oGrid.GetDataTableRowIndex(pVal.Row);

                if (dtRow < 0)
                    return;

                if (pVal.ColUID == "UseType")
                {
                    string code = selectedDt.GetValue("Code", 0).ToString().Trim();
                    dtCAD.SetValue("UseType", dtRow, code);
                }
                else if (pVal.ColUID == "Position")
                {
                    string code = selectedDt.GetValue("Code", 0).ToString().Trim();
                    dtCAD.SetValue("Position", dtRow, code);
                }
                else if (pVal.ColUID == "ItemCode")
                {
                    string itemCode = selectedDt.GetValue("ItemCode", 0).ToString().Trim();
                    string itemName = selectedDt.GetValue("ItemName", 0).ToString().Trim();

                    dtCAD.SetValue("ItemCode", dtRow, itemCode);
                    dtCAD.SetValue("Item Desc", dtRow, itemName);

                    // Add new row only if current row is last row
                    if (dtRow == dtCAD.Rows.Count - 1)
                    {
                        dtCAD.Rows.Add(1);

                        int newRow = dtCAD.Rows.Count - 1;

                        dtCAD.SetValue("UseType", newRow, "");
                        dtCAD.SetValue("Position", newRow, "");
                        dtCAD.SetValue("ItemCode", newRow, "");
                        dtCAD.SetValue("Item Desc", newRow, "");
                        dtCAD.SetValue("UoM", newRow, "");
                        dtCAD.SetValue("Consumption", newRow, 0);
                        dtCAD.SetValue("Shrinkage", newRow, 0);
                        dtCAD.SetValue("Wastage", newRow, 0);
                        dtCAD.SetValue("Cut WD Inch", newRow, 0);
                        dtCAD.SetValue("Cut WD CM", newRow, 0);
                        dtCAD.SetValue("Total Consumption", newRow, 0);
                        dtCAD.SetValue("Remarks", newRow, "");

                        // Keep same colour
                        dtCAD.SetValue(
                            "BSCLR",
                            newRow,
                            dtCAD.GetValue("BSCLR", dtRow).ToString()
                        );
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Grid CFL After error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
        }

        private void BTNLDCAD_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = null;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                oForm.Freeze(true);

                SAPbouiCOM.Grid oGrid =
                    (SAPbouiCOM.Grid)oForm.Items.Item("GRDCDCON").Specific;

                SAPbouiCOM.DataTable dtCAD =
                    oForm.DataSources.DataTables.Item("DT_CADCN");

                string docEntry =
                    ((SAPbouiCOM.EditText)oForm.Items.Item("ETDOCTRY").Specific).Value.Trim();

                string bsClr =
                    ((SAPbouiCOM.EditText)oForm.Items.Item("ETCDCLR").Specific).Value.Trim();

                if (string.IsNullOrEmpty(docEntry))
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Please save or select document first.",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning
                    );
                    return;
                }

                if (string.IsNullOrEmpty(bsClr))
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Please select CAD colour first.",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning
                    );
                    return;
                }

                bsClr = bsClr.Replace("'", "''");

                string query = $@"
                                SELECT
                                    ""U_USETYPE""     AS ""UseType"",
                                    ""U_POSITION""    AS ""Position"",
                                    ""U_ITEMCODE""    AS ""ItemCode"",
                                    ""U_ITEMNAME""    AS ""Item Desc"",
                                    ""U_UOM""         AS ""UoM"",
                                    0                 AS ""Consumption"",
                                    ""U_SHRNLPRC""    AS ""Shrinkage"",
                                    0                 AS ""Wastage"",
                                    ""U_CBLWDIN""     AS ""Cut WD Inch"",
                                    ""U_CBLWDCM""     AS ""Cut WD CM"",
                                    0                 AS ""Total Consumption"",
                                    ''                AS ""Remarks"",
                                    '{bsClr}'         AS ""BSCLR""
                                FROM ""@FIL_DR_CADMFAB""
                                WHERE ""DocEntry"" = '{docEntry}'
                                  AND IFNULL(""U_CLRAPPL"", 'N') = 'Y'
                                ORDER BY ""LineId""
                            ";

                dtCAD.ExecuteQuery(query);

                // Add one empty row at last
                dtCAD.Rows.Add(1);

                int lastRow = dtCAD.Rows.Count - 1;

                dtCAD.SetValue("UseType", lastRow, "");
                dtCAD.SetValue("Position", lastRow, "");
                dtCAD.SetValue("ItemCode", lastRow, "");
                dtCAD.SetValue("Item Desc", lastRow, "");
                dtCAD.SetValue("UoM", lastRow, "");
                dtCAD.SetValue("Consumption", lastRow, 0);
                dtCAD.SetValue("Shrinkage", lastRow, 0);
                dtCAD.SetValue("Wastage", lastRow, 0);
                dtCAD.SetValue("Cut WD Inch", lastRow, 0);
                dtCAD.SetValue("Cut WD CM", lastRow, 0);
                dtCAD.SetValue("Total Consumption", lastRow, 0);
                dtCAD.SetValue("Remarks", lastRow, "");
                dtCAD.SetValue("BSCLR", lastRow, bsClr);

                oGrid.DataTable = dtCAD;

                // optional
                oGrid.AutoResizeColumns();

                Application.SBO_Application.StatusBar.SetText(
                    "CAD data loaded successfully.",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success
                );
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Load CAD Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
            finally
            {
                try
                {
                    if (oForm != null)
                        oForm.Freeze(false);
                }
                catch { }
            }
        }

        private void ADDButton_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            //Series Initialization
            SAPbouiCOM.DBDataSource oDBH = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@FIL_DH_CADFABCN");   //DEFINE  DATASOURCES.
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                SAPbouiCOM.ComboBox ocmb = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBSERIES").Specific;
                Global.GFunc.LoadComboBoxSeries(ocmb, "FIL_D_CADFABCN");  //Object Type
                string ocmbvalue = ocmb.Selected.Value;
                long docno = oForm.BusinessObject.GetNextSerialNumber(ocmbvalue, "FIL_D_CADFABCN");

                oDBH.SetValue("DocNum", 0, docno.ToString()); // only set the value in string.

                //Date
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETDOCDAT").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
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
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                string styleCode = oForm.DataSources.DBDataSources.Item("@FIL_DH_CADFABCN").GetValue("U_STYLECODE", 0).Trim();
                string draftOrder = oForm.DataSources.DBDataSources.Item("@FIL_DH_CADFABCN").GetValue("U_SQNO", 0).Trim();
                string docNum = oForm.DataSources.DBDataSources.Item("@FIL_DH_CADFABCN").GetValue("DocNum", 0).Trim();
                string merc = oForm.DataSources.DBDataSources.Item("@FIL_DH_CADFABCN").GetValue("U_MRCHDNID", 0).Trim();


                if (styleCode == "")
                {
                    Global.GFunc.ShowError("Enter Style Code");
                    oForm.ActiveItem = "ETSTYLCD";
                    return BubbleEvent = false;
                }
                if (draftOrder == "")
                {
                    Global.GFunc.ShowError("Enter Draft Order ");
                    oForm.ActiveItem = "ETDRFTNO";
                    return BubbleEvent = false;
                }
                if (docNum == "")
                {
                    Global.GFunc.ShowError("Enter DocNum ");
                    return BubbleEvent = false;
                }
                if (merc == "")
                {
                    Global.GFunc.ShowError("Enter Merchandiser ");
                    oForm.ActiveItem = "ETDRFTNO";
                    return BubbleEvent = false;
                }


                PreventEmptyLastRow(oForm, "@FIL_DR_CADMFAB", MTXMRCON, "U_ITEMCODE");
                PreventEmptyLastRow(oForm, "@FIL_DR_CADFABCL", MTXCDCLR, "U_COLORCODE");
                PreventEmptyLastRow(oForm, "@FIL_DR_CADFABCN", MTXCDCON, "U_ITEMCODE");

            }

            return BubbleEvent;
        }


        private void MTXMRCON_LostFocusAfter(object sboObject,SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.Row <= 0)
                    return;

                if (pVal.ColUID != "CLCTWDIN" &&
                    pVal.ColUID != "CLCTWDCM")
                    return;

                SAPbouiCOM.Form oForm =Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Matrix oMatrix =(SAPbouiCOM.Matrix)oForm.Items.Item("MTXMRCON").Specific;
                SAPbouiCOM.EditText txtInch =(SAPbouiCOM.EditText)oMatrix.Columns.Item("CLCTWDIN").Cells.Item(pVal.Row).Specific;
                SAPbouiCOM.EditText txtCM =(SAPbouiCOM.EditText)oMatrix.Columns.Item("CLCTWDCM").Cells.Item(pVal.Row).Specific;

                // Inch → CM
                if (pVal.ColUID == "CLCTWDIN")
                {
                    double inch;

                    if (double.TryParse(txtInch.Value.Trim(), out inch))
                    {
                        if (inch < 0)
                        {
                            txtInch.Value = "0";
                            txtCM.Value = "0";
                            return;
                        }

                        txtCM.Value = (inch * 2.54).ToString("0.##");
                    }
                    else
                    {
                        txtCM.Value = "";
                    }
                }

                // CM → Inch
                else if (pVal.ColUID == "CLCTWDCM")
                {
                    double cm;

                    if (double.TryParse(txtCM.Value.Trim(), out cm))
                    {
                        if (cm < 0)
                        {
                            txtCM.Value = "0";
                            txtInch.Value = "0";
                            return;
                        }

                        txtInch.Value = (cm / 2.54).ToString("0.##");
                    }
                    else
                    {
                        txtInch.Value = "";
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Width conversion error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
        }


        private void MTXCDCLR_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = null;

            try
            {
                if (pVal.Row <= 0)
                    return;

                oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix =
                    (SAPbouiCOM.Matrix)oForm.Items.Item("MTXCDCLR").Specific;

                string colourCode =
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLCLRCOD")
                        .Cells.Item(pVal.Row).Specific).Value.Trim();

                SAPbouiCOM.EditText ETCDCLR =
                    (SAPbouiCOM.EditText)oForm.Items.Item("ETCDCLR").Specific;

                ETCDCLR.Value = colourCode;

                LoadCADConsumptionByBaseColour(oForm, colourCode);
                SetCADConsumptionGridEditableForColour(oForm);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "MTXCDCLR Double Click Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
            finally
            {
                try
                {
                    if (oForm != null)
                        oForm.Freeze(false);
                }
                catch { }
            }
        }

        private void SetCADConsumptionGridEditableForColour(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Grid oGrid =
                (SAPbouiCOM.Grid)oForm.Items.Item("GRDCDCON").Specific;

            // First enable all columns
            for (int i = 0; i < oGrid.Columns.Count; i++)
            {
                oGrid.Columns.Item(i).Editable = true;
            }

            // Then disable only these columns
            SetGridColumnEditable(oGrid, "BSCLR", false);
            SetGridColumnEditable(oGrid, "Total Consumption", false);
            SetGridColumnEditable(oGrid, "Item Desc", false);
        }

        private void SetGridColumnEditable(SAPbouiCOM.Grid oGrid, string colUID, bool editable)
        {
            try
            {
                oGrid.Columns.Item(colUID).Editable = editable;
            }
            catch
            {
                // Column not found, ignore
            }
        }

        private void LoadCADConsumptionByBaseColour(SAPbouiCOM.Form oForm, string colourCode)
        {
            try
            {
                SAPbouiCOM.Grid oGrid =
                    (SAPbouiCOM.Grid)oForm.Items.Item("GRDCDCON").Specific;

                SAPbouiCOM.DataTable dtCAD =
                    oForm.DataSources.DataTables.Item("DT_CADCN");

                string docEntry =
                    ((SAPbouiCOM.EditText)oForm.Items.Item("ETDOCTRY").Specific).Value.Trim();

                SAPbouiCOM.Button btnLoadCAD =
                    (SAPbouiCOM.Button)oForm.Items.Item("BTNLDCAD").Specific;

                if (string.IsNullOrEmpty(docEntry) || string.IsNullOrEmpty(colourCode))
                {
                    //btnLoadCAD.Item.Enabled = true;
                    return;
                }

                string safeColourCode = colourCode.Replace("'", "''");

                string query = $@"
                                SELECT
                                    ""U_USETYPE""   AS ""UseType"",
                                    ""U_POSITION""  AS ""Position"",
                                    ""U_ITEMCODE""  AS ""ItemCode"",
                                    ""U_ITEMNAME""  AS ""Item Desc"",
                                    ""U_UOM""       AS ""UoM"",
                                    ""U_ACTCONSM""  AS ""Consumption"",
                                    ""U_SHRNLPRC""  AS ""Shrinkage"",
                                    ""U_CUTWSTG""   AS ""Wastage"",
                                    ""U_CBLWDIN""   AS ""Cut WD Inch"",
                                    ""U_CBLWDCM""   AS ""Cut WD CM"",
                                    ""U_TTLCONSM""  AS ""Total Consumption"",
                                    ""U_REMARKS""   AS ""Remarks"",
                                    ""U_BASECLR""   AS ""BSCLR""
                                FROM ""@FIL_DR_CADFABCN""
                                WHERE ""DocEntry"" = '{docEntry}'
                                  AND ""U_BASECLR"" = '{safeColourCode}'
                                ORDER BY ""LineId""
                            ";

                dtCAD.ExecuteQuery(query);

                bool hasData = false;

                for (int i = 0; i < dtCAD.Rows.Count; i++)
                {
                    string itemCode =
                        dtCAD.GetValue("ItemCode", i).ToString().Trim();

                    if (!string.IsNullOrEmpty(itemCode))
                    {
                        hasData = true;
                        break;
                    }
                }

                if (hasData)
                {
                    dtCAD.Rows.Add(1);

                    int newRow = dtCAD.Rows.Count - 1;

                    dtCAD.SetValue("UseType", newRow, "");
                    dtCAD.SetValue("Position", newRow, "");
                    dtCAD.SetValue("ItemCode", newRow, "");
                    dtCAD.SetValue("Item Desc", newRow, "");
                    dtCAD.SetValue("UoM", newRow, "");
                    dtCAD.SetValue("Consumption", newRow, 0);
                    dtCAD.SetValue("Shrinkage", newRow, 0);
                    dtCAD.SetValue("Wastage", newRow, 0);
                    dtCAD.SetValue("Cut WD Inch", newRow, 0);
                    dtCAD.SetValue("Cut WD CM", newRow, 0);
                    dtCAD.SetValue("Total Consumption", newRow, 0);
                    dtCAD.SetValue("Remarks", newRow, "");
                    dtCAD.SetValue("BSCLR", newRow, colourCode);

                    btnLoadCAD.Item.Enabled = false;
                }
                else
                {
                    btnLoadCAD.Item.Enabled = true;
                }

                oGrid.DataTable = dtCAD;
                oGrid.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Load CAD colour consumption error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
        }

        private void MTXMRCON_ChooseFromListBefore(object sboObject,SAPbouiCOM.SBOItemEventArg pVal,out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

                string cflUID = "";

                if (pVal.ColUID == "CLUSETYP")
                {
                    cflUID = "CFL_USE";
                }
                else if (pVal.ColUID == "CLPOS")
                {
                    cflUID = "CFL_POS";
                }
                else if (pVal.ColUID == "CLITMCOD")
                {
                    cflUID = "CFL_ITEM";
                }
                else
                {
                    return;
                }

                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(cflUID);
                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition oCon1 = oCons.Add();

                if (pVal.ColUID == "CLITMCOD")
                {
                    oCon1.Alias = "InvntItem";
                    oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon1.CondVal = "Y";
                }
                else
                {
                    oCon1.Alias = "U_ACTIVE";
                    oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon1.CondVal = "Y";
                }

                oCFL.SetConditions(oCons);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "CFL condition error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );

                BubbleEvent = false;
            }
        }

        private void MTXMRCON_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg =
                    (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;

                SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
                if (dt == null || dt.Rows.Count == 0)
                    return;

                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Matrix oMatrix =
                    (SAPbouiCOM.Matrix)oForm.Items.Item("MTXMRCON").Specific;

                string selectedCode = "";

                if (pVal.ColUID == "CLUSETYP")
                {
                    // From CFL_USE table
                    selectedCode = dt.GetValue("Code", 0).ToString().Trim();

                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLUSETYP")
                        .Cells.Item(pVal.Row).Specific).Value = selectedCode;
                }
                else if (pVal.ColUID == "CLPOS")
                {
                    // From CFL_POS table
                    selectedCode = dt.GetValue("Code", 0).ToString().Trim();

                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLPOS")
                        .Cells.Item(pVal.Row).Specific).Value = selectedCode;
                }
                else if(pVal.ColUID == "CLITMCOD")
                {
                    string itemCode = dt.GetValue("ItemCode", 0).ToString().Trim();
                    string itemDesc = dt.GetValue("ItemName", 0).ToString().Trim();

                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLITMCOD")
                        .Cells.Item(pVal.Row).Specific).Value = itemCode;

                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLITMNAM")
                        .Cells.Item(pVal.Row).Specific).Value = itemDesc;

                    AddLineIfLastRowHasValue(oForm, "MTXMRCON", "@FIL_DR_CADMFAB", "U_ITEMCODE");
                }

            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "CFL After error: " + ex.Message,
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
                // Add Demo Colour Last Line
                // =========================
                dbClr.InsertRecord(matrixRow);
                dbClr.SetValue("LineId", matrixRow, (matrixRow + 1).ToString());
                dbClr.SetValue("U_COLORCODE", matrixRow, "NoColour");
                dbClr.SetValue("U_COLORNAME", matrixRow, "NoColour");
                matrixRow++;

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
            try
            {
                SAPbouiCOM.Form oForm =
                    Application.SBO_Application.Forms.Item(pVal.FormUID);

                SAPbouiCOM.ISBOChooseFromListEventArg cflArg =
                    (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;

                SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;

                if (dt == null || dt.Rows.Count == 0)
                    return;

                string styleCode = dt.GetValue("U_STYLECODE", 0).ToString().Trim();
                string styleDesc = dt.GetValue("U_STYLENM", 0).ToString().Trim();
                string styleEntry = dt.GetValue("DocEntry", 0).ToString().Trim();
                string merCode = dt.GetValue("U_MARCHEN", 0).ToString().Trim();
                string merName = dt.GetValue("U_MARCHENN", 0).ToString().Trim();

                ((SAPbouiCOM.EditText)oForm.Items.Item("ETSTYLCD").Specific).Value = styleCode;
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETSTYLDS").Specific).Value = styleDesc;
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETSTNTRY").Specific).Value = styleEntry;
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETMERHN").Specific).Value = merCode;
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETMERCNM").Specific).Value = merName;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                // Ignore SAP focus issue [66000-153]
                if (ex.Message.Contains("66000-153"))
                    return;

                Application.SBO_Application.StatusBar.SetText(
                    ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
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
                // Caption Width Fix
                // =========================
                oForm.Items.Item("STSLCLR").Width = 110;
                oForm.Items.Item("FOLMERCON").Width = 180;
                oForm.Items.Item("FOLCANCN").Width = 140;

                // =========================
                // Tab Container: Item_8
                // =========================
                SAPbouiCOM.Item tab = oForm.Items.Item("Item_8");

                tab.Left = margin;
                tab.Top = 170;
                tab.Width = formWidth - (margin * 2);

                int bottomButtonSpace = 55;
                tab.Height = Math.Max(190, formHeight - tab.Top - bottomButtonSpace);

                // =========================
                // Common Tab Content Area
                // =========================
                int folderHeaderHeight = 55;
                int insideMargin = 10;

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
                // Pane 2: FOLCANCN
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

                // =========================
                // Auto Resize All Columns
                // =========================
                AutoResizeAllMatrixAndGridColumns(oForm);
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

        private void AutoResizeAllMatrixAndGridColumns(SAPbouiCOM.Form oForm)
        {
            try
            {
                ((SAPbouiCOM.Grid)oForm.Items.Item("GRDSCLR").Specific).AutoResizeColumns();
                ((SAPbouiCOM.Grid)oForm.Items.Item("GRDSIZE").Specific).AutoResizeColumns();
                ((SAPbouiCOM.Grid)oForm.Items.Item("GRDCDCON").Specific).AutoResizeColumns();

                ((SAPbouiCOM.Matrix)oForm.Items.Item("MTXMRCON").Specific).AutoResizeColumns();
                ((SAPbouiCOM.Matrix)oForm.Items.Item("MTXCDCLR").Specific).AutoResizeColumns();
                ((SAPbouiCOM.Matrix)oForm.Items.Item("MTXCDCON").Specific).AutoResizeColumns();
            }
            catch
            {
            }
        }
    }
}
