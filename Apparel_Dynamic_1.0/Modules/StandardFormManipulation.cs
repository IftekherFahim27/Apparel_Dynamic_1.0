using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;
using Apparel_Dynamic_1._0.Helper;
using Apparel_Dynamic_1._0.Resources;

namespace Apparel_Dynamic_1._0.Modules
{
    class StandardFormManipulation
    {
        public StandardFormManipulation()
        {
            Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
        }

        private string lastFocusedItemUID = "";
        private string lastFocusedColUID = "";

        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.FormUID == "FIL_FRM_ROUTEMSTR")
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS ||
                    pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                {
                    lastFocusedItemUID = pVal.ItemUID;
                    lastFocusedColUID = pVal.ColUID;
                }
            }
            else if (pVal.FormUID == "FIL_FRM_SZTPMSTR")
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS ||
                    pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                {
                    lastFocusedItemUID = pVal.ItemUID;
                    lastFocusedColUID = pVal.ColUID;
                }
            }
            


            if (pVal.FormTypeEx == "9999" && pVal.ItemUID == "1" && pVal.BeforeAction == true && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED ||
                pVal.FormTypeEx == "9999" && pVal.ItemUID == "7" && pVal.BeforeAction == true && pVal.EventType == SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
            {
                try
                {
                    SAPbouiCOM.Form oForm = null;
                    string openForm = "";
                    // Detect which of the three forms is open
                    foreach (string formUID in new[] { "FIL_FRM_ROUTEMSTR", "FIL_FRM_SZTPMSTR" })
                    {
                        try
                        {
                            oForm = Application.SBO_Application.Forms.Item(formUID);
                            openForm = formUID;
                            break; // stop once we found one
                        }
                        catch
                        {
                            // not open → keep checking
                        }
                    }

                    if (oForm == null) return; // no relevant form is open

                    // Branch logic based on which form is open
                    if (openForm == "FIL_FRM_ROUTEMSTR")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                        {
                            oForm.Freeze(true);

                            string routeFind = ((SAPbouiCOM.EditText)oForm.Items.Item("ETCODE").Specific).Value.Trim();
                            string route = "";
                            if (routeFind.Contains("*"))
                            {
                                // ── Get value from Form 9999 Matrix ──
                                SAPbouiCOM.Form frm9999 = Application.SBO_Application.Forms.Item(pVal.FormUID);
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)frm9999.Items.Item("7").Specific;

                                int rowSelected = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);

                                if (rowSelected > 0)
                                {
                                    route = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Code").Cells.Item(rowSelected).Specific).Value;
                                }

                                SAPbouiCOM.Matrix oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXSTAGE").Specific;
                                oMtx.Columns.Item("CLSTGCOD").Editable = true;
                               
                                if (string.IsNullOrEmpty(route))
                                    return;

                                if (IsRouteCodeUsed(route))
                                {
                                    oMtx.Columns.Item("CLSTGCOD").Editable = false;
                                }
                                else
                                {
                                    oMtx.Columns.Item("CLSTGCOD").Editable = true;
                                    SAPbouiCOM.DBDataSource DBDataSourceLine = oForm.DataSources.DBDataSources.Item("@FIL_MR_RSM1");
                                    oMtx.FlushToDataSource();
                                    int lastRow = oMtx.RowCount;
                                    bool lastRowHasData = !string.IsNullOrWhiteSpace(((SAPbouiCOM.EditText)oMtx.Columns.Item("CLSTGCOD").Cells.Item(lastRow).Specific).Value);
                                    if (pVal.Row == lastRow && lastRowHasData)
                                    {
                                        Global.GFunc.SetNewLine(oMtx, DBDataSourceLine, 1, "");
                                    }
                                }
                                oMtx.LoadFromDataSource();
                            }

                            
                            oForm.Freeze(false);
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
    }
}
