using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apparel_Dynamic_1._0.Resources.Transaction
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Transaction.CAD", "Resources/Transaction/CAD.b1f")]
    class CAD : UserFormBase
    {
        public CAD()
        {
        }

        private SAPbouiCOM.StaticText STSTATUS, STSTYLCD, STSTYLDS, STMERCHN, STDOCNUM, STDOCDAT, STDRFTNO, STSLCLR, STCDCLR;
        private SAPbouiCOM.EditText ETSLCLR, ETSTYLCD, ETSTYLDS, ETMERCHN, ETDOCNUM, ETMERCNM, ETDONTRY, ETDOCDAT, ETDRFTNO, ETCDCLR, ETDOCTRY, ETSTNTRY;
        private SAPbouiCOM.ComboBox CBDOCNUM, CBSTATUS;
        private SAPbouiCOM.Folder FOLMERCON, FOLCANCON, FOLTEMP;
        private SAPbouiCOM.Matrix MTXMRCON, MTXCDCLR, MTXCDCON;
        private SAPbouiCOM.Grid GRDCDCON, GRDSIZE, GRDSCLR;
        private SAPbouiCOM.Button ADDButton, CancelButton, BTNFETCH, BTNLDCAD, BTNSAVE;


        public override void OnInitializeComponent()
        {
            //   Static text
            this.STSTATUS = ((SAPbouiCOM.StaticText)(this.GetItem("STSTATUS").Specific));
            this.STSTYLCD = ((SAPbouiCOM.StaticText)(this.GetItem("STSTYLCD").Specific));
            this.STSTYLDS = ((SAPbouiCOM.StaticText)(this.GetItem("STSTYLDS").Specific));
            this.STMERCHN = ((SAPbouiCOM.StaticText)(this.GetItem("STMERCHN").Specific));
            this.STDOCNUM = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCNUM").Specific));
            this.STDOCDAT = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCDAT").Specific));
            this.STDRFTNO = ((SAPbouiCOM.StaticText)(this.GetItem("STDRFTNO").Specific));
            this.STSLCLR = ((SAPbouiCOM.StaticText)(this.GetItem("STSLCLR").Specific));
            this.STCDCLR = ((SAPbouiCOM.StaticText)(this.GetItem("STCDCLR").Specific));
            //   Edittext
            this.ETSLCLR = ((SAPbouiCOM.EditText)(this.GetItem("ETSLCLR").Specific));
            this.ETSTYLCD = ((SAPbouiCOM.EditText)(this.GetItem("ETSTYLCD").Specific));
            this.ETSTYLDS = ((SAPbouiCOM.EditText)(this.GetItem("ETSTYLDS").Specific));
            this.ETMERCHN = ((SAPbouiCOM.EditText)(this.GetItem("ETMERCHN").Specific));
            this.ETDOCNUM = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCNUM").Specific));
            this.ETDOCDAT = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCDAT").Specific));
            this.ETDRFTNO = ((SAPbouiCOM.EditText)(this.GetItem("ETDRFTNO").Specific));
            this.ETCDCLR = ((SAPbouiCOM.EditText)(this.GetItem("ETCDCLR").Specific));
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            //   Combo box
            this.CBDOCNUM = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSERIES").Specific));
            this.CBSTATUS = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSTATUS").Specific));
            //   Folder
            this.FOLMERCON = ((SAPbouiCOM.Folder)(this.GetItem("FOLMERCON").Specific));
            this.FOLCANCON = ((SAPbouiCOM.Folder)(this.GetItem("FOLCANCN").Specific));
            this.FOLTEMP = ((SAPbouiCOM.Folder)(this.GetItem("FOLTEMP").Specific));
            //   Matrix
            this.MTXMRCON = ((SAPbouiCOM.Matrix)(this.GetItem("MTXMRCON").Specific));
            this.MTXCDCLR = ((SAPbouiCOM.Matrix)(this.GetItem("MTXCDCLR").Specific));
            this.MTXCDCON = ((SAPbouiCOM.Matrix)(this.GetItem("MTXCDCON").Specific));
            //   Grid
            this.GRDCDCON = ((SAPbouiCOM.Grid)(this.GetItem("GRDCDCON").Specific));
            this.GRDSIZE = ((SAPbouiCOM.Grid)(this.GetItem("GRDSIZE").Specific));
            //   Button
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.BTNFETCH = ((SAPbouiCOM.Button)(this.GetItem("BTNFETCH").Specific));
            this.BTNLDCAD = ((SAPbouiCOM.Button)(this.GetItem("BTNLDCAD").Specific));
            this.BTNSAVE = ((SAPbouiCOM.Button)(this.GetItem("BTNSAVE").Specific));
            this.GRDSCLR = ((SAPbouiCOM.Grid)(this.GetItem("GRDSCLR").Specific));
            this.ETSTNTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETSTNTRY").Specific));
            this.ETMERCNM = ((SAPbouiCOM.EditText)(this.GetItem("ETMERCNM").Specific));
            this.ETDONTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDONTRY").Specific));
            this.OnCustomInitialize();

        }

        public override void OnInitializeFormEvents()
        {
            this.ResizeAfter += new ResizeAfterHandler(this.Form_ResizeAfter);

        }


        private void OnCustomInitialize()
        {

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
                grdClr.Width = 150;

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
                int insideMargin =10;         // gap after folder header

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
