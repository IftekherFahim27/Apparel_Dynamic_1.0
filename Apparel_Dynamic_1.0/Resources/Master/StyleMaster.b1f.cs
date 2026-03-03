using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apparel_Dynamic_1._0.Resources.Master
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Master.StyleMaster", "Resources/Master/StyleMaster.b1f")]
    class StyleMaster : UserFormBase
    {
        public StyleMaster()
        {
        }

        // -------- Static Text --------
        private SAPbouiCOM.StaticText STSLCODE, STCSCODE, STCSDESC, STGSM, STGENDER,
                                     STSMTPCD, STPDTPCD, STPDLNCD, STPDGPCD,
                                     STBRNDCD, STDEPTCD, STDOCNUM, STSMPBSE,
                                     STSMPLCD, STPDLDTM, STHSCODE, STRTSGCD,
                                     STMERDCD, STBUYRCD, STSBDVSN, STUOM,
                                     STSZTPCD, STSUBCLR;

        // -------- Edit Text --------
        private SAPbouiCOM.EditText ETSLCODE, ETCSCODE, ETCSDESC, ETGSM, ETGENDER,
                                    ETSMTPCD, ETSMTPNM, ETPDTPCD, ETPDTPNM,
                                    ETPDLNCD, ETPDLNNM, ETPDGPCD, ETPDGPNM,
                                    ETBRNDCD, ETBRNDNM, ETDEPTCD, ETDEPTNM,
                                    ETDOCTRY, ETDOCNUM, ETSMPLCD, ETSMPLNM,
                                    ETUOM, ETPDLDTM, ETHSCODE, ETRTSGCD,
                                    ETRTSGNM, ETMERDCD, ETMERDNM, ETBUYRCD,
                                    ETBUYRNM, ETSDSNCD, ETSDSNNM, ETSZTPCD;

        // -------- ComboBox --------
        private SAPbouiCOM.ComboBox CBSERIES, CBSMPBSE;

        // -------- Folder --------
        private SAPbouiCOM.Folder FOLSIZE, FOLCOLOR, FOLITEM, FOLATTAC;

        // -------- Matrix --------
        private SAPbouiCOM.Matrix MTXSIZE, MTXCOLOR, MTXSBCLR, MTXITEM, MTXATTCH;

        // -------- Button --------
        private SAPbouiCOM.Button ADDButton, CancelButton, BTNITMTX, BTNITMCR,
                                  BTNLODSZ, BRWSBTN, DISPBTN, DELBTN;

        public override void OnInitializeComponent()
        {
            // -------- Static Text --------
            this.STSLCODE = (SAPbouiCOM.StaticText)this.GetItem("STSLCODE").Specific;
            this.STCSCODE = (SAPbouiCOM.StaticText)this.GetItem("STCSCODE").Specific;
            this.STCSDESC = (SAPbouiCOM.StaticText)this.GetItem("STCSDESC").Specific;
            this.STGSM = (SAPbouiCOM.StaticText)this.GetItem("STGSM").Specific;
            this.STGENDER = (SAPbouiCOM.StaticText)this.GetItem("STGENDER").Specific;
            this.STSMTPCD = (SAPbouiCOM.StaticText)this.GetItem("STSMTPCD").Specific;
            this.STPDTPCD = (SAPbouiCOM.StaticText)this.GetItem("STPDTPCD").Specific;
            this.STPDLNCD = (SAPbouiCOM.StaticText)this.GetItem("STPDLNCD").Specific;
            this.STPDGPCD = (SAPbouiCOM.StaticText)this.GetItem("STPDGPCD").Specific;
            this.STBRNDCD = (SAPbouiCOM.StaticText)this.GetItem("STBRNDCD").Specific;
            this.STDEPTCD = (SAPbouiCOM.StaticText)this.GetItem("STDEPTCD").Specific;
            this.STDOCNUM = (SAPbouiCOM.StaticText)this.GetItem("STDOCNUM").Specific;
            this.STSMPBSE = (SAPbouiCOM.StaticText)this.GetItem("STSMPBSE").Specific;
            this.STSMPLCD = (SAPbouiCOM.StaticText)this.GetItem("STSMPLCD").Specific;
            this.STPDLDTM = (SAPbouiCOM.StaticText)this.GetItem("STPDLDTM").Specific;
            this.STHSCODE = (SAPbouiCOM.StaticText)this.GetItem("STHSCODE").Specific;
            this.STRTSGCD = (SAPbouiCOM.StaticText)this.GetItem("STRTSGCD").Specific;
            this.STMERDCD = (SAPbouiCOM.StaticText)this.GetItem("STMERDCD").Specific;
            this.STBUYRCD = (SAPbouiCOM.StaticText)this.GetItem("STBUYRCD").Specific;
            this.STSBDVSN = (SAPbouiCOM.StaticText)this.GetItem("STSBDVSN").Specific;
            this.STUOM = (SAPbouiCOM.StaticText)this.GetItem("STUOM").Specific;
            this.STSZTPCD = (SAPbouiCOM.StaticText)this.GetItem("STSZTPCD").Specific;
            this.STSUBCLR = (SAPbouiCOM.StaticText)this.GetItem("STSUBCLR").Specific;

            // -------- Edit Text --------
            this.ETSLCODE = (SAPbouiCOM.EditText)this.GetItem("ETSLCODE").Specific;
            this.ETCSCODE = (SAPbouiCOM.EditText)this.GetItem("ETCSCODE").Specific;
            this.ETCSDESC = (SAPbouiCOM.EditText)this.GetItem("ETCSDESC").Specific;
            this.ETGSM = (SAPbouiCOM.EditText)this.GetItem("ETGSM").Specific;
            this.ETGENDER = (SAPbouiCOM.EditText)this.GetItem("ETGENDER").Specific;
            this.ETSMTPCD = (SAPbouiCOM.EditText)this.GetItem("ETSMTPCD").Specific;
            this.ETSMTPNM = (SAPbouiCOM.EditText)this.GetItem("ETSMTPNM").Specific;
            this.ETPDTPCD = (SAPbouiCOM.EditText)this.GetItem("ETPDTPCD").Specific;
            this.ETPDTPNM = (SAPbouiCOM.EditText)this.GetItem("ETPDTPNM").Specific;
            this.ETPDLNCD = (SAPbouiCOM.EditText)this.GetItem("ETPDLNCD").Specific;
            this.ETPDLNNM = (SAPbouiCOM.EditText)this.GetItem("ETPDLNNM").Specific;
            this.ETPDGPCD = (SAPbouiCOM.EditText)this.GetItem("ETPDGPCD").Specific;
            this.ETPDGPNM = (SAPbouiCOM.EditText)this.GetItem("ETPDGPNM").Specific;
            this.ETBRNDCD = (SAPbouiCOM.EditText)this.GetItem("ETBRNDCD").Specific;
            this.ETBRNDNM = (SAPbouiCOM.EditText)this.GetItem("STBRNDNM").Specific; // (your item id is STBRNDNM)
            this.ETDEPTCD = (SAPbouiCOM.EditText)this.GetItem("ETDEPTCD").Specific;
            this.ETDEPTNM = (SAPbouiCOM.EditText)this.GetItem("ETDEPTNM").Specific;
            this.ETDOCTRY = (SAPbouiCOM.EditText)this.GetItem("ETDOCTRY").Specific;
            this.ETDOCNUM = (SAPbouiCOM.EditText)this.GetItem("ETDOCNUM").Specific;
            this.ETSMPLCD = (SAPbouiCOM.EditText)this.GetItem("ETSMPLCD").Specific;
            this.ETSMPLNM = (SAPbouiCOM.EditText)this.GetItem("STSMPLNM ").Specific; // NOTE: your ID has a trailing space!
            this.ETUOM = (SAPbouiCOM.EditText)this.GetItem("ETUOM").Specific;
            this.ETPDLDTM = (SAPbouiCOM.EditText)this.GetItem("ETPDLDTM").Specific;
            this.ETHSCODE = (SAPbouiCOM.EditText)this.GetItem("ETHSCODE").Specific;
            this.ETRTSGCD = (SAPbouiCOM.EditText)this.GetItem("ETRTSGCD").Specific;
            this.ETRTSGNM = (SAPbouiCOM.EditText)this.GetItem("ETRTSGNM").Specific;
            this.ETMERDCD = (SAPbouiCOM.EditText)this.GetItem("ETMERDCD").Specific;
            this.ETMERDNM = (SAPbouiCOM.EditText)this.GetItem("ETMERDNM").Specific;
            this.ETBUYRCD = (SAPbouiCOM.EditText)this.GetItem("ETBUYRCD").Specific;
            this.ETBUYRNM = (SAPbouiCOM.EditText)this.GetItem("ETBUYRNM").Specific;
            this.ETSDSNCD = (SAPbouiCOM.EditText)this.GetItem("ETSDSNCD").Specific;
            this.ETSDSNNM = (SAPbouiCOM.EditText)this.GetItem("ETSDSNNM").Specific;
            this.ETSZTPCD = (SAPbouiCOM.EditText)this.GetItem("ETSZTPCD").Specific;

            // -------- ComboBox --------
            this.CBSERIES = (SAPbouiCOM.ComboBox)this.GetItem("CBSERIES").Specific;
            this.CBSMPBSE = (SAPbouiCOM.ComboBox)this.GetItem("CBSMPBSE").Specific;

            // -------- Folder --------
            this.FOLSIZE = (SAPbouiCOM.Folder)this.GetItem("FOLSIZE").Specific;
            this.FOLCOLOR = (SAPbouiCOM.Folder)this.GetItem("FOLCOLOR").Specific;
            this.FOLITEM = (SAPbouiCOM.Folder)this.GetItem("FOLITEM").Specific;
            this.FOLATTAC = (SAPbouiCOM.Folder)this.GetItem("FOLATTAC").Specific;

            // -------- Matrix --------
            this.MTXSIZE = (SAPbouiCOM.Matrix)this.GetItem("MTXSIZE").Specific;
            this.MTXCOLOR = (SAPbouiCOM.Matrix)this.GetItem("MTXCOLOR").Specific;
            this.MTXSBCLR = (SAPbouiCOM.Matrix)this.GetItem("MTXSBCLR").Specific;
            this.MTXITEM = (SAPbouiCOM.Matrix)this.GetItem("MTXITEM").Specific;
            this.MTXATTCH = (SAPbouiCOM.Matrix)this.GetItem("MTXATTCH").Specific;

            // -------- Button --------
            this.ADDButton = (SAPbouiCOM.Button)this.GetItem("1").Specific;
            this.CancelButton = (SAPbouiCOM.Button)this.GetItem("2").Specific;

            this.BTNITMTX = (SAPbouiCOM.Button)this.GetItem("BTNITMTX").Specific;
            this.BTNITMCR = (SAPbouiCOM.Button)this.GetItem("BTNITMCR").Specific;
            this.BTNLODSZ = (SAPbouiCOM.Button)this.GetItem("BTNLODSZ").Specific;
            this.BRWSBTN = (SAPbouiCOM.Button)this.GetItem("BRWSBTN").Specific;
            this.DISPBTN = (SAPbouiCOM.Button)this.GetItem("DISPBTN").Specific;
            this.DELBTN = (SAPbouiCOM.Button)this.GetItem("DELBTN").Specific;

            this.OnCustomInitialize();
        }
        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
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
                if (oForm == null) return;

                oForm.Freeze(true);

                // --- IDs ---
                string folderId = "FOLCOLOR";     // your Folder(tab)
                string paneAreaId = "2";   // <<< CHANGE THIS to your tab content area item id (8,178,970,210)

                string m1Id = "MTXCOLOR";
                string m2Id = "MTXSBCLR";

                // --- Keep TOP fixed as you requested ---
                int m1TopFixed = 215;
                int m2TopFixed = 235;

                // Heights can stay fixed too (or clamp if tab small)
                int m1HeightFixed = 150;
                int m2HeightFixed = 130;

                int paddingLeft = 12;
                int paddingRight = 12;
                int gapBetween = 12;

                // Put matrices on the folder pane (important)
                var fol = (SAPbouiCOM.Folder)oForm.Items.Item(folderId).Specific;
                int pane = fol.Pane;

                oForm.Items.Item(m1Id).FromPane = pane; oForm.Items.Item(m1Id).ToPane = pane;
                oForm.Items.Item(m2Id).FromPane = pane; oForm.Items.Item(m2Id).ToPane = pane;

                // --- Get REAL tab area bounds ---
                var area = oForm.Items.Item(paneAreaId);
                int L = area.Left + paddingLeft;
                int T = area.Top;
                int R = area.Left + area.Width - paddingRight;
                int B = area.Top + area.Height;

                int usableW = Math.Max(1, R - L);
                int halfW = usableW / 2;

                // --- Left half rect ---
                int leftL = L;
                int leftR = L + halfW - (gapBetween / 2);

                // --- Right half rect ---
                int rightL = L + halfW + (gapBetween / 2);
                int rightR = R;

                // Safety if form becomes too small
                if (leftR < leftL) leftR = leftL + 1;
                if (rightR < rightL) rightR = rightL + 1;

                // --- Apply to MTZCOLOR (left half) ---
                var m1 = oForm.Items.Item(m1Id);
                m1.Left = leftL;                 // always inside left half
                m1.Top = m1TopFixed;            // FIXED top
                m1.Width = Math.Max(1, leftR - leftL);

                // keep height fixed but clamp inside tab bottom (so it won't go outside)
                m1.Height = Math.Min(m1HeightFixed, Math.Max(1, B - m1.Top));

                // --- Apply to MTXSBCLR (right half) ---
                var m2 = oForm.Items.Item(m2Id);
                m2.Left = rightL;                // always inside right half
                m2.Top = m2TopFixed;            // FIXED top
                m2.Width = Math.Max(1, rightR - rightL);

                m2.Height = Math.Min(m2HeightFixed, Math.Max(1, B - m2.Top));
            }
            catch(Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Form_DataLoadAfter error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                try { if (oForm != null) oForm.Freeze(false); } catch { }
            }
        }
    }
}
