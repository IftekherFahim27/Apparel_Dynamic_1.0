using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Apparel_Dynamic_1._0.Helper;

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
                                     STMERDCD, STBUYRCD, STSBDVSN, STUOM, STSLDESC,
                                     STSZTPCD, STSUBCLR;

       



        // -------- Edit Text --------
        private SAPbouiCOM.EditText ETSLCODE, ETCSCODE, ETCSDESC, ETGSM, ETGENDER,
                                    ETSMTPCD, ETSMTPNM, ETPDTPCD, ETPDTPNM,
                                    ETPDLNCD, ETPDLNNM, ETPDGPCD, ETPDGPNM,
                                    ETBRNDCD, ETBRNDNM, ETDEPTCD, ETDEPTNM,
                                    ETDOCTRY, ETDOCNUM, ETSMPLCD, ETSMPLNM,
                                    ETUOM, ETPDLDTM, ETHSCODE, ETRTSGCD, ETSLDESC,
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
            //      -------- Static Text --------
            this.STSLCODE = ((SAPbouiCOM.StaticText)(this.GetItem("STSLCODE").Specific));
            this.STCSCODE = ((SAPbouiCOM.StaticText)(this.GetItem("STCSCODE").Specific));
            this.STCSDESC = ((SAPbouiCOM.StaticText)(this.GetItem("STCSDESC").Specific));
            this.STGSM = ((SAPbouiCOM.StaticText)(this.GetItem("STGSM").Specific));
            this.STGENDER = ((SAPbouiCOM.StaticText)(this.GetItem("STGENDER").Specific));
            this.STSMTPCD = ((SAPbouiCOM.StaticText)(this.GetItem("STSMTPCD").Specific));
            this.STPDTPCD = ((SAPbouiCOM.StaticText)(this.GetItem("STPDTPCD").Specific));
            this.STPDLNCD = ((SAPbouiCOM.StaticText)(this.GetItem("STPDLNCD").Specific));
            this.STPDGPCD = ((SAPbouiCOM.StaticText)(this.GetItem("STPDGPCD").Specific));
            this.STBRNDCD = ((SAPbouiCOM.StaticText)(this.GetItem("STBRNDCD").Specific));
            this.STDEPTCD = ((SAPbouiCOM.StaticText)(this.GetItem("STDEPTCD").Specific));
            this.STDOCNUM = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCNUM").Specific));
            this.STSMPBSE = ((SAPbouiCOM.StaticText)(this.GetItem("STSMPBSE").Specific));
            this.STSMPLCD = ((SAPbouiCOM.StaticText)(this.GetItem("STSMPLCD").Specific));
            this.STPDLDTM = ((SAPbouiCOM.StaticText)(this.GetItem("STPDLDTM").Specific));
            this.STHSCODE = ((SAPbouiCOM.StaticText)(this.GetItem("STHSCODE").Specific));
            this.STRTSGCD = ((SAPbouiCOM.StaticText)(this.GetItem("STRTSGCD").Specific));
            this.STMERDCD = ((SAPbouiCOM.StaticText)(this.GetItem("STMERDCD").Specific));
            this.STBUYRCD = ((SAPbouiCOM.StaticText)(this.GetItem("STBUYRCD").Specific));
            this.STSBDVSN = ((SAPbouiCOM.StaticText)(this.GetItem("STSBDVSN").Specific));
            this.STUOM = ((SAPbouiCOM.StaticText)(this.GetItem("STUOM").Specific));
            this.STSZTPCD = ((SAPbouiCOM.StaticText)(this.GetItem("STSZTPCD").Specific));
            this.STSUBCLR = ((SAPbouiCOM.StaticText)(this.GetItem("STSUBCLR").Specific));
            //      -------- Edit Text --------
            this.ETSLCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETSLCODE").Specific));
            this.ETCSCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETCSCODE").Specific));
            this.ETCSDESC = ((SAPbouiCOM.EditText)(this.GetItem("ETCSDESC").Specific));
            this.ETGSM = ((SAPbouiCOM.EditText)(this.GetItem("ETGSM").Specific));
            this.ETGENDER = ((SAPbouiCOM.EditText)(this.GetItem("ETGENDER").Specific));
            this.ETGENDER.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETGENDER_ChooseFromListAfter);
            this.ETGENDER.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETGENDER_ChooseFromListBefore);
            this.ETSMTPCD = ((SAPbouiCOM.EditText)(this.GetItem("ETSMTPCD").Specific));
            this.ETSMTPCD.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETSMTPCD_ChooseFromListAfter);
            this.ETSMTPCD.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETSMTPCD_ChooseFromListBefore);
            this.ETSMTPNM = ((SAPbouiCOM.EditText)(this.GetItem("ETSMTPNM").Specific));
            this.ETPDTPCD = ((SAPbouiCOM.EditText)(this.GetItem("ETPDTPCD").Specific));
            this.ETPDTPNM = ((SAPbouiCOM.EditText)(this.GetItem("ETPDTPNM").Specific));
            this.ETPDLNCD = ((SAPbouiCOM.EditText)(this.GetItem("ETPDLNCD").Specific));
            this.ETPDLNNM = ((SAPbouiCOM.EditText)(this.GetItem("ETPDLNNM").Specific));
            this.ETPDGPCD = ((SAPbouiCOM.EditText)(this.GetItem("ETPDGPCD").Specific));
            this.ETPDGPCD.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETPDGPCD_ChooseFromListAfter);
            this.ETPDGPCD.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETPDGPCD_ChooseFromListBefore);
            this.ETPDGPNM = ((SAPbouiCOM.EditText)(this.GetItem("ETPDGPNM").Specific));
            this.ETBRNDCD = ((SAPbouiCOM.EditText)(this.GetItem("ETBRNDCD").Specific));
            this.ETBRNDCD.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETBRNDCD_ChooseFromListAfter);
            this.ETBRNDCD.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETBRNDCD_ChooseFromListBefore);
            this.ETBRNDNM = ((SAPbouiCOM.EditText)(this.GetItem("ETBRNDNM").Specific));
            this.ETDEPTCD = ((SAPbouiCOM.EditText)(this.GetItem("ETDEPTCD").Specific));
            this.ETDEPTCD.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETDEPTCD_ChooseFromListAfter);
            this.ETDEPTNM = ((SAPbouiCOM.EditText)(this.GetItem("ETDEPTNM").Specific));
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            this.ETDOCNUM = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCNUM").Specific));
            this.ETSMPLCD = ((SAPbouiCOM.EditText)(this.GetItem("ETSMPLCD").Specific));
            this.ETSMPLCD.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETSMPLCD_ChooseFromListAfter);
            this.ETSMPLCD.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETSMPLCD_ChooseFromListBefore);
            this.ETSMPLNM = ((SAPbouiCOM.EditText)(this.GetItem("STSMPLNM ").Specific));
            this.ETUOM = ((SAPbouiCOM.EditText)(this.GetItem("ETUOM").Specific));
            this.ETPDLDTM = ((SAPbouiCOM.EditText)(this.GetItem("ETPDLDTM").Specific));
            this.ETHSCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETHSCODE").Specific));
            this.ETRTSGCD = ((SAPbouiCOM.EditText)(this.GetItem("ETRTSGCD").Specific));
            this.ETRTSGCD.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETRTSGCD_ChooseFromListAfter);
            this.ETRTSGCD.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETRTSGCD_ChooseFromListBefore);
            this.ETRTSGNM = ((SAPbouiCOM.EditText)(this.GetItem("ETRTSGNM").Specific));
            this.ETMERDCD = ((SAPbouiCOM.EditText)(this.GetItem("ETMERDCD").Specific));
            this.ETMERDCD.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETMERDCD_ChooseFromListAfter);
            this.ETMERDNM = ((SAPbouiCOM.EditText)(this.GetItem("ETMERDNM").Specific));
            this.ETBUYRCD = ((SAPbouiCOM.EditText)(this.GetItem("ETBUYRCD").Specific));
            this.ETBUYRCD.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETBUYRCD_ChooseFromListAfter);
            this.ETBUYRNM = ((SAPbouiCOM.EditText)(this.GetItem("ETBUYRNM").Specific));
            this.ETSDSNCD = ((SAPbouiCOM.EditText)(this.GetItem("ETSDSNCD").Specific));
            this.ETSDSNCD.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETSDSNCD_ChooseFromListAfter);
            this.ETSDSNCD.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETSDSNCD_ChooseFromListBefore);
            this.ETSDSNNM = ((SAPbouiCOM.EditText)(this.GetItem("ETSDSNNM").Specific));
            this.ETSZTPCD = ((SAPbouiCOM.EditText)(this.GetItem("ETSZTPCD").Specific));
            this.ETSZTPCD.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETSZTPCD_ChooseFromListAfter);
            this.ETSZTPCD.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETSZTPCD_ChooseFromListBefore);
            //      -------- ComboBox --------
            this.CBSERIES = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSERIES").Specific));
            this.CBSMPBSE = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSMPBSE").Specific));
            //      -------- Folder --------
            this.FOLSIZE = ((SAPbouiCOM.Folder)(this.GetItem("FOLSIZE").Specific));
            this.FOLCOLOR = ((SAPbouiCOM.Folder)(this.GetItem("FOLCOLOR").Specific));
            this.FOLITEM = ((SAPbouiCOM.Folder)(this.GetItem("FOLITEM").Specific));
            this.FOLATTAC = ((SAPbouiCOM.Folder)(this.GetItem("FOLATTAC").Specific));
            //      -------- Matrix --------
            this.MTXSIZE = ((SAPbouiCOM.Matrix)(this.GetItem("MTXSIZE").Specific));
            this.MTXCOLOR = ((SAPbouiCOM.Matrix)(this.GetItem("MTXCOLOR").Specific));
            this.MTXCOLOR.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.MTXCOLOR_ChooseFromListAfter);
            this.MTXCOLOR.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.MTXCOLOR_ChooseFromListBefore);
            this.MTXSBCLR = ((SAPbouiCOM.Matrix)(this.GetItem("MTXSBCLR").Specific));
            this.MTXITEM = ((SAPbouiCOM.Matrix)(this.GetItem("MTXITEM").Specific));
            this.MTXATTCH = ((SAPbouiCOM.Matrix)(this.GetItem("MTXATTCH").Specific));
            //      -------- Button --------
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.BTNITMTX = ((SAPbouiCOM.Button)(this.GetItem("BTNITMTX").Specific));
            this.BTNITMCR = ((SAPbouiCOM.Button)(this.GetItem("BTNITMCR").Specific));
            this.BTNLODSZ = ((SAPbouiCOM.Button)(this.GetItem("BTNLODSZ").Specific));
            this.BRWSBTN = ((SAPbouiCOM.Button)(this.GetItem("BRWSBTN").Specific));
            this.DISPBTN = ((SAPbouiCOM.Button)(this.GetItem("DISPBTN").Specific));
            this.DELBTN = ((SAPbouiCOM.Button)(this.GetItem("DELBTN").Specific));
            this.STSLDESC = ((SAPbouiCOM.StaticText)(this.GetItem("STSLDESC").Specific));
            this.ETSLDESC = ((SAPbouiCOM.EditText)(this.GetItem("ETSLDESC").Specific));
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

        private void ETSMPLCD_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_SMST")
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
                    "Error filtering Sample Master CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }
        }

        private void ETSMPLCD_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("Code", 0).ToString().Trim();
            string Name = dt.GetValue("Name", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETSMPLCD").Specific;
            ETCD.Value = Code;
            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETSMPLNM").Specific;
            ETNM.Value = Name;

        }

        private void ETRTSGCD_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_ROUT")
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
                    "Error filtering Route CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }

        }

        private void ETRTSGCD_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("Code", 0).ToString().Trim();
            string Name = dt.GetValue("Name", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETRTSGCD").Specific;
            ETCD.Value = Code;
            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETRTSGNM").Specific;
            ETNM.Value = Name;

        }

        private void ETMERDCD_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("empID", 0).ToString().Trim();
            string Name = dt.GetValue("U_FNAME", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETMERDCD").Specific;
            ETCD.Value = Code;
            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETMERDNM").Specific;
            ETNM.Value = Name;

        }



        private void ETBUYRCD_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("CardCode", 0).ToString().Trim();
            string Name = dt.GetValue("CardName", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETBUYRCD").Specific;
            ETCD.Value = Code;
            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETBUYRNM").Specific;
            ETNM.Value = Name;

        }

        private void ETSDSNCD_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_DVSN")
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
                    "Error filtering Sub Division CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }

        }

        private void ETSDSNCD_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("Code", 0).ToString().Trim();
            string Name = dt.GetValue("Name", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETSDSNCD").Specific;
            ETCD.Value = Code;
            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETSDSNNM").Specific;
            ETNM.Value = Name;

        }

        private void ETSZTPCD_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_SZTP")
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
                    "Error filtering Size Type CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }

        }

        private void ETSZTPCD_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("Code", 0).ToString().Trim();
            

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETSZTPCD").Specific;
            ETCD.Value = Code;
            

        }

        private void MTXCOLOR_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID != "CFL_CLOR")
                    return;

                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

                // Matrix + Column where Color code is stored
                SAPbouiCOM.Matrix oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXCOLOR").Specific;
                string colCode = "CLCLRCOD"; 

                HashSet<string> usedCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                for (int i = 1; i <= oMtx.RowCount; i++)
                {
                    var cell = (SAPbouiCOM.EditText)oMtx.Columns.Item(colCode).Cells.Item(i).Specific;
                    string code = (cell.Value ?? "").Trim();
                    if (!string.IsNullOrEmpty(code))
                        usedCodes.Add(code);
                }

                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(cflUID);
                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                //Active only
                SAPbouiCOM.Condition cActive = oCons.Add();
                cActive.Alias = "U_ACTIVE";
                cActive.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                cActive.CondVal = "Y";

                // If there are used codes, add AND + NOT EQUAL for each
                if (usedCodes.Count > 0)
                {
                    cActive.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    int idx = 0;
                    foreach (string c in usedCodes)
                    {
                        SAPbouiCOM.Condition cond = oCons.Add();
                        cond.Alias = "Code"; // or "U_SIZECODE" if CFL object doesn't have Code
                        cond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                        cond.CondVal = c;

                        idx++;
                        if (idx < usedCodes.Count)
                            cond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                    }
                }

                oCFL.SetConditions(oCons);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error filtering Colour CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }

        }

        private void MTXCOLOR_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXCOLOR").Specific;
                SAPbouiCOM.DBDataSource DBDataSourceLine = oForm.DataSources.DBDataSources.Item("@FIL_DR_SMPLCOLO");
                SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;

                if (dt == null || dt.Rows.Count == 0)
                    return;

                string code = dt.GetValue("Code", 0).ToString();
                string name = dt.GetValue("Name", 0).ToString();
                string pantone = dt.GetValue("U_PANTONE", 0).ToString();
                int row = pVal.Row;
                //Set Values
                oMatrix.SetCellWithoutValidation(row, "CLCLRCOD", code);
                oMatrix.SetCellWithoutValidation(row, "CLCLRNAM", name);
                oMatrix.SetCellWithoutValidation(row, "CLPANTON", pantone);
                oMatrix.FlushToDataSource();

                // Add new row if last row has data
                int lastRow = oMatrix.RowCount;
                bool lastRowHasData = !string.IsNullOrWhiteSpace(((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLCLRCOD").Cells.Item(lastRow).Specific).Value);
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
                Application.SBO_Application.StatusBar.SetText("Color Matrix CFL Error: " + ex.Message,
                   SAPbouiCOM.BoMessageTime.bmt_Short,
                   SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void ETDEPTCD_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("Code", 0).ToString().Trim();
            string Name = dt.GetValue("Name", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETDEPTCD").Specific;
            ETCD.Value = Code;
            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETDEPTNM").Specific;
            ETNM.Value = Name;
        }



        private void ETBRNDCD_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("Code", 0).ToString().Trim();
            string Name = dt.GetValue("Name", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETBRNDCD").Specific;
            ETCD.Value = Code;
            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETBRNDNM").Specific;
            ETNM.Value = Name;

        }

        private void ETBRNDCD_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_BRND")
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
                    "Error filtering Brand CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }
        }

        private void ETPDGPCD_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;

                if (dt == null || dt.Rows.Count == 0)
                    return;

                string code = dt.GetValue("Code", 0).ToString().Trim();
                string name = dt.GetValue("Name", 0).ToString().Trim();
                string prdLine = dt.GetValue("U_PRDLINE", 0).ToString().Trim(); 
                string prdType = dt.GetValue("U_PRDTYPE", 0).ToString().Trim();

                ((SAPbouiCOM.EditText)oForm.Items.Item("ETPDGPCD").Specific).Value = code;
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETPDGPNM").Specific).Value = name;
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETPDLNCD").Specific).Value = prdLine;
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETPDTPCD").Specific).Value = prdType;

                string prdLineName = "";
                string prdTypeName = "";

                SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                // Product Line Name from @FIL_MH_OPLM
                if (!string.IsNullOrWhiteSpace(prdLine))
                {
                    string q1 = $@"SELECT ""Name"" FROM ""@FIL_MH_OPLM"" WHERE ""Code"" = '{prdLine}'";
                    oRS.DoQuery(q1);

                    if (!oRS.EoF)
                        prdLineName = oRS.Fields.Item("Name").Value.ToString().Trim();
                }

                // Product Type Name from @FIL_MH_PRDTYPE
                if (!string.IsNullOrWhiteSpace(prdType))
                {
                    string q2 = $@"SELECT ""Name"" FROM ""@FIL_MH_PRDTYPE"" WHERE ""Code"" = '{prdType}'";
                    oRS.DoQuery(q2);

                    if (!oRS.EoF)
                        prdTypeName = oRS.Fields.Item("Name").Value.ToString().Trim();
                }

                ((SAPbouiCOM.EditText)oForm.Items.Item("ETPDLNNM").Specific).Value = prdLineName;
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETPDTPNM").Specific).Value = prdTypeName;
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error in CFL After event: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void ETPDGPCD_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                string gender = ((SAPbouiCOM.EditText)oForm.Items.Item("ETGENDER").Specific).Value.Trim();

                // Validation
                if (string.IsNullOrEmpty(gender))
                {
                    Application.SBO_Application.MessageBox("Select Gender First.");
                    oForm.Items.Item("ETGENDER").Click();
                    BubbleEvent = false;
                    return;
                }

                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_PGRP")
                {
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(cflUID);

                    SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                    // Condition 1 
                    SAPbouiCOM.Condition oCon1 = oCons.Add();
                    oCon1.Alias = "U_ACTIVE";
                    oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon1.CondVal = "Y";
                    oCon1.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    // Condition 2 
                    SAPbouiCOM.Condition oCon2 = oCons.Add();
                    oCon2.Alias = "U_GENDER";
                    oCon2.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon2.CondVal = gender;
                    

                    oCFL.SetConditions(oCons);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error filtering Product Group CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
        }





        private void ETSMTPCD_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("Code", 0).ToString().Trim();
            string Name = dt.GetValue("Name", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETSMTPCD").Specific;
            ETCD.Value = Code;
            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETSMTPNM").Specific;
            ETNM.Value = Name;

        }

        private void ETSMTPCD_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_SMPTP")
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
                    "Error filtering Sample Type CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }

        }

        private void ETGENDER_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("Code", 0).ToString();
            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETGENDER").Specific;
            ETCD.Value = Code;

        }

        private void ETGENDER_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_GEN")
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
                    "Error filtering Gender CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }

        }

        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = null;

                
        }
    }
}
