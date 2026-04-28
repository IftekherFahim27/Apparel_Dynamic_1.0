using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Apparel_Dynamic_1._0.Helper;
using Apparel_Dynamic_1._0.Resources.Version;

namespace Apparel_Dynamic_1._0.Resources.Transaction
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Transaction.SalesContract", "Resources/Transaction/SalesContract.b1f")]
    class SalesContract : UserFormBase
    {
        public SalesContract()
        {
        }

        // -------- Static Text --------
        private SAPbouiCOM.StaticText STBRANCH, STMRSTAT, STCMSTAT, STCUSTMR, STBRNDCD,
                                     STSCNO, STSCDESC, STREFNCE, STSCVAL, STDOVAL,
                                     STB2BPER, STB2BVAL, STDSNBNK, STDOCNUM, STDOCDAT,
                                     STCBNKAC, STCUSBNK, STOBNKAC, STOWNBNK, STCURR,
                                     STISUDAT, STSHPDAT, STEXPDAT, STTOLPER, STAMNDNO,
                                     STOPNAMT, STINTRMS, STPYTRMS, STMDSHIP, STPORTLD,
                                     STCNDEST, STPRTDIS, STINSNCE, STSHPTOL, STHSCODE,
                                     STDOCREQ, STRMSCON, STSHPADD, STPRTSHP;

       



        // -------- ComboBox --------
        private SAPbouiCOM.ComboBox CBBRANCH, CBMRSTAT, CBCMSTAT, CBINTRMS, CBDSNBNK,
                                    CBPYTRMS, CBMDSHIP, CBPRTSHP, CBSERIES, CBSHPADD;



        // -------- Edit Text --------
        private SAPbouiCOM.EditText ETCUSTMR, ETCUSTNM, ETBRNDCD, ETBRNDNM, ETSCNO,
                                    ETSCDESC, ETREFNCE, ETSCVAL, ETDOVAL, ETB2BPER,
                                    ETB2BVAL, ETDOCNUM, ETDOCDAT, ETCBNKAC,
                                    ETCUSBNK, ETOBNKAC, ETOWNBNK, ETCURR, ETISUDAT,
                                    ETSHPDAT, ETEXPDAT, ETTOLPER, ETAMNDNO, ETOPNAMT,
                                    ETPORTLD, ETCNDEST, ETPRTDIS, ETINSNCE, ETSHPTOL,
                                    ETHSCODE, ETDOCREQ, ETRMSCON, ETDOCTRY;

        // -------- Folder --------
        private SAPbouiCOM.Folder FOLORDTL, FOLAMEND, FOLB2BDL, FOLATTCH, FOLCDTLS;

        // -------- Button --------
        private SAPbouiCOM.Button ADDButton, CancelButton, BTNAMEND, BRWSBTN, DISPBTN, DELBTN;

        // -------- Matrix --------
        private SAPbouiCOM.Matrix MTXORDTL, MTXATTCH, MTXB2BDL;

        // -------- Grid --------
        private SAPbouiCOM.Grid GRDAMEND;
        public override void OnInitializeComponent()
        {
            //               -------- Static Text --------
            this.STBRANCH = ((SAPbouiCOM.StaticText)(this.GetItem("STBRANCH").Specific));
            this.STMRSTAT = ((SAPbouiCOM.StaticText)(this.GetItem("STMRSTAT").Specific));
            this.STCMSTAT = ((SAPbouiCOM.StaticText)(this.GetItem("STCMSTAT").Specific));
            this.STCUSTMR = ((SAPbouiCOM.StaticText)(this.GetItem("STCUSTMR").Specific));
            this.STBRNDCD = ((SAPbouiCOM.StaticText)(this.GetItem("STBRNDCD").Specific));
            this.STSCNO = ((SAPbouiCOM.StaticText)(this.GetItem("STSCNO").Specific));
            this.STSCDESC = ((SAPbouiCOM.StaticText)(this.GetItem("STSCDESC").Specific));
            this.STREFNCE = ((SAPbouiCOM.StaticText)(this.GetItem("STREFNCE").Specific));
            this.STSCVAL = ((SAPbouiCOM.StaticText)(this.GetItem("STSCVAL").Specific));
            this.STDOVAL = ((SAPbouiCOM.StaticText)(this.GetItem("STDOVAL").Specific));
            this.STB2BPER = ((SAPbouiCOM.StaticText)(this.GetItem("STB2BPER").Specific));
            this.STB2BVAL = ((SAPbouiCOM.StaticText)(this.GetItem("STB2BVAL").Specific));
            this.STDSNBNK = ((SAPbouiCOM.StaticText)(this.GetItem("STDSNBNK").Specific));
            this.STDOCNUM = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCNUM").Specific));
            this.STDOCDAT = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCDAT").Specific));
            this.STCBNKAC = ((SAPbouiCOM.StaticText)(this.GetItem("STCBNKAC").Specific));
            this.STCUSBNK = ((SAPbouiCOM.StaticText)(this.GetItem("STCUSBNK").Specific));
            this.STOBNKAC = ((SAPbouiCOM.StaticText)(this.GetItem("STOBNKAC").Specific));
            this.STOWNBNK = ((SAPbouiCOM.StaticText)(this.GetItem("STOWNBNK").Specific));
            this.STCURR = ((SAPbouiCOM.StaticText)(this.GetItem("STCURR").Specific));
            this.STISUDAT = ((SAPbouiCOM.StaticText)(this.GetItem("STISUDAT").Specific));
            this.STSHPDAT = ((SAPbouiCOM.StaticText)(this.GetItem("STSHPDAT").Specific));
            this.STEXPDAT = ((SAPbouiCOM.StaticText)(this.GetItem("STEXPDAT").Specific));
            this.STTOLPER = ((SAPbouiCOM.StaticText)(this.GetItem("STTOLPER").Specific));
            this.STAMNDNO = ((SAPbouiCOM.StaticText)(this.GetItem("STAMNDNO").Specific));
            this.STOPNAMT = ((SAPbouiCOM.StaticText)(this.GetItem("STOPNAMT").Specific));
            this.STINTRMS = ((SAPbouiCOM.StaticText)(this.GetItem("STINTRMS").Specific));
            this.STPYTRMS = ((SAPbouiCOM.StaticText)(this.GetItem("STPYTRMS").Specific));
            this.STMDSHIP = ((SAPbouiCOM.StaticText)(this.GetItem("STMDSHIP").Specific));
            this.STPORTLD = ((SAPbouiCOM.StaticText)(this.GetItem("STPORTLD").Specific));
            this.STCNDEST = ((SAPbouiCOM.StaticText)(this.GetItem("STCNDEST").Specific));
            this.STPRTDIS = ((SAPbouiCOM.StaticText)(this.GetItem("STPRTDIS").Specific));
            this.STINSNCE = ((SAPbouiCOM.StaticText)(this.GetItem("STINSNCE").Specific));
            this.STSHPTOL = ((SAPbouiCOM.StaticText)(this.GetItem("STSHPTOL").Specific));
            this.STHSCODE = ((SAPbouiCOM.StaticText)(this.GetItem("STHSCODE").Specific));
            this.STDOCREQ = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCREQ").Specific));
            this.STRMSCON = ((SAPbouiCOM.StaticText)(this.GetItem("STRMSCON").Specific));
            this.STSHPADD = ((SAPbouiCOM.StaticText)(this.GetItem("STSHPADD").Specific));
            this.STPRTSHP = ((SAPbouiCOM.StaticText)(this.GetItem("STPRTSHP").Specific));
            //               -------- Folder --------
            this.FOLORDTL = ((SAPbouiCOM.Folder)(this.GetItem("FOLORDTL").Specific));
            this.FOLAMEND = ((SAPbouiCOM.Folder)(this.GetItem("FOLAMEND").Specific));
            this.FOLB2BDL = ((SAPbouiCOM.Folder)(this.GetItem("FOLB2BDL").Specific));
            this.FOLATTCH = ((SAPbouiCOM.Folder)(this.GetItem("FOLATTCH").Specific));
            this.FOLCDTLS = ((SAPbouiCOM.Folder)(this.GetItem("FOLCDTLS").Specific));
            //               -------- Button --------
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.ADDButton.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.ADDButton_PressedAfter);
            this.ADDButton.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.ADDButton_PressedBefore);
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.BTNAMEND = ((SAPbouiCOM.Button)(this.GetItem("BTNAMEND").Specific));
            this.BTNAMEND.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.BTNAMEND_PressedAfter);
            this.BRWSBTN = ((SAPbouiCOM.Button)(this.GetItem("BRWSBTN").Specific));
            this.BRWSBTN.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.BRWSBTN_ClickAfter);
            this.DISPBTN = ((SAPbouiCOM.Button)(this.GetItem("DISPBTN").Specific));
            this.DISPBTN.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.DISPBTN_ClickAfter);
            this.DELBTN = ((SAPbouiCOM.Button)(this.GetItem("DELBTN").Specific));
            this.DELBTN.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.DELBTN_ClickAfter);
            //               -------- ComboBox --------
            this.CBBRANCH = ((SAPbouiCOM.ComboBox)(this.GetItem("CBBRANCH").Specific));
            this.CBMRSTAT = ((SAPbouiCOM.ComboBox)(this.GetItem("CBMRSTAT").Specific));
            this.CBMRSTAT.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.CBMRSTAT_ComboSelectAfter);
            this.CBCMSTAT = ((SAPbouiCOM.ComboBox)(this.GetItem("CBCMSTAT").Specific));
            this.CBCMSTAT.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.CBCMSTAT_ComboSelectAfter);
            this.CBINTRMS = ((SAPbouiCOM.ComboBox)(this.GetItem("CBINTRMS").Specific));
            this.CBPYTRMS = ((SAPbouiCOM.ComboBox)(this.GetItem("CBPYTRMS").Specific));
            this.CBMDSHIP = ((SAPbouiCOM.ComboBox)(this.GetItem("CBMDSHIP").Specific));
            this.CBPRTSHP = ((SAPbouiCOM.ComboBox)(this.GetItem("CBPRTSHP").Specific));
            this.CBSERIES = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSERIES").Specific));
            this.CBDSNBNK = ((SAPbouiCOM.ComboBox)(this.GetItem("CBDSNBNK").Specific));
            this.CBSHPADD = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSHPADD").Specific));
            this.CBDSNBNK.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.CBDSNBNK_ComboSelectAfter);
            //               -------- Edit Text --------
            this.ETCUSTMR = ((SAPbouiCOM.EditText)(this.GetItem("ETCUSTMR").Specific));
            this.ETCUSTMR.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETCUSTMR_ChooseFromListBefore);
            this.ETCUSTMR.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETCUSTMR_ChooseFromListAfter);
            this.ETCUSTNM = ((SAPbouiCOM.EditText)(this.GetItem("ETCUSTNM").Specific));
            this.ETBRNDCD = ((SAPbouiCOM.EditText)(this.GetItem("ETBRNDCD").Specific));
            this.ETBRNDCD.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETBRNDCD_ChooseFromListBefore);
            this.ETBRNDCD.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETBRNDCD_ChooseFromListAfter);
            this.ETBRNDNM = ((SAPbouiCOM.EditText)(this.GetItem("ETBRNDNM").Specific));
            this.ETSCNO = ((SAPbouiCOM.EditText)(this.GetItem("ETSCNO").Specific));
            this.ETSCNO.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.ETSCNO_LostFocusAfter);
            this.ETSCDESC = ((SAPbouiCOM.EditText)(this.GetItem("ETSCDESC").Specific));
            this.ETREFNCE = ((SAPbouiCOM.EditText)(this.GetItem("ETREFNCE").Specific));
            this.ETSCVAL = ((SAPbouiCOM.EditText)(this.GetItem("ETSCVAL").Specific));
            this.ETDOVAL = ((SAPbouiCOM.EditText)(this.GetItem("ETDOVAL").Specific));
            this.ETB2BPER = ((SAPbouiCOM.EditText)(this.GetItem("ETB2BPER").Specific));
            this.ETB2BVAL = ((SAPbouiCOM.EditText)(this.GetItem("ETB2BVAL").Specific));
            this.ETDOCNUM = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCNUM").Specific));
            this.ETDOCDAT = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCDAT").Specific));
            this.ETCBNKAC = ((SAPbouiCOM.EditText)(this.GetItem("ETCBNKAC").Specific));
            this.ETCBNKAC.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETCBNKAC_ChooseFromListAfter);
            this.ETCBNKAC.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETCBNKAC_ChooseFromListBefore);
            this.ETCUSBNK = ((SAPbouiCOM.EditText)(this.GetItem("ETCUSBNK").Specific));
            this.ETOBNKAC = ((SAPbouiCOM.EditText)(this.GetItem("ETOBNKAC").Specific));
            this.ETOBNKAC.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETOBNKAC_ChooseFromListAfter);
            this.ETOWNBNK = ((SAPbouiCOM.EditText)(this.GetItem("ETOWNBNK").Specific));
            this.ETCURR = ((SAPbouiCOM.EditText)(this.GetItem("ETCURR").Specific));
            this.ETCURR.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETCURR_ChooseFromListAfter);
            this.ETISUDAT = ((SAPbouiCOM.EditText)(this.GetItem("ETISUDAT").Specific));
            this.ETSHPDAT = ((SAPbouiCOM.EditText)(this.GetItem("ETSHPDAT").Specific));
            this.ETEXPDAT = ((SAPbouiCOM.EditText)(this.GetItem("ETEXPDAT").Specific));
            this.ETTOLPER = ((SAPbouiCOM.EditText)(this.GetItem("ETTOLPER").Specific));
            this.ETAMNDNO = ((SAPbouiCOM.EditText)(this.GetItem("ETAMNDNO").Specific));
            this.ETOPNAMT = ((SAPbouiCOM.EditText)(this.GetItem("ETOPNAMT").Specific));
            this.ETPORTLD = ((SAPbouiCOM.EditText)(this.GetItem("ETPORTLD").Specific));
            this.ETCNDEST = ((SAPbouiCOM.EditText)(this.GetItem("ETCNDEST").Specific));
            this.ETCNDEST.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETCNDEST_ChooseFromListAfter);
            this.ETPRTDIS = ((SAPbouiCOM.EditText)(this.GetItem("ETPRTDIS").Specific));
            this.ETINSNCE = ((SAPbouiCOM.EditText)(this.GetItem("ETINSNCE").Specific));
            this.ETSHPTOL = ((SAPbouiCOM.EditText)(this.GetItem("ETSHPTOL").Specific));
            this.ETHSCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETHSCODE").Specific));
            this.ETDOCREQ = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCREQ").Specific));
            this.ETRMSCON = ((SAPbouiCOM.EditText)(this.GetItem("ETRMSCON").Specific));
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            //               -------- Matrix --------
            this.MTXORDTL = ((SAPbouiCOM.Matrix)(this.GetItem("MTXORDTL").Specific));
            this.MTXATTCH = ((SAPbouiCOM.Matrix)(this.GetItem("MTXATTCH").Specific));
            this.MTXB2BDL = ((SAPbouiCOM.Matrix)(this.GetItem("MTXB2BDL").Specific));
            //               -------- Grid --------
            this.GRDAMEND = ((SAPbouiCOM.Grid)(this.GetItem("GRDAMEND").Specific));
            this.GRDAMEND.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.GRDAMEND_DoubleClickAfter);
            this.OnCustomInitialize();

        }

        public override void OnInitializeFormEvents()
        {
            this.DataLoadAfter += new SAPbouiCOM.Framework.FormBase.DataLoadAfterHandler(this.Form_DataLoadAfter);
            this.DataUpdateAfter += new DataUpdateAfterHandler(this.Form_DataUpdateAfter);

        }



        private void OnCustomInitialize()
        {

        }
        private void GRDAMEND_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

            try
            {
                if (pVal.Row < 0)
                    return;

                int ret = Application.SBO_Application.MessageBox(
                    "Are you sure you want to see the Amendment Details?",
                    1, "OK", "Cancel");

                if (ret != 1)
                    return;

                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("GRDAMEND").Specific;
                SAPbouiCOM.DataTable oDT = oGrid.DataTable;

                string docEntry = oDT.GetValue("DocEntry", pVal.Row).ToString().Trim();
                string docNum = oDT.GetValue("DocNum", pVal.Row).ToString().Trim();
                string amendNo = oDT.GetValue("U_AMNDMNT", pVal.Row).ToString().Trim();
                string logInst = oDT.GetValue("LogInst", pVal.Row).ToString().Trim();

                OpenAmendmentInNewForm(docEntry, docNum, amendNo, logInst);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "GRDAMEND_DoubleClickAfter Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Long,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void OpenAmendmentInNewForm(string docEntry, string docNum, string amendNo, string logInst)
        {
            SAPbouiCOM.Form newForm = null;

            try
            {
                SalesContractAmendLog frm = new SalesContractAmendLog();
                frm.Show();

                newForm = (SAPbouiCOM.Form)frm.UIAPIRawForm;

                if (newForm == null)
                    throw new Exception("Amendment log form could not be opened.");

                newForm.Freeze(true);

                newForm.Title = "Sales Contract - Amendment " + amendNo;

                ClearAmendmentForm(newForm);

                newForm.PaneLevel = 1;

                LoadAmendmentHeader(newForm, docEntry, docNum, amendNo, logInst);
                LoadAmendmentOrderMatrix(newForm, docEntry, logInst);
                LoadAmendmentB2BMatrix(newForm, docEntry, logInst);
                LoadAmendmentAttachmentMatrix(newForm, docEntry, logInst);

                newForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE;
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "OpenAmendmentInNewForm Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Long,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (newForm != null)
                    newForm.Freeze(false);
            }
        }

        private void LoadAmendmentHeader(SAPbouiCOM.Form oForm, string docEntry, string docNum, string amendNo, string logInst)
        {
             string sql = $@"
                SELECT
                    T0.""DocEntry"",
                    T0.""DocNum"",
                    T0.""Series"",
                    S.""SeriesName"",

                    T0.""U_BRANCH"",
                    B.""BPLName"" AS ""BranchName"",

                    T0.""U_MSCSTTS"",
                    CASE T0.""U_MSCSTTS""
                        WHEN 'D' THEN 'Draft'
                        WHEN 'C' THEN 'Complete'
                        ELSE IFNULL(T0.""U_MSCSTTS"", '')
                    END AS ""MarketingStatus"",

                    T0.""U_CSCSTTS"",
                    CASE T0.""U_CSCSTTS""
                        WHEN 'D' THEN 'Draft'
                        WHEN 'C' THEN 'Complete'
                        ELSE IFNULL(T0.""U_CSCSTTS"", '')
                    END AS ""CommercialStatus"",

                    T0.""U_CARDCODE"",
                    T0.""U_CARDNAME"",
                    T0.""U_BRAND"",
                    T0.""U_BRANDN"",
                    T0.""U_SCNO"",
                    T0.""U_SCDESC"",
                    T0.""U_REFERANCE"",
                    T0.""U_SCVALUE"",
                    T0.""U_DRFTVALUE"",
                    T0.""U_BTOBLCPER"",
                    T0.""U_BTOBLCVALUE"",

                    T0.""U_DOCSNDBK"",
                    CASE T0.""U_DOCSNDBK""
                        WHEN 'Y' THEN 'Yes'
                        WHEN 'N' THEN 'No'
                        ELSE IFNULL(T0.""U_DOCSNDBK"", '')
                    END AS ""DocSendToBank"",

                    T0.""U_DOCDATE"",
                    T0.""U_CUSBKACT"",
                    T0.""U_CUSBANK"",
                    T0.""U_OWNBKACT"",
                    T0.""U_OWNBNAME"",
                    T0.""U_CURRENCY"",
                    T0.""U_ISSUEDATE"",
                    T0.""U_SHIPDATE"",
                    T0.""U_EXPDATE"",
                    T0.""U_TOLRPERC"",
                    T0.""U_AMNDMNT"",
                    T0.""U_OPBTBAMT"",

                    T0.""U_INCOTRMS"",
                    T0.""U_TRMSOFPM"",
                    P.""PymntGroup"" AS ""PaymentTermsName"",
                    T0.""U_SHIPTRMS"",
                    SH.""TrnspName"" AS ""ShipmentName"",

                    T0.""U_PRTOFLDG"",
                    T0.""U_CNTROFDT"",
                    T0.""U_PORTOFDS"",
                    T0.""U_INSRNCE"",
                    T0.""U_SHIPTLRN"",
                    T0.""U_HSCODE"",
                    T0.""U_DOCRQRD"",
                    T0.""U_TRMSCOND"",
                    T0.""U_SHPADDRS"",
                    T0.""U_PRTSHP"",

                    CASE T0.""U_PRTSHP""
                        WHEN 'A' THEN 'Allowed'
                        WHEN 'NA' THEN 'Not Allowed'
                        ELSE IFNULL(T0.""U_PRTSHP"", '')
                    END AS ""PartialShipment""

                FROM ""@AFIL_DH_OSCM"" T0
                LEFT JOIN ""NNM1"" S 
                    ON T0.""Series"" = S.""Series""
                LEFT JOIN ""OBPL"" B 
                    ON TO_NVARCHAR(T0.""U_BRANCH"") = TO_NVARCHAR(B.""BPLId"")
                LEFT JOIN ""OCTG"" P 
                    ON TO_NVARCHAR(T0.""U_TRMSOFPM"") = TO_NVARCHAR(P.""GroupNum"")
                LEFT JOIN ""OSHP"" SH 
                    ON TO_NVARCHAR(T0.""U_SHIPTRMS"") = TO_NVARCHAR(SH.""TrnspCode"")
                WHERE T0.""DocEntry"" = '{docEntry}'
                  AND T0.""DocNum"" = '{docNum}'
                  AND T0.""U_AMNDMNT"" = '{amendNo}'
                  AND T0.""LogInst"" = '{logInst}'";

            SAPbobsCOM.Recordset rs =
                (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            rs.DoQuery(sql);

            if (!rs.EoF)
            {
                SetEdit(oForm, "ETDOCTRY", rs.Fields.Item("DocEntry").Value);
                SetEdit(oForm, "ETDOCNUM", rs.Fields.Item("DocNum").Value);
                SetEdit(oForm, "ETSERIES", rs.Fields.Item("SeriesName").Value);

                SetEdit(oForm, "ETBRANCH", rs.Fields.Item("BranchName").Value);
                SetEdit(oForm, "ETMRSTAT", rs.Fields.Item("MarketingStatus").Value);
                SetEdit(oForm, "ETCMSTAT", rs.Fields.Item("CommercialStatus").Value);

                SetEdit(oForm, "ETCUSTMR", rs.Fields.Item("U_CARDCODE").Value);
                SetEdit(oForm, "ETCUSTNM", rs.Fields.Item("U_CARDNAME").Value);
                SetEdit(oForm, "ETBRNDCD", rs.Fields.Item("U_BRAND").Value);
                SetEdit(oForm, "ETBRNDNM", rs.Fields.Item("U_BRANDN").Value);
                SetEdit(oForm, "ETSCNO", rs.Fields.Item("U_SCNO").Value);
                SetEdit(oForm, "ETSCDESC", rs.Fields.Item("U_SCDESC").Value);
                SetEdit(oForm, "ETREFNCE", rs.Fields.Item("U_REFERANCE").Value);
                SetEdit(oForm, "ETSCVAL", rs.Fields.Item("U_SCVALUE").Value);
                SetEdit(oForm, "ETDOVAL", rs.Fields.Item("U_DRFTVALUE").Value);
                SetEdit(oForm, "ETB2BPER", rs.Fields.Item("U_BTOBLCPER").Value);
                SetEdit(oForm, "ETB2BVAL", rs.Fields.Item("U_BTOBLCVALUE").Value);
                SetEdit(oForm, "ETDSNBNK", rs.Fields.Item("DocSendToBank").Value);

                SetEditDate(oForm, "ETDOCDAT", rs.Fields.Item("U_DOCDATE").Value);
                SetEdit(oForm, "ETCBNKAC", rs.Fields.Item("U_CUSBKACT").Value);
                SetEdit(oForm, "ETCUSBNK", rs.Fields.Item("U_CUSBANK").Value);
                SetEdit(oForm, "ETOBNKAC", rs.Fields.Item("U_OWNBKACT").Value);
                SetEdit(oForm, "ETOWNBNK", rs.Fields.Item("U_OWNBNAME").Value);
                SetEdit(oForm, "ETCURR", rs.Fields.Item("U_CURRENCY").Value);
                SetEditDate(oForm, "ETISUDAT", rs.Fields.Item("U_ISSUEDATE").Value);
                SetEditDate(oForm, "ETSHPDAT", rs.Fields.Item("U_SHIPDATE").Value);
                SetEditDate(oForm, "ETEXPDAT", rs.Fields.Item("U_EXPDATE").Value);
                SetEdit(oForm, "ETTOLPER", rs.Fields.Item("U_TOLRPERC").Value);
                SetEdit(oForm, "ETAMNDNO", rs.Fields.Item("U_AMNDMNT").Value);
                SetEdit(oForm, "ETOPNAMT", rs.Fields.Item("U_OPBTBAMT").Value);

                SetEdit(oForm, "ETINTRMS", rs.Fields.Item("U_INCOTRMS").Value);
                SetEdit(oForm, "ETPYTRMS", rs.Fields.Item("PaymentTermsName").Value);
                SetEdit(oForm, "ETMDSHIP", rs.Fields.Item("ShipmentName").Value);
                SetEdit(oForm, "ETPORTLD", rs.Fields.Item("U_PRTOFLDG").Value);
                SetEdit(oForm, "ETCNDEST", rs.Fields.Item("U_CNTROFDT").Value);
                SetEdit(oForm, "ETPRTDIS", rs.Fields.Item("U_PORTOFDS").Value);
                SetEdit(oForm, "ETINSNCE", rs.Fields.Item("U_INSRNCE").Value);
                SetEdit(oForm, "ETSHPTOL", rs.Fields.Item("U_SHIPTLRN").Value);
                SetEdit(oForm, "ETHSCODE", rs.Fields.Item("U_HSCODE").Value);
                SetEdit(oForm, "ETDOCREQ", rs.Fields.Item("U_DOCRQRD").Value);
                SetEdit(oForm, "ETRMSCON", rs.Fields.Item("U_TRMSCOND").Value);
                SetEdit(oForm, "ETSHPADD", rs.Fields.Item("U_SHPADDRS").Value);
                SetEdit(oForm, "ETPRTSHP", rs.Fields.Item("PartialShipment").Value);
            }
        }

        private void LoadAmendmentOrderMatrix(SAPbouiCOM.Form oForm, string docEntry, string logInst)
        {
            SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXORDTL").Specific;
            mtx.Clear();

            string sql = $@"
                            SELECT
                                ""LineId"",
                                ""U_DRFTORDN"",
                                ""U_DRFTNRY"",
                                ""U_STYLECODE"",
                                ""U_STYLENTRY"",
                                ""U_DOCDATE"",
                                ""U_DUEDATE"",
                                ""U_TOTALQTY"",
                                ""U_TOTALVALUE""
                            FROM ""@AFIL_DR_SCM1""
                            WHERE ""DocEntry"" = '{docEntry}'
                              AND ""LogInst"" = '{logInst}'
                            ORDER BY ""LineId""";

            SAPbobsCOM.Recordset rs =
                (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            rs.DoQuery(sql);

            int row = 1;

            while (!rs.EoF)
            {
                mtx.AddRow();

                SetMatrixValue(mtx, "#", row, rs.Fields.Item("LineId").Value);
                SetMatrixValue(mtx, "CLDONO", row, rs.Fields.Item("U_DRFTORDN").Value);
                SetMatrixValue(mtx, "CLDONTRY", row, rs.Fields.Item("U_DRFTNRY").Value);
                SetMatrixValue(mtx, "CLSTYLNO", row, rs.Fields.Item("U_STYLECODE").Value);
                SetMatrixValue(mtx, "CLSLNTRY", row, rs.Fields.Item("U_STYLENTRY").Value);
                SetMatrixDate(mtx, "CLDOPSTD", row, rs.Fields.Item("U_DOCDATE").Value);
                SetMatrixDate(mtx, "CLDODLVD", row, rs.Fields.Item("U_DUEDATE").Value);
                SetMatrixValue(mtx, "CLTTLQTY", row, rs.Fields.Item("U_TOTALQTY").Value);
                SetMatrixValue(mtx, "CLTTLVAL", row, rs.Fields.Item("U_TOTALVALUE").Value);

                row++;
                rs.MoveNext();
            }

            mtx.AutoResizeColumns();
        }

        private void LoadAmendmentB2BMatrix(SAPbouiCOM.Form oForm, string docEntry, string logInst)
        {
            SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXB2BDL").Specific;
            mtx.Clear();

            string sql = $@"
                            SELECT
                                ""LineId"",
                                ""U_BTBLCDNO"",
                                ""U_LCNO"",
                                ""U_BTBLCNTRY"",
                                ""U_ISSUEDATE"",
                                ""U_LCVALUE""
                            FROM ""@AFIL_DR_SCM2""
                            WHERE ""DocEntry"" = '{docEntry}'
                              AND ""LogInst"" = '{logInst}'
                            ORDER BY ""LineId""";

            SAPbobsCOM.Recordset rs =
                (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            rs.DoQuery(sql);

            int row = 1;

            while (!rs.EoF)
            {
                mtx.AddRow();

                SetMatrixValue(mtx, "#", row, rs.Fields.Item("LineId").Value);
                SetMatrixValue(mtx, "CLB2BDNO", row, rs.Fields.Item("U_BTBLCDNO").Value);
                SetMatrixValue(mtx, "CLBLCNO", row, rs.Fields.Item("U_LCNO").Value);
                SetMatrixValue(mtx, "CLBCNTRY", row, rs.Fields.Item("U_BTBLCNTRY").Value);
                SetMatrixDate(mtx, "CLISUDAT", row, rs.Fields.Item("U_ISSUEDATE").Value);
                SetMatrixValue(mtx, "CLVALUE", row, rs.Fields.Item("U_LCVALUE").Value);

                row++;
                rs.MoveNext();
            }

            mtx.AutoResizeColumns();
        }

        private void LoadAmendmentAttachmentMatrix(SAPbouiCOM.Form oForm, string docEntry, string logInst)
        {
            SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXATTCH").Specific;
            mtx.Clear();

            string sql = $@"
                            SELECT
                                ""LineId"",
                                ""U_ATCHMENT"",
                                ""U_REMARKS""
                            FROM ""@AFIL_DR_SCM3""
                            WHERE ""DocEntry"" = '{docEntry}'
                              AND ""LogInst"" = '{logInst}'
                            ORDER BY ""LineId""";

            SAPbobsCOM.Recordset rs =
                (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            rs.DoQuery(sql);

            int row = 1;

            while (!rs.EoF)
            {
                mtx.AddRow();

                SetMatrixValue(mtx, "#", row, rs.Fields.Item("LineId").Value);
                SetMatrixValue(mtx, "CLATTACH", row, rs.Fields.Item("U_ATCHMENT").Value);
                SetMatrixValue(mtx, "CLREMARK", row, rs.Fields.Item("U_REMARKS").Value);

                row++;
                rs.MoveNext();
            }

            mtx.AutoResizeColumns();
        }

        private void ClearAmendmentForm(SAPbouiCOM.Form oForm)
        {
            string[] editIds =
            {
        "ETDOCTRY","ETDOCNUM","ETSERIES","ETBRANCH","ETMRSTAT","ETCMSTAT",
        "ETCUSTMR","ETCUSTNM","ETBRNDCD","ETBRNDNM","ETSCNO","ETSCDESC",
        "ETREFNCE","ETSCVAL","ETDOVAL","ETB2BPER","ETB2BVAL","ETDSNBNK",
        "ETDOCDAT","ETCBNKAC","ETCUSBNK","ETOBNKAC","ETOWNBNK","ETCURR",
        "ETISUDAT","ETSHPDAT","ETEXPDAT","ETTOLPER","ETAMNDNO","ETOPNAMT",
        "ETINTRMS","ETPYTRMS","ETMDSHIP","ETPORTLD","ETCNDEST","ETPRTDIS",
        "ETINSNCE","ETSHPTOL","ETHSCODE","ETDOCREQ","ETRMSCON","ETSHPADD",
        "ETPRTSHP"
    };

            foreach (string id in editIds)
            {
                try
                {
                    ((SAPbouiCOM.EditText)oForm.Items.Item(id).Specific).Value = "";
                }
                catch { }
            }

            try { ((SAPbouiCOM.Matrix)oForm.Items.Item("MTXORDTL").Specific).Clear(); } catch { }
            try { ((SAPbouiCOM.Matrix)oForm.Items.Item("MTXB2BDL").Specific).Clear(); } catch { }
            try { ((SAPbouiCOM.Matrix)oForm.Items.Item("MTXATTCH").Specific).Clear(); } catch { }
        }

        private void SetEdit(SAPbouiCOM.Form oForm, string itemId, object value)
        {
            try
            {
                ((SAPbouiCOM.EditText)oForm.Items.Item(itemId).Specific).Value =
                    value == null ? "" : value.ToString();
            }
            catch { }
        }

        private void SetEditDate(SAPbouiCOM.Form oForm, string itemId, object value)
        {
            try
            {
                if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                {
                    SetEdit(oForm, itemId, "");
                    return;
                }

                DateTime dt = Convert.ToDateTime(value);
                SetEdit(oForm, itemId, dt.ToString("yyyyMMdd"));
            }
            catch
            {
                SetEdit(oForm, itemId, value);
            }
        }

        private void SetMatrixValue(SAPbouiCOM.Matrix mtx, string colId, int row, object value)
        {
            try
            {
                ((SAPbouiCOM.EditText)mtx.Columns.Item(colId).Cells.Item(row).Specific).Value =
                    value == null ? "" : value.ToString();
            }
            catch { }
        }

        private void SetMatrixDate(SAPbouiCOM.Matrix mtx, string colId, int row, object value)
        {
            try
            {
                if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                {
                    SetMatrixValue(mtx, colId, row, "");
                    return;
                }

                DateTime dt = Convert.ToDateTime(value);
                SetMatrixValue(mtx, colId, row, dt.ToString("yyyyMMdd"));
            }
            catch
            {
                SetMatrixValue(mtx, colId, row, value);
            }
        }


        private void Form_DataUpdateAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            if (!pVal.ActionSuccess)
                return;

            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            CheckAndLoadAmendmentGrid(oForm);
        }

        private void CheckAndLoadAmendmentGrid(SAPbouiCOM.Form oForm)
        {
            try
            {
                string amendNoStr = ((SAPbouiCOM.EditText)oForm.Items.Item("ETAMNDNO").Specific).Value.Trim();

                int amendNo = 0;
                int.TryParse(amendNoStr, out amendNo);

                if (amendNo > 1)
                {
                    LoadAmendmentGrid(oForm);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Amendment check error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void LoadAmendmentGrid(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                string docEntry = ((SAPbouiCOM.EditText)oForm.Items.Item("ETDOCTRY").Specific).Value.Trim();
                string docNum = ((SAPbouiCOM.EditText)oForm.Items.Item("ETDOCNUM").Specific).Value.Trim();

                if (string.IsNullOrWhiteSpace(docEntry) || string.IsNullOrWhiteSpace(docNum))
                {
                    return;
                }

                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("GRDAMEND").Specific;

                SAPbouiCOM.DataTable oDT;
                try
                {
                    oDT = oForm.DataSources.DataTables.Item("DT_AMEN");
                }
                catch
                {
                    oDT = oForm.DataSources.DataTables.Add("DT_AMEN");
                }

                // Clear grid/data table first
                oGrid.DataTable = null;
                oDT.Clear();

                string query = $@"
                                SELECT 
                                    ROW_NUMBER() OVER (ORDER BY ""DocEntry"") AS ""#"",
                                    ""DocEntry"",
                                    ""DocNum"",
                                    ""Creator"",
                                    ""CreateDate"",
                                    ""UpdateDate"",
                                    ""LogInst"",
                                    ""U_AMNDMNT""
                                FROM
                                (
                                    SELECT 
                                        ""DocEntry"",
                                        ""DocNum"",
                                        ""Creator"",
                                        ""CreateDate"",
                                        ""UpdateDate"",
                                        MAX(""LogInst"") AS ""LogInst"",
                                        ""U_AMNDMNT""
                                    FROM ""@AFIL_DH_OSCM""
                                    WHERE ""DocEntry"" = {docEntry}
                                      AND ""DocNum"" = {docNum}
                                    GROUP BY 
                                        ""DocEntry"",
                                        ""DocNum"",
                                        ""Creator"",
                                        ""CreateDate"",
                                        ""UpdateDate"",
                                        ""U_AMNDMNT""
                                ) T
                                ORDER BY ""DocEntry"";
        ";

                oDT.ExecuteQuery(query);

                oGrid.DataTable = oDT;

                for (int i = 0; i < oGrid.Columns.Count; i++)
                {
                    oGrid.Columns.Item(i).Editable = false;
                }

                oGrid.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error loading amendment grid: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private void BTNAMEND_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

                SAPbouiCOM.ComboBox cbDSNBNK = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBDSNBNK").Specific;
                SAPbouiCOM.ComboBox cbMRSTAT = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBMRSTAT").Specific;
                SAPbouiCOM.ComboBox cbCMSTAT = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBCMSTAT").Specific;
                SAPbouiCOM.EditText etAmndNo = (SAPbouiCOM.EditText)oForm.Items.Item("ETAMNDNO").Specific;

                string sendToBank = "";

                if (cbDSNBNK.Selected != null)
                    sendToBank = cbDSNBNK.Selected.Value.Trim();

                // If not sent to bank
                if (sendToBank == "N")
                {
                    Application.SBO_Application.MessageBox("This Document is not sent to Bank Yet.");
                    return;
                }

                // If sent to bank
                if (sendToBank == "Y")
                {
                    cbMRSTAT.Select("D", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    cbCMSTAT.Select("D", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    cbDSNBNK.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);

                    int amendNo = 0;

                    if (!string.IsNullOrWhiteSpace(etAmndNo.Value))
                        amendNo = Convert.ToInt32(etAmndNo.Value);

                    amendNo++;

                    etAmndNo.Value = amendNo.ToString();

                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "BTNAMEND Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
        }

        private void ControlSendToBank(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.ComboBox cbMR =
                    (SAPbouiCOM.ComboBox)oForm.Items.Item("CBMRSTAT").Specific;

                SAPbouiCOM.ComboBox cbCM =
                    (SAPbouiCOM.ComboBox)oForm.Items.Item("CBCMSTAT").Specific;

                string mr = cbMR.Selected == null ? "" : cbMR.Selected.Value.Trim();
                string cm = cbCM.Selected == null ? "" : cbCM.Selected.Value.Trim();

                if (mr == "C" && cm == "C")
                {
                    oForm.Items.Item("CBDSNBNK").Enabled = true;
                }
                else
                {
                    oForm.Items.Item("CBDSNBNK").Enabled = false;
                }
            }
            catch { }
        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            SAPbouiCOM.Form oForm = null;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

                string scNo = ((SAPbouiCOM.EditText)oForm.Items.Item("ETSCNO").Specific).Value.Trim();

                if (string.IsNullOrWhiteSpace(scNo))
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Sales Contract No is empty.",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning
                    );
                    return;
                }

                // SC saved or not check
                bool scExists = IsSCNoExistsInOSCM(scNo);
                DisableFieldsIfSCExists(oForm, scExists);

                // SC used in Draft Order or not check
                bool scOQUT = IsSCNoExistsInOQUT(scNo);

                if (!scOQUT)
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Sales Contract is not used in Draft Order.",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning
                    );
                    return;
                }

                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix =
                    (SAPbouiCOM.Matrix)oForm.Items.Item("MTXORDTL").Specific;

                SAPbouiCOM.DBDataSource oDBDS =
                    oForm.DataSources.DBDataSources.Item("@FIL_DR_SCM1");

                oDBDS.Clear();
                oMatrix.LoadFromDataSource();

                string safeSCNo = scNo.Replace("'", "''");

                string qStr = @"
                                SELECT 
                                    ROW_NUMBER() OVER (ORDER BY A.""DocEntry"") AS ""LineId"",
                                    A.""DocNum"",
                                    A.""DocEntry"",
                                    A.""U_STYLECODE"",
                                    A.""U_STYLENTRY"",
                                    A.""DocDate"",
                                    A.""DocDueDate"",
                                    SUM(B.""Quantity"") AS ""Quantity"",
                                    CASE 
                                        WHEN A.""DocCur"" = 'BDT' THEN A.""DocTotal""
                                        ELSE A.""DocTotalFC""
                                    END AS ""TotalValue""
                                FROM OQUT AS A
                                INNER JOIN QUT1 AS B
                                    ON A.""DocEntry"" = B.""DocEntry""
                                WHERE A.""U_SCNO"" = '" + safeSCNo + @"'
                                GROUP BY 
                                    A.""DocNum"",
                                    A.""DocEntry"",
                                    A.""U_STYLECODE"",
                                    A.""U_STYLENTRY"",
                                    A.""DocDate"",
                                    A.""DocDueDate"",
                                    A.""DocCur"",
                                    A.""DocTotal"",
                                    A.""DocTotalFC""
                                ORDER BY A.""DocEntry""";

                SAPbobsCOM.Recordset rs =
                    (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                rs.DoQuery(qStr);

                int row = 0;

                while (!rs.EoF)
                {
                    oDBDS.InsertRecord(row);

                    oDBDS.SetValue("LineId", row, (row + 1).ToString());
                    oDBDS.SetValue("U_DRFTORDN", row, rs.Fields.Item("DocNum").Value.ToString());
                    oDBDS.SetValue("U_DRFTNRY", row, rs.Fields.Item("DocEntry").Value.ToString());
                    oDBDS.SetValue("U_STYLECODE", row, rs.Fields.Item("U_STYLECODE").Value.ToString());
                    oDBDS.SetValue("U_STYLENTRY", row, rs.Fields.Item("U_STYLENTRY").Value.ToString());

                    DateTime docDate = Convert.ToDateTime(rs.Fields.Item("DocDate").Value);
                    DateTime dueDate = Convert.ToDateTime(rs.Fields.Item("DocDueDate").Value);

                    oDBDS.SetValue("U_DOCDATE", row, docDate.ToString("yyyyMMdd"));
                    oDBDS.SetValue("U_DUEDATE", row, dueDate.ToString("yyyyMMdd"));

                    oDBDS.SetValue("U_TOTALQTY", row, rs.Fields.Item("Quantity").Value.ToString());
                    oDBDS.SetValue("U_TOTALVALUE", row, rs.Fields.Item("TotalValue").Value.ToString());

                    row++;
                    rs.MoveNext();
                }

                oMatrix.LoadFromDataSource();
                SetDraftOrderTotalValue(oForm);
                oMatrix.AutoResizeColumns();
                ControlSendToBank(oForm);
                CheckAndLoadAmendmentGrid(oForm);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Form_DataLoadAfter Error: " + ex.Message,
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

        private void SetDraftOrderTotalValue(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix =
                    (SAPbouiCOM.Matrix)oForm.Items.Item("MTXORDTL").Specific;

                double totalValue = 0;

                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    string val = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLTTLVAL").Cells.Item(i).Specific).Value.Trim();

                    if (!string.IsNullOrEmpty(val))
                    {
                        double rowValue = 0;
                        double.TryParse(val, out rowValue);
                        totalValue += rowValue;
                    }
                }

                ((SAPbouiCOM.EditText)oForm.Items.Item("ETDOVAL").Specific).Value =
                    totalValue.ToString("0.00");
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "SetDraftOrderTotalValue Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
        }

        private bool IsSCNoExistsInOQUT(string scNo)
        {
            string safeSCNo = scNo.Replace("'", "''");

            string qStr = $@"
                            SELECT COUNT(*) AS ""Cnt""
                            FROM ""OQUT""
                            WHERE ""U_SCNO"" = '{safeSCNo}'";

            SAPbobsCOM.Recordset rs =
                (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            rs.DoQuery(qStr);
            return Convert.ToInt32(rs.Fields.Item("Cnt").Value) > 0;
        }

        private bool IsSCNoExistsInOSCM(string scNo)
        {
            string safeSCNo = scNo.Replace("'", "''");

            string qStr = $@"
                            SELECT COUNT(*) AS ""Cnt""
                            FROM ""@FIL_DH_OSCM""
                            WHERE ""U_SCNO"" = '{safeSCNo}'";

            SAPbobsCOM.Recordset rs =
                (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            rs.DoQuery(qStr);
            return Convert.ToInt32(rs.Fields.Item("Cnt").Value) > 0;
        }

        private void DisableFieldsIfSCExists(SAPbouiCOM.Form oForm, bool exists)
        {
            string[] itemIds = {"ETSCNO","ETCUSTMR","ETBRNDCD","ETISUDAT"};

            foreach (string itemId in itemIds)
            {
                try
                {
                    oForm.Items.Item(itemId).Enabled = !exists;
                }
                catch { }
            }
        }

        private void ADDButton_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.DBDataSource oDBH = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@FIL_DH_OSCM");
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                SAPbouiCOM.ComboBox ocmb = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBSERIES").Specific;
                Global.GFunc.LoadComboBoxSeries(ocmb, "FIL_D_OSCM");  //Object Type
                string ocmbvalue = ocmb.Selected.Value;
                long docno = oForm.BusinessObject.GetNextSerialNumber(ocmbvalue, "FIL_D_OSCM");

                oDBH.SetValue("DocNum", 0, docno.ToString()); // only set the value in string.

                //Date
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETDOCDAT").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
                //Amendment No
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETAMNDNO").Specific).Value = "1";

                // Branch combo
                string sqlQuerybpl = @"SELECT ""BPLId"", ""BPLName"" FROM ""OBPL""";
                SAPbouiCOM.ComboBox CBCMPANY = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBBRANCH").Specific;
                Global.GFunc.setComboBoxValue(CBCMPANY, sqlQuerybpl);
                CBCMPANY.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

                //Payment Terms
                string payTerms = @"SELECT ""GroupNum"", ""PymntGroup"" FROM ""OCTG""";
                SAPbouiCOM.ComboBox CBPYTRMS = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBPYTRMS").Specific;
                Global.GFunc.setComboBoxValue(CBPYTRMS, payTerms);

                //Shipping Type
                string shipType = @"SELECT ""TrnspCode"", ""TrnspName"" FROM ""OSHP""";
                SAPbouiCOM.ComboBox CBMDSHIP = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBMDSHIP").Specific;
                Global.GFunc.setComboBoxValue(CBMDSHIP, shipType);
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
                string branch = oForm.DataSources.DBDataSources.Item("@FIL_DH_OSCM").GetValue("U_BRANCH", 0).Trim();
                string customer = oForm.DataSources.DBDataSources.Item("@FIL_DH_OSCM").GetValue("U_CARDCODE", 0).Trim();
                string brand = oForm.DataSources.DBDataSources.Item("@FIL_DH_OSCM").GetValue("U_BRAND", 0).Trim();
                string scNo = oForm.DataSources.DBDataSources.Item("@FIL_DH_OSCM").GetValue("U_SCNO", 0).Trim();
                string issueDate = oForm.DataSources.DBDataSources.Item("@FIL_DH_OSCM").GetValue("U_ISSUEDATE", 0).Trim();
                

                if (branch == "")
                {
                    Global.GFunc.ShowError("Select Branch");
                    oForm.ActiveItem = "CBBRANCH";
                    return BubbleEvent = false;
                }
                if (customer == "")
                {
                    Global.GFunc.ShowError("Enter Customer Code");
                    oForm.ActiveItem = "ETCUSTMR";
                    return BubbleEvent = false;
                }
                else if (brand == "")
                {
                    Global.GFunc.ShowError("Enter Brand Code");
                    oForm.ActiveItem = "ETBRNDCD";
                    return BubbleEvent = false;
                }
                else if (scNo == "")
                {
                    Global.GFunc.ShowError("Enter Sales Contract No ");
                    oForm.ActiveItem = "ETSCNO";
                    return BubbleEvent = false;
                }
                else if (issueDate == "")
                {
                    Global.GFunc.ShowError("Enter Issue Date");
                    oForm.ActiveItem = "ETISUDAT";
                    return BubbleEvent = false;
                }

                //Prevent New Empty Free Line
                PreventEmptyLastRow(oForm, "@FIL_DR_SCM1", MTXORDTL, "U_DRFTORDN");
                PreventEmptyLastRow(oForm, "@FIL_DR_SCM2", MTXB2BDL, "U_BTBLCDNO");
                PreventEmptyLastRow(oForm, "@FIL_DR_SCM3", MTXATTCH, "U_ATCHMENT");
                
            }

            return BubbleEvent;
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

        private void DELBTN_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.Matrix MTXATTCH = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXATTCH").Specific;

            for (int i = 1; i <= MTXATTCH.RowCount; i++)
            {
                if (MTXATTCH.IsRowSelected(i))
                {
                    string filePath = ((SAPbouiCOM.EditText)MTXATTCH.Columns.Item("CLATTACH").Cells.Item(i).Specific).Value;
                    if (!string.IsNullOrEmpty(filePath) && System.IO.File.Exists(filePath))
                    {
                        System.Diagnostics.Process.Start(filePath);
                    }
                    else
                    {
                        Application.SBO_Application.MessageBox("File does not exist or path is empty.");
                    }
                    break;
                }
            }

        }

        private void DISPBTN_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.Matrix MTXATTCH = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXATTCH").Specific;

            for (int i = 1; i <= MTXATTCH.RowCount; i++)
            {
                if (MTXATTCH.IsRowSelected(i))
                {
                    string filePath = ((SAPbouiCOM.EditText)MTXATTCH.Columns.Item("CLATTACH").Cells.Item(i).Specific).Value;
                    if (!string.IsNullOrEmpty(filePath) && System.IO.File.Exists(filePath))
                    {
                        System.Diagnostics.Process.Start(filePath);
                    }
                    else
                    {
                        Application.SBO_Application.MessageBox("File does not exist or path is empty.");
                    }
                    break;
                }
            };

        }

        private void BRWSBTN_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.DBDataSource DBDataSourceLine = oForm.DataSources.DBDataSources.Item("@FIL_DR_SCM3");
            SAPbouiCOM.Matrix MTXATTCH = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXATTCH").Specific;

            string filePath = FileDialogHelper.ShowFileDialog();
            if (!string.IsNullOrEmpty(filePath))
            {
                int lastRow = MTXATTCH.VisualRowCount;
                bool needNewRow = (lastRow == 0) ||
                                  !string.IsNullOrEmpty(((SAPbouiCOM.EditText)MTXATTCH.Columns.Item("CLATTACH").Cells.Item(lastRow).Specific).Value);
                if (needNewRow)
                {
                    Global.GFunc.SetNewLine(MTXATTCH, DBDataSourceLine, 1, "");
                    lastRow = MTXATTCH.VisualRowCount;
                }

                ((SAPbouiCOM.EditText)MTXATTCH.Columns.Item("CLATTACH").Cells.Item(lastRow).Specific).Value = filePath;
                MTXATTCH.FlushToDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
            }

        }

        private void CBCMSTAT_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.ComboBox oCmb = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBCMSTAT").Specific;

                string selectedValue = "";
                if (oCmb.Selected != null)
                    selectedValue = oCmb.Selected.Value.Trim();

                if (selectedValue == "C")
                {
                    int result = Application.SBO_Application.MessageBox(
                        "Are you sure want to change?",
                        1,
                        "Yes",
                        "No"
                    );

                    if (result != 1)
                    {
                        oCmb.Select("D", SAPbouiCOM.BoSearchKey.psk_ByValue);

                    }
                }
                ControlSendToBank(oForm);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error in CM Status selection: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }

        }

        private void CBMRSTAT_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.ComboBox oCmb = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBMRSTAT").Specific;

                string selectedValue = "";
                if (oCmb.Selected != null)
                    selectedValue = oCmb.Selected.Value.Trim();

                if (selectedValue == "C")
                {
                    int result = Application.SBO_Application.MessageBox(
                        "Are you sure want to change?",
                        1,
                        "Yes",
                        "No"
                    );

                    if (result != 1)
                    {
                        oCmb.Select("D", SAPbouiCOM.BoSearchKey.psk_ByValue);
                       
                    }                    
                }
                ControlSendToBank(oForm);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error in MR Status selection: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }

        }

        private void CBDSNBNK_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.ComboBox oCmb = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBDSNBNK").Specific;

                string selectedValue = "";
                if (oCmb.Selected != null)
                    selectedValue = oCmb.Selected.Value.Trim();

                if (selectedValue == "Y")
                {
                    int result = Application.SBO_Application.MessageBox(
                        "Are you sure want to change?",
                        1,
                        "Yes",
                        "No"
                    );

                    if (result != 1)
                    {
                        oCmb.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }

                    
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error in Doc Send to Bank selection: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
        }

        private void ETSCNO_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    string code = ((SAPbouiCOM.EditText)oForm.Items.Item("ETSCNO").Specific).Value.Trim();
                    string UCode = Global.GFunc.ToUpperCase(code);
                    ((SAPbouiCOM.EditText)oForm.Items.Item("ETSCNO").Specific).Value = UCode;
                    if (!string.IsNullOrEmpty(UCode))
                    {
                        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string query = $@"SELECT 1 FROM ""@FIL_DH_OSCM"" WHERE ""U_SCNO"" = '{UCode.Replace("'", "''")}'";
                        oRS.DoQuery(query);
                        if (!oRS.EoF)
                        {
                            Application.SBO_Application.StatusBar.SetText("Code already exists!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            ((SAPbouiCOM.EditText)oForm.Items.Item("ETSCNO").Specific).Value = "";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Error: Sales Contract Code  " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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

        private void ETCNDEST_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("Name", 0).ToString().Trim();
            SAPbouiCOM.EditText ETCNTRY = (SAPbouiCOM.EditText)oForm.Items.Item("ETCNDEST").Specific;
            ETCNTRY.Value = Code;

        }

        private void ETOBNKAC_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Account = dt.GetValue("Account", 0).ToString().Trim();
            string Bank = dt.GetValue("BankCode", 0).ToString().Trim();

            SAPbouiCOM.EditText ETBAC = (SAPbouiCOM.EditText)oForm.Items.Item("ETOBNKAC").Specific;
            ETBAC.Value = Account;
            SAPbouiCOM.EditText ETBCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETOWNBNK").Specific;
            ETBCD.Value = Bank;

        }

        private void ETCBNKAC_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Account = dt.GetValue("Account", 0).ToString().Trim();
            string Bank = dt.GetValue("BankCode", 0).ToString().Trim();

            SAPbouiCOM.EditText ETBAC = (SAPbouiCOM.EditText)oForm.Items.Item("ETCBNKAC").Specific;
            ETBAC.Value = Account;
            SAPbouiCOM.EditText ETBCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETCUSBNK").Specific;
            ETBCD.Value = Bank;

        }

        private void ETCBNKAC_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

                string customerCode = ((SAPbouiCOM.EditText)oForm.Items.Item("ETCUSTMR").Specific).Value.Trim();

                if (string.IsNullOrEmpty(customerCode))
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Please select Customer first.",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Error
                    );
                    BubbleEvent = false;
                    return;
                }

                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_OCRB")
                {
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(cflUID);

                    SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                    SAPbouiCOM.Condition oCon = oCons.Add();

                    oCon.Alias = "CardCode";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = customerCode;

                    oCFL.SetConditions(oCons);
                }
            }
            catch (Exception ex)
            {   
                Application.SBO_Application.StatusBar.SetText(
                    "Error filtering Bank Account CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }
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

        private void ETCUSTMR_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_OCRD")
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(cflUID);
                    SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                    SAPbouiCOM.Condition oCon1 = oCons.Add();
                    oCon1.Alias = "CardType";
                    oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon1.CondVal = "C";
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

        private void ETCUSTMR_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("CardCode", 0).ToString().Trim();
            string Name = dt.GetValue("CardName", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETCUSTMR").Specific;
            ETCD.Value = Code;

            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETCUSTNM").Specific;
            ETNM.Value = Name;


            SAPbouiCOM.Item CBSHPADD = oForm.Items.Item("CBSHPADD");
            CBSHPADD.Enabled = true;

            //Shipping Type
            string address = $@"SELECT ""Address"", ""Street"" FROM ""CRD1"" where ""CardCode""='{Code}'";
            SAPbouiCOM.ComboBox CBADD= (SAPbouiCOM.ComboBox)oForm.Items.Item("CBSHPADD").Specific;
            Global.GFunc.setComboBoxValue(CBADD, address);
        }

        
    }
}
