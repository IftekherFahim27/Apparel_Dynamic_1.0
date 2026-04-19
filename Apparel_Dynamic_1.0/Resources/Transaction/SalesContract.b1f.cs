using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
        private SAPbouiCOM.ComboBox CBBRANCH, CBMRSTAT, CBCMSTAT, CBINTRMS,
                                    CBPYTRMS, CBMDSHIP, CBPRTSHP, CBSERIES;

        // -------- Edit Text --------
        private SAPbouiCOM.EditText ETCUSTMR, ETCUSTNM, ETBRNDCD, ETBRNDNM, ETSCNO,
                                    ETSCDESC, ETREFNCE, ETSCVAL, ETDOVAL, ETB2BPER,
                                    ETB2BVAL, ETDSNBNK, ETDOCNUM, ETDOCDAT, ETCBNKAC,
                                    ETCUSBNK, ETOBNKAC, ETOWNBNK, ETCURR, ETISUDAT,
                                    ETSHPDAT, ETEXPDAT, ETTOLPER, ETAMNDNO, ETOPNAMT,
                                    ETPORTLD, ETCNDEST, ETPRTDIS, ETINSNCE, ETSHPTOL,
                                    ETHSCODE, ETDOCREQ, ETRMSCON, ETSHPADD, ETDOCTRY;

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
            // -------- Static Text --------
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

            // -------- Folder --------
            this.FOLORDTL = ((SAPbouiCOM.Folder)(this.GetItem("FOLORDTL").Specific));
            this.FOLAMEND = ((SAPbouiCOM.Folder)(this.GetItem("FOLAMEND").Specific));
            this.FOLB2BDL = ((SAPbouiCOM.Folder)(this.GetItem("FOLB2BDL").Specific));
            this.FOLATTCH = ((SAPbouiCOM.Folder)(this.GetItem("FOLATTCH").Specific));
            this.FOLCDTLS = ((SAPbouiCOM.Folder)(this.GetItem("FOLCDTLS").Specific));

            // -------- Button --------
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.BTNAMEND = ((SAPbouiCOM.Button)(this.GetItem("BTNAMEND").Specific));
            this.BRWSBTN = ((SAPbouiCOM.Button)(this.GetItem("BRWSBTN").Specific));
            this.DISPBTN = ((SAPbouiCOM.Button)(this.GetItem("DISPBTN").Specific));
            this.DELBTN = ((SAPbouiCOM.Button)(this.GetItem("DELBTN").Specific));

            // -------- ComboBox --------
            this.CBBRANCH = ((SAPbouiCOM.ComboBox)(this.GetItem("CBBRANCH").Specific));
            this.CBMRSTAT = ((SAPbouiCOM.ComboBox)(this.GetItem("CBMRSTAT").Specific));
            this.CBCMSTAT = ((SAPbouiCOM.ComboBox)(this.GetItem("CBCMSTAT").Specific));
            this.CBINTRMS = ((SAPbouiCOM.ComboBox)(this.GetItem("CBINTRMS").Specific));
            this.CBPYTRMS = ((SAPbouiCOM.ComboBox)(this.GetItem("CBPYTRMS").Specific));
            this.CBMDSHIP = ((SAPbouiCOM.ComboBox)(this.GetItem("CBMDSHIP").Specific));
            this.CBPRTSHP = ((SAPbouiCOM.ComboBox)(this.GetItem("CBPRTSHP").Specific));
            this.CBSERIES = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSERIES").Specific));

            // -------- Edit Text --------
            this.ETCUSTMR = ((SAPbouiCOM.EditText)(this.GetItem("ETCUSTMR").Specific));
            this.ETCUSTNM = ((SAPbouiCOM.EditText)(this.GetItem("ETCUSTNM").Specific));
            this.ETBRNDCD = ((SAPbouiCOM.EditText)(this.GetItem("ETBRNDCD").Specific));
            this.ETBRNDNM = ((SAPbouiCOM.EditText)(this.GetItem("ETBRNDNM").Specific));
            this.ETSCNO   = ((SAPbouiCOM.EditText)(this.GetItem("ETSCNO").Specific));
            this.ETSCDESC = ((SAPbouiCOM.EditText)(this.GetItem("ETSCDESC").Specific));
            this.ETREFNCE = ((SAPbouiCOM.EditText)(this.GetItem("ETREFNCE").Specific));
            this.ETSCVAL  = ((SAPbouiCOM.EditText)(this.GetItem("ETSCVAL").Specific));
            this.ETDOVAL  = ((SAPbouiCOM.EditText)(this.GetItem("ETDOVAL").Specific));
            this.ETB2BPER = ((SAPbouiCOM.EditText)(this.GetItem("ETB2BPER").Specific));
            this.ETB2BVAL = ((SAPbouiCOM.EditText)(this.GetItem("ETB2BVAL").Specific));
            this.ETDSNBNK = ((SAPbouiCOM.EditText)(this.GetItem("ETDSNBNK").Specific));
            this.ETDOCNUM = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCNUM").Specific));
            this.ETDOCDAT = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCDAT").Specific));
            this.ETCBNKAC = ((SAPbouiCOM.EditText)(this.GetItem("ETCBNKAC").Specific));
            this.ETCUSBNK = ((SAPbouiCOM.EditText)(this.GetItem("ETCUSBNK").Specific));
            this.ETOBNKAC = ((SAPbouiCOM.EditText)(this.GetItem("ETOBNKAC").Specific));
            this.ETOWNBNK = ((SAPbouiCOM.EditText)(this.GetItem("ETOWNBNK").Specific));
            this.ETCURR   = ((SAPbouiCOM.EditText)(this.GetItem("ETCURR").Specific));
            this.ETISUDAT = ((SAPbouiCOM.EditText)(this.GetItem("ETISUDAT").Specific));
            this.ETSHPDAT = ((SAPbouiCOM.EditText)(this.GetItem("ETSHPDAT").Specific));
            this.ETEXPDAT = ((SAPbouiCOM.EditText)(this.GetItem("ETEXPDAT").Specific));
            this.ETTOLPER = ((SAPbouiCOM.EditText)(this.GetItem("ETTOLPER").Specific));
            this.ETAMNDNO = ((SAPbouiCOM.EditText)(this.GetItem("ETAMNDNO").Specific));
            this.ETOPNAMT = ((SAPbouiCOM.EditText)(this.GetItem("ETOPNAMT").Specific));
            this.ETPORTLD = ((SAPbouiCOM.EditText)(this.GetItem("ETPORTLD").Specific));
            this.ETCNDEST = ((SAPbouiCOM.EditText)(this.GetItem("ETCNDEST").Specific));
            this.ETPRTDIS = ((SAPbouiCOM.EditText)(this.GetItem("ETPRTDIS").Specific));
            this.ETINSNCE = ((SAPbouiCOM.EditText)(this.GetItem("ETINSNCE").Specific));
            this.ETSHPTOL = ((SAPbouiCOM.EditText)(this.GetItem("ETSHPTOL").Specific));
            this.ETHSCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETHSCODE").Specific));
            this.ETDOCREQ = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCREQ").Specific));
            this.ETRMSCON = ((SAPbouiCOM.EditText)(this.GetItem("ETRMSCON").Specific));
            this.ETSHPADD = ((SAPbouiCOM.EditText)(this.GetItem("ETSHPADD").Specific));
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));

            // -------- Matrix --------
            this.MTXORDTL = ((SAPbouiCOM.Matrix)(this.GetItem("MTXORDTL").Specific));
            this.MTXATTCH = ((SAPbouiCOM.Matrix)(this.GetItem("MTXATTCH").Specific));
            this.MTXB2BDL = ((SAPbouiCOM.Matrix)(this.GetItem("MTXB2BDL").Specific));

            // -------- Grid --------
            this.GRDAMEND = ((SAPbouiCOM.Grid)(this.GetItem("GRDAMEND").Specific));

           
            this.OnCustomInitialize();

        }

        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {

        }

    }
}
