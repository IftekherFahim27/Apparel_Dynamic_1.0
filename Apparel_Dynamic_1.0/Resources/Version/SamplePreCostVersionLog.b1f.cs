using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apparel_Dynamic_1._0.Resources.Version
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Version.SamplePreCostVersionLog", "Resources/Version/SamplePreCostVersionLog.b1f")]
    class SamplePreCostVersionLog : UserFormBase
    {
        public SamplePreCostVersionLog()
        {
        }


        private SAPbouiCOM.StaticText STSMPLCD, STCURR, STTCNAMT, STNO, STDOCNUM, STDATE, STVERSON, STBUYER;
        private SAPbouiCOM.EditText ETCURR, ETTCNAMT, ETNO, ETBYRNM, ETSMPLNM, ETDOCNUM, ETDATE, ETVERSON, ETBUYER, ETDOCTRY, ETSMPLCD;
        private SAPbouiCOM.Folder FOLCMPNT, FOLOTCST;
        private SAPbouiCOM.Matrix MTXCMPNT, MTXOTCST;
        private SAPbouiCOM.Button ADDButton, CancelButton;

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
            this.ETCURR = ((SAPbouiCOM.EditText)(this.GetItem("ETCURR").Specific));
            this.ETTCNAMT = ((SAPbouiCOM.EditText)(this.GetItem("ETTCNAMT").Specific));
            this.ETDOCNUM = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCNUM").Specific));
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            this.ETDATE = ((SAPbouiCOM.EditText)(this.GetItem("ETDATE").Specific));
            this.ETVERSON = ((SAPbouiCOM.EditText)(this.GetItem("ETVERSON").Specific));
            this.ETBUYER = ((SAPbouiCOM.EditText)(this.GetItem("ETBUYER").Specific));
            this.FOLCMPNT = ((SAPbouiCOM.Folder)(this.GetItem("FOLCMPNT").Specific));
            this.FOLOTCST = ((SAPbouiCOM.Folder)(this.GetItem("FOLOTCST").Specific));
            this.MTXCMPNT = ((SAPbouiCOM.Matrix)(this.GetItem("MTXCMPNT").Specific));
            this.MTXOTCST = ((SAPbouiCOM.Matrix)(this.GetItem("MTXOTCST").Specific));
            this.ETSMPLNM = ((SAPbouiCOM.EditText)(this.GetItem("ETSMPLNM").Specific));
            this.ETBYRNM = ((SAPbouiCOM.EditText)(this.GetItem("ETBYRNM").Specific));
            this.ETSMPLCD = ((SAPbouiCOM.EditText)(this.GetItem("ETSMPLCD").Specific));
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("ETSERIES").Specific));
            this.OnCustomInitialize();

        }

        public override void OnInitializeFormEvents()
        {
        }


        private void OnCustomInitialize()
        {

        }

        private SAPbouiCOM.EditText EditText0;
    }
}
