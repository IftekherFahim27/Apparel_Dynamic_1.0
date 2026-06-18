using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apparel_Dynamic_1._0.Resources.Setup
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Setup.LeadTime", "Resources/Setup/LeadTime.b1f")]
    class LeadTime : UserFormBase
    {
        public LeadTime()
        {
        }

        private SAPbouiCOM.StaticText STEFRMDT, STEFTODT, STDOCNUM;
        private SAPbouiCOM.EditText ETEFRMDT, ETEFTODT, ETDOCTRY, ETDOCNUM;
      
        private SAPbouiCOM.ComboBox CBSERIES;
        private SAPbouiCOM.Matrix MTXLEDTM;

        private SAPbouiCOM.Button ADDButton,CancelButton;


        public override void OnInitializeComponent()
        {
            this.STEFRMDT = ((SAPbouiCOM.StaticText)(this.GetItem("STEFRMDT").Specific));
            this.STEFTODT = ((SAPbouiCOM.StaticText)(this.GetItem("STEFTODT").Specific));
            this.ETEFRMDT = ((SAPbouiCOM.EditText)(this.GetItem("ETEFRMDT").Specific));
            this.ETEFTODT = ((SAPbouiCOM.EditText)(this.GetItem("ETEFTODT").Specific));
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            this.ETDOCNUM = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCNUM").Specific));
            this.CBSERIES = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSERIES").Specific));
            this.MTXLEDTM = ((SAPbouiCOM.Matrix)(this.GetItem("MTXLEDTM").Specific));
            this.STDOCNUM = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCNUM").Specific));
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
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
