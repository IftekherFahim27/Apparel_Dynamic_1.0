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

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("STEFRMDT").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("STEFTODT").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("ETEFRMDT").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("ETEFTODT").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCNUM").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSERIES").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("MTXLEDTM").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCNUM").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {

        }

        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
    }
}
