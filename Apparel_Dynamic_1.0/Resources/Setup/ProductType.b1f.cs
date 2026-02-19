using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apparel_Dynamic_1._0.Resources.Setup
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Setup.ProductType", "Resources/Setup/ProductType.b1f")]
    class ProductType : UserFormBase
    {
        public ProductType()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("STCODE").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("STNAME").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("ETCODE").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("ETNAME").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("CKACTIVE").Specific));
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
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.CheckBox CheckBox0;
    }
}
