using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apparel_Dynamic_1._0.Resources.Setup
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Setup.ProductGroup", "Resources/Setup/ProductGroup.b1f")]
    class ProductGroup : UserFormBase
    {
        public ProductGroup()
        {
        }


        private SAPbouiCOM.StaticText STCODE, STNAME, STPDTYPE, STPDLINE;
        private SAPbouiCOM.EditText ETCODE, ETNAME, ETDOCTRY, ETPDTYPE, ETPDLINE;
        private SAPbouiCOM.Button ADDButton, CancelButton;
        private SAPbouiCOM.CheckBox CKACTIVE;
        public override void OnInitializeComponent()
        {
            this.STCODE = ((SAPbouiCOM.StaticText)(this.GetItem("STCODE").Specific));
            this.STNAME = ((SAPbouiCOM.StaticText)(this.GetItem("STNAME").Specific));
            this.ETCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETCODE").Specific));
            this.ETNAME = ((SAPbouiCOM.EditText)(this.GetItem("ETNAME").Specific));
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.CKACTIVE = ((SAPbouiCOM.CheckBox)(this.GetItem("CKACTIVE").Specific));
            this.STPDTYPE = ((SAPbouiCOM.StaticText)(this.GetItem("STPDTYPE").Specific));
            this.STPDLINE = ((SAPbouiCOM.StaticText)(this.GetItem("STPDLINE").Specific));
            this.ETPDTYPE = ((SAPbouiCOM.EditText)(this.GetItem("ETPDTYPE").Specific));
            this.ETPDLINE = ((SAPbouiCOM.EditText)(this.GetItem("ETPDLINE").Specific));
            this.OnCustomInitialize();

        }

        public override void OnInitializeFormEvents()
        {
        }



        private void OnCustomInitialize()
        {

        }


    }
}
