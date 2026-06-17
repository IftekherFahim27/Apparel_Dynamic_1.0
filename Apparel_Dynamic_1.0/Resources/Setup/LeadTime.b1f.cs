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
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_2").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_3").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_4").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("Item_5").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_6").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_7").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_9").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_10").Specific));
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
