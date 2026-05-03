using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apparel_Dynamic_1._0.Resources.Transaction
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Transaction.CAD", "Resources/Transaction/CAD.b1f")]
    class CAD : UserFormBase
    {
        public CAD()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Item_9").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_10").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_11").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_12").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("Item_13").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("Item_14").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("Item_15").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("Item_16").Specific));
            this.EditText7 = ((SAPbouiCOM.EditText)(this.GetItem("Item_17").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_18").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_19").Specific));
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_20").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_21").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_22").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_23").Specific));
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("Item_24").Specific));
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_25").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_26").Specific));
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("Item_27").Specific));
            this.Folder2 = ((SAPbouiCOM.Folder)(this.GetItem("Item_28").Specific));
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_29").Specific));
            this.Matrix2 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_30").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_31").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("Item_32").Specific));
            this.Grid1 = ((SAPbouiCOM.Grid)(this.GetItem("Item_33").Specific));
            this.Matrix3 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_34").Specific));
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
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.EditText EditText7;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Matrix Matrix1;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button Button3;
        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.Folder Folder1;
        private SAPbouiCOM.Folder Folder2;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Matrix Matrix2;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.EditText EditText8;
        private SAPbouiCOM.Grid Grid1;
        private SAPbouiCOM.Matrix Matrix3;
    }
}
