using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apparel_Dynamic_1._0.Resources.Master
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Master.StyleMaster", "Resources/Master/StyleMaster.b1f")]
    class StyleMaster : UserFormBase
    {
        public StyleMaster()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("STSLCODE").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("STCSCODE").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("STCSDESC").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("STGSM").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("STGENDER").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("STSMTPCD").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("STPDTPCD").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("STPDLNCD").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_9").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_10").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_11").Specific));
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_12").Specific));
            this.StaticText13 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_13").Specific));
            this.StaticText14 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_14").Specific));
            this.StaticText15 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_15").Specific));
            this.StaticText16 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_16").Specific));
            this.StaticText17 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_17").Specific));
            this.StaticText18 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_18").Specific));
            this.StaticText19 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_19").Specific));
            this.StaticText20 = ((SAPbouiCOM.StaticText)(this.GetItem("STUOM").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("ETSLCODE").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("ETCSCODE").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("ETCSDESC").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("ETGSM").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("ETGENDER").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("ETSMTPCD").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("ETSMTPNM").Specific));
            this.EditText7 = ((SAPbouiCOM.EditText)(this.GetItem("ETPDTPCD").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("ETPDTPNM").Specific));
            this.EditText9 = ((SAPbouiCOM.EditText)(this.GetItem("ETPDLNCD").Specific));
            this.EditText10 = ((SAPbouiCOM.EditText)(this.GetItem("ETPDLNNM").Specific));
            this.EditText11 = ((SAPbouiCOM.EditText)(this.GetItem("Item_33").Specific));
            this.EditText12 = ((SAPbouiCOM.EditText)(this.GetItem("Item_34").Specific));
            this.EditText13 = ((SAPbouiCOM.EditText)(this.GetItem("Item_35").Specific));
            this.EditText14 = ((SAPbouiCOM.EditText)(this.GetItem("Item_36").Specific));
            this.EditText15 = ((SAPbouiCOM.EditText)(this.GetItem("Item_37").Specific));
            this.EditText16 = ((SAPbouiCOM.EditText)(this.GetItem("Item_38").Specific));
            this.EditText17 = ((SAPbouiCOM.EditText)(this.GetItem("Item_39").Specific));
            this.EditText18 = ((SAPbouiCOM.EditText)(this.GetItem("Item_40").Specific));
            this.EditText19 = ((SAPbouiCOM.EditText)(this.GetItem("Item_41").Specific));
            this.EditText20 = ((SAPbouiCOM.EditText)(this.GetItem("Item_42").Specific));
            this.EditText21 = ((SAPbouiCOM.EditText)(this.GetItem("Item_43").Specific));
            this.EditText22 = ((SAPbouiCOM.EditText)(this.GetItem("Item_44").Specific));
            this.EditText23 = ((SAPbouiCOM.EditText)(this.GetItem("Item_45").Specific));
            this.EditText24 = ((SAPbouiCOM.EditText)(this.GetItem("Item_46").Specific));
            this.EditText25 = ((SAPbouiCOM.EditText)(this.GetItem("Item_47").Specific));
            this.EditText26 = ((SAPbouiCOM.EditText)(this.GetItem("Item_48").Specific));
            this.EditText27 = ((SAPbouiCOM.EditText)(this.GetItem("Item_49").Specific));
            this.EditText28 = ((SAPbouiCOM.EditText)(this.GetItem("Item_50").Specific));
            this.EditText29 = ((SAPbouiCOM.EditText)(this.GetItem("Item_51").Specific));
            this.EditText30 = ((SAPbouiCOM.EditText)(this.GetItem("Item_52").Specific));
            this.EditText31 = ((SAPbouiCOM.EditText)(this.GetItem("Item_53").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_54").Specific));
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_55").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("FOLSIZE").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_58").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_59").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_60").Specific));
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("Item_61").Specific));
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("FOLCOLOR").Specific));
            this.Folder2 = ((SAPbouiCOM.Folder)(this.GetItem("FOLITEM").Specific));
            this.Folder3 = ((SAPbouiCOM.Folder)(this.GetItem("FOLATTAC").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_65").Specific));
            this.EditText32 = ((SAPbouiCOM.EditText)(this.GetItem("Item_66").Specific));
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("Item_67").Specific));
            this.StaticText22 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_68").Specific));
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("MTXCOLOR").Specific));
            this.Matrix2 = ((SAPbouiCOM.Matrix)(this.GetItem("MTXSBCLR").Specific));
            this.Matrix3 = ((SAPbouiCOM.Matrix)(this.GetItem("MTXITEM").Specific));
            this.Matrix4 = ((SAPbouiCOM.Matrix)(this.GetItem("MTXATTCH").Specific));
            this.Button5 = ((SAPbouiCOM.Button)(this.GetItem("Item_73").Specific));
            this.Button6 = ((SAPbouiCOM.Button)(this.GetItem("Item_74").Specific));
            this.Button7 = ((SAPbouiCOM.Button)(this.GetItem("Item_75").Specific));
            this.StaticText23 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_76").Specific));
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
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.StaticText StaticText10;
        private SAPbouiCOM.StaticText StaticText11;
        private SAPbouiCOM.StaticText StaticText12;
        private SAPbouiCOM.StaticText StaticText13;
        private SAPbouiCOM.StaticText StaticText14;
        private SAPbouiCOM.StaticText StaticText15;
        private SAPbouiCOM.StaticText StaticText16;
        private SAPbouiCOM.StaticText StaticText17;
        private SAPbouiCOM.StaticText StaticText18;
        private SAPbouiCOM.StaticText StaticText19;
        private SAPbouiCOM.StaticText StaticText20;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.EditText EditText7;
        private SAPbouiCOM.EditText EditText8;
        private SAPbouiCOM.EditText EditText9;
        private SAPbouiCOM.EditText EditText10;
        private SAPbouiCOM.EditText EditText11;
        private SAPbouiCOM.EditText EditText12;
        private SAPbouiCOM.EditText EditText13;
        private SAPbouiCOM.EditText EditText14;
        private SAPbouiCOM.EditText EditText15;
        private SAPbouiCOM.EditText EditText16;
        private SAPbouiCOM.EditText EditText17;
        private SAPbouiCOM.EditText EditText18;
        private SAPbouiCOM.EditText EditText19;
        private SAPbouiCOM.EditText EditText20;
        private SAPbouiCOM.EditText EditText21;
        private SAPbouiCOM.EditText EditText22;
        private SAPbouiCOM.EditText EditText23;
        private SAPbouiCOM.EditText EditText24;
        private SAPbouiCOM.EditText EditText25;
        private SAPbouiCOM.EditText EditText26;
        private SAPbouiCOM.EditText EditText27;
        private SAPbouiCOM.EditText EditText28;
        private SAPbouiCOM.EditText EditText29;
        private SAPbouiCOM.EditText EditText30;
        private SAPbouiCOM.EditText EditText31;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button Button3;
        private SAPbouiCOM.Folder Folder1;
        private SAPbouiCOM.Folder Folder2;
        private SAPbouiCOM.Folder Folder3;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.EditText EditText32;
        private SAPbouiCOM.Button Button4;
        private SAPbouiCOM.StaticText StaticText22;
        private SAPbouiCOM.Matrix Matrix1;
        private SAPbouiCOM.Matrix Matrix2;
        private SAPbouiCOM.Matrix Matrix3;
        private SAPbouiCOM.Matrix Matrix4;
        private SAPbouiCOM.Button Button5;
        private SAPbouiCOM.Button Button6;
        private SAPbouiCOM.Button Button7;
        private SAPbouiCOM.StaticText StaticText23;
    }
}
