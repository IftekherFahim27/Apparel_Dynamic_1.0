using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;

namespace Apparel_Dynamic_1._0.Helper
{
    class Global
    {
        public static SAPbouiCOM.Application G_UI_Application;
        public static SAPbobsCOM.Company oComp; // Varible for company 
        public static GlobalFunction GFunc = new GlobalFunction();

        public static string Password = "";
        public static SAPbouiCOM.SboGuiApi G_UI_SboGuiApi;
        public static SAPbobsCOM.SBObob G_DI_SBObob;

        public static SAPbouiCOM.Form G_Form;
        public static SAPbouiCOM.Form G_ParentForm;
        public static SAPbouiCOM.Form G_TargetForm;
        public static SAPbouiCOM.Form GRN_Form;

        public static string qStr;
        public static SAPbobsCOM.Recordset rSet;

        public static SAPbouiCOM.Folder oFolder;
        public static SAPbouiCOM.StaticText oStatic;
        public static SAPbouiCOM.EditText oEdit;
        public static SAPbouiCOM.Item oItem;
        public static SAPbouiCOM.Matrix oMatrix;
        public static SAPbouiCOM.Column oColumn;
        public static SAPbouiCOM.Cell oCell;
        public static SAPbouiCOM.ChooseFromListCollection oCfls;
        public static SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams;
        public static SAPbouiCOM.ChooseFromList oCfl;
        public static SAPbouiCOM.Conditions oCons;
        public static SAPbouiCOM.Condition oCon;
        public static SAPbouiCOM.ComboBox oCombo;
        public static SAPbouiCOM.CheckBox oCheck;
        public static SAPbouiCOM.OptionBtn oOptButton;
        public static SAPbouiCOM.LinkedButton oLinkButton;
        public static SAPbouiCOM.Button oButton;
        public static SAPbouiCOM.Grid oGrid;
        public static SAPbouiCOM.GridColumn oGridColumn;
        public static SAPbouiCOM.DataTable oDataTable;
        public static SAPbouiCOM.ProgressBar oProgressBar;
        public static SAPbouiCOM.EditTextColumn oEditTextColumn;
        public static SAPbouiCOM.ComboBoxColumn oComboBoxColumn;
        public static SAPbobsCOM.Message oMessage;

        public static int G_LastFormId;

        //public static OpenFileDialog legacyFileDialog = new OpenFileDialog();
        public static string legacyFileName;

        public static int currentProcess;
        public static bool FromSoMenu = false;
        public static bool ContFromAmmed = false;
        public static bool GRNBatch = false;
        public static string GRNNewBatchNo;
        public static bool GRNAddd = false;
        public static string GRNSeries;
        public static string GRNItem;
        public static bool MarchendiserUser = false;
        public static int SalesEmpUser = -1;
        public static int EmpIdUser = -1;
    }
}
