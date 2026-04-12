using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apparel_Dynamic_1._0.Resources.Transaction
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Transaction.OTT", "Resources/Transaction/OTT.b1f")]
    class OTT : UserFormBase
    {
        public OTT()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.STOTDATE = ((SAPbouiCOM.StaticText)(this.GetItem("STOTDATE").Specific));
            this.STSTATUS = ((SAPbouiCOM.StaticText)(this.GetItem("STSTATUS").Specific));
            this.STDOCNUM = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCNUM").Specific));
            this.STMERCHN = ((SAPbouiCOM.StaticText)(this.GetItem("STMERCHN").Specific));
            this.FOLOTDTL = ((SAPbouiCOM.Folder)(this.GetItem("FOLOTDTL").Specific));
            this.FOLDODTL = ((SAPbouiCOM.Folder)(this.GetItem("FOLDODTL").Specific));
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.ETOTDATE = ((SAPbouiCOM.EditText)(this.GetItem("ETOTDATE").Specific));
            this.CBSTATUS = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSTATUS").Specific));
            this.ETDOCNUM = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCNUM").Specific));
            this.CBSERIES = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSERIES").Specific));
            this.ETMERCNM = ((SAPbouiCOM.EditText)(this.GetItem("ETMERCNM").Specific));
            this.ETMERCHN = ((SAPbouiCOM.EditText)(this.GetItem("ETMERCHN").Specific));
            this.ETMERCHN.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETMERCHN_ChooseFromListAfter);
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            this.MTXOTDTL = ((SAPbouiCOM.Matrix)(this.GetItem("MTXOTDTL").Specific));
            this.MTXOTDTL.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.MTXOTDTL_ChooseFromListAfter);
            this.GRDDODTL = ((SAPbouiCOM.Grid)(this.GetItem("GRDDODTL").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.DataLoadAfter += new DataLoadAfterHandler(this.Form_DataLoadAfter);

        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {

        }

        private SAPbouiCOM.StaticText STOTDATE, STSTATUS, STDOCNUM, STMERCHN;
        private SAPbouiCOM.Folder FOLOTDTL, FOLDODTL;
        private SAPbouiCOM.Button ADDButton, CancelButton;

 
        private SAPbouiCOM.EditText ETOTDATE, ETDOCNUM, ETMERCNM, ETMERCHN, ETDOCTRY;



        private SAPbouiCOM.ComboBox CBSTATUS, CBSERIES;
        private SAPbouiCOM.Matrix MTXOTDTL;
        private SAPbouiCOM.Grid GRDDODTL;

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            //throw new System.NotImplementedException();

        }

        private void ETMERCHN_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("empID", 0).ToString().Trim();
            string Name = dt.GetValue("U_FNAME", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETMERCHN").Specific;
            ETCD.Value = Code;
            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETMERCNM").Specific;
            ETNM.Value = Name;

        }

        private void MTXOTDTL_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            throw new System.NotImplementedException();

        }

    }
}
