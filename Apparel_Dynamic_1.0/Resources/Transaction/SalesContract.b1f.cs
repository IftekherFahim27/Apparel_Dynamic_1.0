using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Apparel_Dynamic_1._0.Helper;

namespace Apparel_Dynamic_1._0.Resources.Transaction
{
    [FormAttribute("Apparel_Dynamic_1._0.Resources.Transaction.SalesContract", "Resources/Transaction/SalesContract.b1f")]
    class SalesContract : UserFormBase
    {
        public SalesContract()
        {
        }

        // -------- Static Text --------
        private SAPbouiCOM.StaticText STBRANCH, STMRSTAT, STCMSTAT, STCUSTMR, STBRNDCD,
                                     STSCNO, STSCDESC, STREFNCE, STSCVAL, STDOVAL,
                                     STB2BPER, STB2BVAL, STDSNBNK, STDOCNUM, STDOCDAT,
                                     STCBNKAC, STCUSBNK, STOBNKAC, STOWNBNK, STCURR,
                                     STISUDAT, STSHPDAT, STEXPDAT, STTOLPER, STAMNDNO,
                                     STOPNAMT, STINTRMS, STPYTRMS, STMDSHIP, STPORTLD,
                                     STCNDEST, STPRTDIS, STINSNCE, STSHPTOL, STHSCODE,
                                     STDOCREQ, STRMSCON, STSHPADD, STPRTSHP;

       











        // -------- ComboBox --------
        private SAPbouiCOM.ComboBox CBBRANCH, CBMRSTAT, CBCMSTAT, CBINTRMS, CBDSNBNK,
                                    CBPYTRMS, CBMDSHIP, CBPRTSHP, CBSERIES, CBSHPADD;

        // -------- Edit Text --------
        private SAPbouiCOM.EditText ETCUSTMR, ETCUSTNM, ETBRNDCD, ETBRNDNM, ETSCNO,
                                    ETSCDESC, ETREFNCE, ETSCVAL, ETDOVAL, ETB2BPER,
                                    ETB2BVAL, ETDOCNUM, ETDOCDAT, ETCBNKAC,
                                    ETCUSBNK, ETOBNKAC, ETOWNBNK, ETCURR, ETISUDAT,
                                    ETSHPDAT, ETEXPDAT, ETTOLPER, ETAMNDNO, ETOPNAMT,
                                    ETPORTLD, ETCNDEST, ETPRTDIS, ETINSNCE, ETSHPTOL,
                                    ETHSCODE, ETDOCREQ, ETRMSCON, ETDOCTRY;

        // -------- Folder --------
        private SAPbouiCOM.Folder FOLORDTL, FOLAMEND, FOLB2BDL, FOLATTCH, FOLCDTLS;

        // -------- Button --------
        private SAPbouiCOM.Button ADDButton, CancelButton, BTNAMEND, BRWSBTN, DISPBTN, DELBTN;

        // -------- Matrix --------
        private SAPbouiCOM.Matrix MTXORDTL, MTXATTCH, MTXB2BDL;

        // -------- Grid --------
        private SAPbouiCOM.Grid GRDAMEND;
        public override void OnInitializeComponent()
        {
            //             -------- Static Text --------
            this.STBRANCH = ((SAPbouiCOM.StaticText)(this.GetItem("STBRANCH").Specific));
            this.STMRSTAT = ((SAPbouiCOM.StaticText)(this.GetItem("STMRSTAT").Specific));
            this.STCMSTAT = ((SAPbouiCOM.StaticText)(this.GetItem("STCMSTAT").Specific));
            this.STCUSTMR = ((SAPbouiCOM.StaticText)(this.GetItem("STCUSTMR").Specific));
            this.STBRNDCD = ((SAPbouiCOM.StaticText)(this.GetItem("STBRNDCD").Specific));
            this.STSCNO = ((SAPbouiCOM.StaticText)(this.GetItem("STSCNO").Specific));
            this.STSCDESC = ((SAPbouiCOM.StaticText)(this.GetItem("STSCDESC").Specific));
            this.STREFNCE = ((SAPbouiCOM.StaticText)(this.GetItem("STREFNCE").Specific));
            this.STSCVAL = ((SAPbouiCOM.StaticText)(this.GetItem("STSCVAL").Specific));
            this.STDOVAL = ((SAPbouiCOM.StaticText)(this.GetItem("STDOVAL").Specific));
            this.STB2BPER = ((SAPbouiCOM.StaticText)(this.GetItem("STB2BPER").Specific));
            this.STB2BVAL = ((SAPbouiCOM.StaticText)(this.GetItem("STB2BVAL").Specific));
            this.STDSNBNK = ((SAPbouiCOM.StaticText)(this.GetItem("STDSNBNK").Specific));
            this.STDOCNUM = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCNUM").Specific));
            this.STDOCDAT = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCDAT").Specific));
            this.STCBNKAC = ((SAPbouiCOM.StaticText)(this.GetItem("STCBNKAC").Specific));
            this.STCUSBNK = ((SAPbouiCOM.StaticText)(this.GetItem("STCUSBNK").Specific));
            this.STOBNKAC = ((SAPbouiCOM.StaticText)(this.GetItem("STOBNKAC").Specific));
            this.STOWNBNK = ((SAPbouiCOM.StaticText)(this.GetItem("STOWNBNK").Specific));
            this.STCURR = ((SAPbouiCOM.StaticText)(this.GetItem("STCURR").Specific));
            this.STISUDAT = ((SAPbouiCOM.StaticText)(this.GetItem("STISUDAT").Specific));
            this.STSHPDAT = ((SAPbouiCOM.StaticText)(this.GetItem("STSHPDAT").Specific));
            this.STEXPDAT = ((SAPbouiCOM.StaticText)(this.GetItem("STEXPDAT").Specific));
            this.STTOLPER = ((SAPbouiCOM.StaticText)(this.GetItem("STTOLPER").Specific));
            this.STAMNDNO = ((SAPbouiCOM.StaticText)(this.GetItem("STAMNDNO").Specific));
            this.STOPNAMT = ((SAPbouiCOM.StaticText)(this.GetItem("STOPNAMT").Specific));
            this.STINTRMS = ((SAPbouiCOM.StaticText)(this.GetItem("STINTRMS").Specific));
            this.STPYTRMS = ((SAPbouiCOM.StaticText)(this.GetItem("STPYTRMS").Specific));
            this.STMDSHIP = ((SAPbouiCOM.StaticText)(this.GetItem("STMDSHIP").Specific));
            this.STPORTLD = ((SAPbouiCOM.StaticText)(this.GetItem("STPORTLD").Specific));
            this.STCNDEST = ((SAPbouiCOM.StaticText)(this.GetItem("STCNDEST").Specific));
            this.STPRTDIS = ((SAPbouiCOM.StaticText)(this.GetItem("STPRTDIS").Specific));
            this.STINSNCE = ((SAPbouiCOM.StaticText)(this.GetItem("STINSNCE").Specific));
            this.STSHPTOL = ((SAPbouiCOM.StaticText)(this.GetItem("STSHPTOL").Specific));
            this.STHSCODE = ((SAPbouiCOM.StaticText)(this.GetItem("STHSCODE").Specific));
            this.STDOCREQ = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCREQ").Specific));
            this.STRMSCON = ((SAPbouiCOM.StaticText)(this.GetItem("STRMSCON").Specific));
            this.STSHPADD = ((SAPbouiCOM.StaticText)(this.GetItem("STSHPADD").Specific));
            this.STPRTSHP = ((SAPbouiCOM.StaticText)(this.GetItem("STPRTSHP").Specific));
            //             -------- Folder --------
            this.FOLORDTL = ((SAPbouiCOM.Folder)(this.GetItem("FOLORDTL").Specific));
            this.FOLAMEND = ((SAPbouiCOM.Folder)(this.GetItem("FOLAMEND").Specific));
            this.FOLB2BDL = ((SAPbouiCOM.Folder)(this.GetItem("FOLB2BDL").Specific));
            this.FOLATTCH = ((SAPbouiCOM.Folder)(this.GetItem("FOLATTCH").Specific));
            this.FOLCDTLS = ((SAPbouiCOM.Folder)(this.GetItem("FOLCDTLS").Specific));
            //             -------- Button --------
            this.ADDButton = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.ADDButton.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.ADDButton_PressedAfter);
            this.ADDButton.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.ADDButton_PressedBefore);
            this.CancelButton = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.BTNAMEND = ((SAPbouiCOM.Button)(this.GetItem("BTNAMEND").Specific));
            this.BTNAMEND.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.BTNAMEND_PressedAfter);
            this.BRWSBTN = ((SAPbouiCOM.Button)(this.GetItem("BRWSBTN").Specific));
            this.BRWSBTN.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.BRWSBTN_ClickAfter);
            this.DISPBTN = ((SAPbouiCOM.Button)(this.GetItem("DISPBTN").Specific));
            this.DISPBTN.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.DISPBTN_ClickAfter);
            this.DELBTN = ((SAPbouiCOM.Button)(this.GetItem("DELBTN").Specific));
            this.DELBTN.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.DELBTN_ClickAfter);
            //             -------- ComboBox --------
            this.CBBRANCH = ((SAPbouiCOM.ComboBox)(this.GetItem("CBBRANCH").Specific));
            this.CBMRSTAT = ((SAPbouiCOM.ComboBox)(this.GetItem("CBMRSTAT").Specific));
            this.CBMRSTAT.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.CBMRSTAT_ComboSelectAfter);
            this.CBCMSTAT = ((SAPbouiCOM.ComboBox)(this.GetItem("CBCMSTAT").Specific));
            this.CBCMSTAT.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.CBCMSTAT_ComboSelectAfter);
            this.CBINTRMS = ((SAPbouiCOM.ComboBox)(this.GetItem("CBINTRMS").Specific));
            this.CBPYTRMS = ((SAPbouiCOM.ComboBox)(this.GetItem("CBPYTRMS").Specific));
            this.CBMDSHIP = ((SAPbouiCOM.ComboBox)(this.GetItem("CBMDSHIP").Specific));
            this.CBPRTSHP = ((SAPbouiCOM.ComboBox)(this.GetItem("CBPRTSHP").Specific));
            this.CBSERIES = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSERIES").Specific));
            this.CBDSNBNK = ((SAPbouiCOM.ComboBox)(this.GetItem("CBDSNBNK").Specific));
            this.CBSHPADD = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSHPADD").Specific));
            this.CBDSNBNK.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.CBDSNBNK_ComboSelectAfter);
            //             -------- Edit Text --------
            this.ETCUSTMR = ((SAPbouiCOM.EditText)(this.GetItem("ETCUSTMR").Specific));
            this.ETCUSTMR.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETCUSTMR_ChooseFromListBefore);
            this.ETCUSTMR.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETCUSTMR_ChooseFromListAfter);
            this.ETCUSTNM = ((SAPbouiCOM.EditText)(this.GetItem("ETCUSTNM").Specific));
            this.ETBRNDCD = ((SAPbouiCOM.EditText)(this.GetItem("ETBRNDCD").Specific));
            this.ETBRNDCD.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETBRNDCD_ChooseFromListBefore);
            this.ETBRNDCD.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETBRNDCD_ChooseFromListAfter);
            this.ETBRNDNM = ((SAPbouiCOM.EditText)(this.GetItem("ETBRNDNM").Specific));
            this.ETSCNO = ((SAPbouiCOM.EditText)(this.GetItem("ETSCNO").Specific));
            this.ETSCNO.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.ETSCNO_LostFocusAfter);
            this.ETSCDESC = ((SAPbouiCOM.EditText)(this.GetItem("ETSCDESC").Specific));
            this.ETREFNCE = ((SAPbouiCOM.EditText)(this.GetItem("ETREFNCE").Specific));
            this.ETSCVAL = ((SAPbouiCOM.EditText)(this.GetItem("ETSCVAL").Specific));
            this.ETDOVAL = ((SAPbouiCOM.EditText)(this.GetItem("ETDOVAL").Specific));
            this.ETB2BPER = ((SAPbouiCOM.EditText)(this.GetItem("ETB2BPER").Specific));
            this.ETB2BVAL = ((SAPbouiCOM.EditText)(this.GetItem("ETB2BVAL").Specific));
            this.ETDOCNUM = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCNUM").Specific));
            this.ETDOCDAT = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCDAT").Specific));
            this.ETCBNKAC = ((SAPbouiCOM.EditText)(this.GetItem("ETCBNKAC").Specific));
            this.ETCBNKAC.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETCBNKAC_ChooseFromListAfter);
            this.ETCBNKAC.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETCBNKAC_ChooseFromListBefore);
            this.ETCUSBNK = ((SAPbouiCOM.EditText)(this.GetItem("ETCUSBNK").Specific));
            this.ETOBNKAC = ((SAPbouiCOM.EditText)(this.GetItem("ETOBNKAC").Specific));
            this.ETOBNKAC.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETOBNKAC_ChooseFromListAfter);
            this.ETOWNBNK = ((SAPbouiCOM.EditText)(this.GetItem("ETOWNBNK").Specific));
            this.ETCURR = ((SAPbouiCOM.EditText)(this.GetItem("ETCURR").Specific));
            this.ETCURR.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETCURR_ChooseFromListAfter);
            this.ETISUDAT = ((SAPbouiCOM.EditText)(this.GetItem("ETISUDAT").Specific));
            this.ETSHPDAT = ((SAPbouiCOM.EditText)(this.GetItem("ETSHPDAT").Specific));
            this.ETEXPDAT = ((SAPbouiCOM.EditText)(this.GetItem("ETEXPDAT").Specific));
            this.ETTOLPER = ((SAPbouiCOM.EditText)(this.GetItem("ETTOLPER").Specific));
            this.ETAMNDNO = ((SAPbouiCOM.EditText)(this.GetItem("ETAMNDNO").Specific));
            this.ETOPNAMT = ((SAPbouiCOM.EditText)(this.GetItem("ETOPNAMT").Specific));
            this.ETPORTLD = ((SAPbouiCOM.EditText)(this.GetItem("ETPORTLD").Specific));
            this.ETCNDEST = ((SAPbouiCOM.EditText)(this.GetItem("ETCNDEST").Specific));
            this.ETPRTDIS = ((SAPbouiCOM.EditText)(this.GetItem("ETPRTDIS").Specific));
            this.ETINSNCE = ((SAPbouiCOM.EditText)(this.GetItem("ETINSNCE").Specific));
            this.ETSHPTOL = ((SAPbouiCOM.EditText)(this.GetItem("ETSHPTOL").Specific));
            this.ETHSCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETHSCODE").Specific));
            this.ETDOCREQ = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCREQ").Specific));
            this.ETRMSCON = ((SAPbouiCOM.EditText)(this.GetItem("ETRMSCON").Specific));
            this.ETDOCTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCTRY").Specific));
            //             -------- Matrix --------
            this.MTXORDTL = ((SAPbouiCOM.Matrix)(this.GetItem("MTXORDTL").Specific));
            this.MTXATTCH = ((SAPbouiCOM.Matrix)(this.GetItem("MTXATTCH").Specific));
            this.MTXB2BDL = ((SAPbouiCOM.Matrix)(this.GetItem("MTXB2BDL").Specific));
            //             -------- Grid --------
            this.GRDAMEND = ((SAPbouiCOM.Grid)(this.GetItem("GRDAMEND").Specific));
            this.OnCustomInitialize();

        }

        public override void OnInitializeFormEvents()
        {
            this.DataLoadAfter += new DataLoadAfterHandler(this.Form_DataLoadAfter);

        }



        private void OnCustomInitialize()
        {

        }

        private void BTNAMEND_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            throw new System.NotImplementedException();

        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

                string scNo = ((SAPbouiCOM.EditText)oForm.Items.Item("ETSCNO").Specific).Value.Trim();

                if (string.IsNullOrWhiteSpace(scNo))
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Sales Contract is not used in Draft Order.",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning
                    );
                    return;
                }

                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix =
                    (SAPbouiCOM.Matrix)oForm.Items.Item("MTXORDTL").Specific;

                SAPbouiCOM.DBDataSource oDBDS =
                    oForm.DataSources.DBDataSources.Item("@FIL_DR_SCM1"); // change this

                oDBDS.Clear();
                oMatrix.Clear();

                string safeSCNo = scNo.Replace("'", "''");

                string qStr = @"
                                SELECT 
                                    ROW_NUMBER() OVER (ORDER BY A.""DocEntry"") AS ""LineId"",
                                    A.""DocNum"",
                                    A.""DocEntry"",
                                    A.""U_STYLECODE"",
                                    A.""U_STYLENTRY"",
                                    A.""DocDate"",
                                    A.""DocDueDate"",
                                    SUM(B.""Quantity"") AS ""Quantity"",
                                    CASE 
                                        WHEN A.""DocCur"" = 'BDT' THEN A.""DocTotal""
                                        ELSE A.""DocTotalFC""
                                    END AS ""TotalValue""
                                FROM OQUT AS A
                                INNER JOIN QUT1 AS B
                                    ON A.""DocEntry"" = B.""DocEntry""
                                WHERE A.""U_SCNO"" = '" + safeSCNo + @"'
                                GROUP BY 
                                    A.""DocNum"",
                                    A.""DocEntry"",
                                    A.""U_STYLECODE"",
                                    A.""U_STYLENTRY"",
                                    A.""DocDate"",
                                    A.""DocDueDate"",
                                    A.""DocCur"",
                                    A.""DocTotal"",
                                    A.""DocTotalFC""
                                ORDER BY A.""DocEntry""";

                SAPbobsCOM.Recordset rs =
                    (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                rs.DoQuery(qStr);

                int row = 0;

                while (!rs.EoF)
                {
                    oDBDS.InsertRecord(row);

                    oDBDS.SetValue("LineId", row, rs.Fields.Item("LineId").Value.ToString());
                    oDBDS.SetValue("U_DRFTORDN", row, rs.Fields.Item("DocNum").Value.ToString());
                    oDBDS.SetValue("U_DRFTNRY", row, rs.Fields.Item("DocEntry").Value.ToString());
                    oDBDS.SetValue("U_STYLECODE", row, rs.Fields.Item("U_STYLECODE").Value.ToString());
                    oDBDS.SetValue("U_STYLENTRY", row, rs.Fields.Item("U_STYLENTRY").Value.ToString());

                    DateTime docDate = Convert.ToDateTime(rs.Fields.Item("DocDate").Value);
                    DateTime dueDate = Convert.ToDateTime(rs.Fields.Item("DocDueDate").Value);

                    oDBDS.SetValue("U_DOCDATE", row, docDate.ToString("yyyyMMdd"));
                    oDBDS.SetValue("U_DUEDATE", row, dueDate.ToString("yyyyMMdd"));

                    oDBDS.SetValue("U_TOTALQTY", row, rs.Fields.Item("Quantity").Value.ToString());
                    oDBDS.SetValue("U_TOTALVALUE", row, rs.Fields.Item("TotalValue").Value.ToString());

                    row++;
                    rs.MoveNext();
                }

                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();

                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                try { Application.SBO_Application.Forms.Item(pVal.FormUID).Freeze(false); } catch { }

                Application.SBO_Application.StatusBar.SetText(
                    "Form_DataLoadAfter Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
        }

        private void ADDButton_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.DBDataSource oDBH = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@FIL_DH_OSCM");
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                SAPbouiCOM.ComboBox ocmb = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBSERIES").Specific;
                Global.GFunc.LoadComboBoxSeries(ocmb, "FIL_D_OSCM");  //Object Type
                string ocmbvalue = ocmb.Selected.Value;
                long docno = oForm.BusinessObject.GetNextSerialNumber(ocmbvalue, "FIL_D_OSCM");

                oDBH.SetValue("DocNum", 0, docno.ToString()); // only set the value in string.

                //Date
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETDOCDAT").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
                //Amendment No
                ((SAPbouiCOM.EditText)oForm.Items.Item("ETAMNDNO").Specific).Value = "1";

                // Branch combo
                string sqlQuerybpl = @"SELECT ""BPLId"", ""BPLName"" FROM ""OBPL""";
                SAPbouiCOM.ComboBox CBCMPANY = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBBRANCH").Specific;
                Global.GFunc.setComboBoxValue(CBCMPANY, sqlQuerybpl);
                CBCMPANY.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

                //Payment Terms
                string payTerms = @"SELECT ""GroupNum"", ""PymntGroup"" FROM ""OCTG""";
                SAPbouiCOM.ComboBox CBPYTRMS = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBPYTRMS").Specific;
                Global.GFunc.setComboBoxValue(CBPYTRMS, payTerms);

                //Shipping Type
                string shipType = @"SELECT ""TrnspCode"", ""TrnspName"" FROM ""OSHP""";
                SAPbouiCOM.ComboBox CBMDSHIP = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBMDSHIP").Specific;
                Global.GFunc.setComboBoxValue(CBMDSHIP, shipType);
            }
                       
        }

        private void ADDButton_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                ValidateForm(ref oForm, ref BubbleEvent);
            }

        }

        private bool ValidateForm(ref SAPbouiCOM.Form oForm, ref bool BubbleEvent)
        {
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                string branch = oForm.DataSources.DBDataSources.Item("@FIL_DH_OSCM").GetValue("U_BRANCH", 0).Trim();
                string customer = oForm.DataSources.DBDataSources.Item("@FIL_DH_OSCM").GetValue("U_CARDCODE", 0).Trim();
                string brand = oForm.DataSources.DBDataSources.Item("@FIL_DH_OSCM").GetValue("U_BRAND", 0).Trim();
                string scNo = oForm.DataSources.DBDataSources.Item("@FIL_DH_OSCM").GetValue("U_SCNO", 0).Trim();
                string issueDate = oForm.DataSources.DBDataSources.Item("@FIL_DH_OSCM").GetValue("U_ISSUEDATE", 0).Trim();
                

                if (branch == "")
                {
                    Global.GFunc.ShowError("Select Branch");
                    oForm.ActiveItem = "CBBRANCH";
                    return BubbleEvent = false;
                }
                if (customer == "")
                {
                    Global.GFunc.ShowError("Enter Customer Code");
                    oForm.ActiveItem = "ETCUSTMR";
                    return BubbleEvent = false;
                }
                else if (brand == "")
                {
                    Global.GFunc.ShowError("Enter Brand Code");
                    oForm.ActiveItem = "ETBRNDCD";
                    return BubbleEvent = false;
                }
                else if (scNo == "")
                {
                    Global.GFunc.ShowError("Enter Sales Contract No ");
                    oForm.ActiveItem = "ETSCNO";
                    return BubbleEvent = false;
                }
                else if (issueDate == "")
                {
                    Global.GFunc.ShowError("Enter Issue Date");
                    oForm.ActiveItem = "ETISUDAT";
                    return BubbleEvent = false;
                }

                //Prevent New Empty Free Line
                PreventEmptyLastRow(oForm, "@FIL_DR_SCM1", MTXORDTL, "U_DRFTORDN");
                PreventEmptyLastRow(oForm, "@FIL_DR_SCM2", MTXB2BDL, "U_BTBLCDNO");
                PreventEmptyLastRow(oForm, "@FIL_DR_SCM3", MTXATTCH, "U_ATCHMENT");
                
            }

            return BubbleEvent;
        }

        private void PreventEmptyLastRow(SAPbouiCOM.Form oForm, string dbDatasourceUID, SAPbouiCOM.Matrix matrix, string columnName)
        {
            SAPbouiCOM.DBDataSource oDB = oForm.DataSources.DBDataSources.Item(dbDatasourceUID);
            int rowCount = matrix.VisualRowCount;

            if (rowCount > 0)
            {
                string lastValue = oDB.GetValue(columnName, rowCount - 1).Trim();

                if (string.IsNullOrEmpty(lastValue) || lastValue.Equals("0.0"))
                {
                    matrix.DeleteRow(rowCount);
                    oDB.RemoveRecord(rowCount - 1);
                }
            }
        }

        private void DELBTN_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.Matrix MTXATTCH = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXATTCH").Specific;

            for (int i = 1; i <= MTXATTCH.RowCount; i++)
            {
                if (MTXATTCH.IsRowSelected(i))
                {
                    string filePath = ((SAPbouiCOM.EditText)MTXATTCH.Columns.Item("CLATTACH").Cells.Item(i).Specific).Value;
                    if (!string.IsNullOrEmpty(filePath) && System.IO.File.Exists(filePath))
                    {
                        System.Diagnostics.Process.Start(filePath);
                    }
                    else
                    {
                        Application.SBO_Application.MessageBox("File does not exist or path is empty.");
                    }
                    break;
                }
            }

        }

        private void DISPBTN_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.Matrix MTXATTCH = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXATTCH").Specific;

            for (int i = 1; i <= MTXATTCH.RowCount; i++)
            {
                if (MTXATTCH.IsRowSelected(i))
                {
                    string filePath = ((SAPbouiCOM.EditText)MTXATTCH.Columns.Item("CLATTACH").Cells.Item(i).Specific).Value;
                    if (!string.IsNullOrEmpty(filePath) && System.IO.File.Exists(filePath))
                    {
                        System.Diagnostics.Process.Start(filePath);
                    }
                    else
                    {
                        Application.SBO_Application.MessageBox("File does not exist or path is empty.");
                    }
                    break;
                }
            };

        }

        private void BRWSBTN_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.DBDataSource DBDataSourceLine = oForm.DataSources.DBDataSources.Item("@FIL_DR_SCM3");
            SAPbouiCOM.Matrix MTXATTCH = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXATTCH").Specific;

            string filePath = FileDialogHelper.ShowFileDialog();
            if (!string.IsNullOrEmpty(filePath))
            {
                int lastRow = MTXATTCH.VisualRowCount;
                bool needNewRow = (lastRow == 0) ||
                                  !string.IsNullOrEmpty(((SAPbouiCOM.EditText)MTXATTCH.Columns.Item("CLATTACH").Cells.Item(lastRow).Specific).Value);
                if (needNewRow)
                {
                    Global.GFunc.SetNewLine(MTXATTCH, DBDataSourceLine, 1, "");
                    lastRow = MTXATTCH.VisualRowCount;
                }

                ((SAPbouiCOM.EditText)MTXATTCH.Columns.Item("CLATTACH").Cells.Item(lastRow).Specific).Value = filePath;
                MTXATTCH.FlushToDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
            }

        }

        private void CBCMSTAT_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.ComboBox oCmb = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBCMSTAT").Specific;

                string selectedValue = "";
                if (oCmb.Selected != null)
                    selectedValue = oCmb.Selected.Value.Trim();

                if (selectedValue == "C")
                {
                    int result = Application.SBO_Application.MessageBox(
                        "Are you sure want to change?",
                        1,
                        "Yes",
                        "No"
                    );

                    if (result == 1)
                    {
                        oForm.Items.Item("CBCMSTAT").Enabled = false;
                    }
                    else
                    {
                        oCmb.Select("D", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error in CM Status selection: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }

        }

        private void CBMRSTAT_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.ComboBox oCmb = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBMRSTAT").Specific;

                string selectedValue = "";
                if (oCmb.Selected != null)
                    selectedValue = oCmb.Selected.Value.Trim();

                if (selectedValue == "C")
                {
                    int result = Application.SBO_Application.MessageBox(
                        "Are you sure want to change?",
                        1,
                        "Yes",
                        "No"
                    );

                    if (result == 1)
                    {
                        oForm.Items.Item("CBMRSTAT").Enabled = false;
                    }
                    else
                    {
                        oCmb.Select("D", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error in MR Status selection: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }

        }

        private void CBDSNBNK_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.ComboBox oCmb = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBDSNBNK").Specific;

                string selectedValue = "";
                if (oCmb.Selected != null)
                    selectedValue = oCmb.Selected.Value.Trim();

                if (selectedValue == "Y")
                {
                    int result = Application.SBO_Application.MessageBox(
                        "Are you sure want to change?",
                        1,
                        "Yes",
                        "No"
                    );

                    if (result == 1)
                    {
                        oForm.Items.Item("CBDSNBNK").Enabled = false;
                    }
                    else
                    {
                        oCmb.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error in Doc Send to Bank selection: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
        }

        private void ETSCNO_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    string code = ((SAPbouiCOM.EditText)oForm.Items.Item("ETSCNO").Specific).Value.Trim();
                    string UCode = Global.GFunc.ToUpperCase(code);
                    ((SAPbouiCOM.EditText)oForm.Items.Item("ETSCNO").Specific).Value = UCode;
                    if (!string.IsNullOrEmpty(UCode))
                    {
                        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string query = $@"SELECT 1 FROM ""@FIL_DH_OSCM"" WHERE ""U_SCNO"" = '{UCode.Replace("'", "''")}'";
                        oRS.DoQuery(query);
                        if (!oRS.EoF)
                        {
                            Application.SBO_Application.StatusBar.SetText("Code already exists!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            ((SAPbouiCOM.EditText)oForm.Items.Item("ETSCNO").Specific).Value = "";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Error: Sales Contract Code  " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }


        }

        private void ETCURR_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("CurrCode", 0).ToString().Trim();
            SAPbouiCOM.EditText ETBCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETCURR").Specific;
            ETBCD.Value = Code;

        }

        private void ETOBNKAC_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Account = dt.GetValue("Account", 0).ToString().Trim();
            string Bank = dt.GetValue("BankCode", 0).ToString().Trim();

            SAPbouiCOM.EditText ETBAC = (SAPbouiCOM.EditText)oForm.Items.Item("ETOBNKAC").Specific;
            ETBAC.Value = Account;
            SAPbouiCOM.EditText ETBCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETOWNBNK").Specific;
            ETBCD.Value = Bank;

        }

        private void ETCBNKAC_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Account = dt.GetValue("Account", 0).ToString().Trim();
            string Bank = dt.GetValue("BankCode", 0).ToString().Trim();

            SAPbouiCOM.EditText ETBAC = (SAPbouiCOM.EditText)oForm.Items.Item("ETCBNKAC").Specific;
            ETBAC.Value = Account;
            SAPbouiCOM.EditText ETBCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETCUSBNK").Specific;
            ETBCD.Value = Bank;

        }

        private void ETCBNKAC_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

                string customerCode = ((SAPbouiCOM.EditText)oForm.Items.Item("ETCUSTMR").Specific).Value.Trim();

                if (string.IsNullOrEmpty(customerCode))
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Please select Customer first.",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Error
                    );
                    BubbleEvent = false;
                    return;
                }

                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_OCRB")
                {
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(cflUID);

                    SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                    SAPbouiCOM.Condition oCon = oCons.Add();

                    oCon.Alias = "CardCode";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = customerCode;

                    oCFL.SetConditions(oCons);
                }
            }
            catch (Exception ex)
            {   
                Application.SBO_Application.StatusBar.SetText(
                    "Error filtering Bank Account CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }
        }

        private void ETBRNDCD_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("Code", 0).ToString().Trim();
            string Name = dt.GetValue("Name", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETBRNDCD").Specific;
            ETCD.Value = Code;
            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETBRNDNM").Specific;
            ETNM.Value = Name;

        }
        private void ETBRNDCD_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_BRND")
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(cflUID);
                    SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                    SAPbouiCOM.Condition oCon1 = oCons.Add();
                    oCon1.Alias = "U_ACTIVE";
                    oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon1.CondVal = "Y";
                    oCFL.SetConditions(oCons);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error filtering Brand CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }
        }

        private void ETCUSTMR_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string cflUID = cflArg.ChooseFromListUID;

                if (cflUID == "CFL_OCRD")
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(cflUID);
                    SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                    SAPbouiCOM.Condition oCon1 = oCons.Add();
                    oCon1.Alias = "CardType";
                    oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon1.CondVal = "C";
                    oCFL.SetConditions(oCons);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error filtering Brand CFL: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
                BubbleEvent = false;
            }

        }

        private void ETCUSTMR_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ISBOChooseFromListEventArg cflArg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            SAPbouiCOM.DataTable dt = cflArg.SelectedObjects;
            if (dt == null || dt.Rows.Count == 0)
                return;

            string Code = dt.GetValue("CardCode", 0).ToString().Trim();
            string Name = dt.GetValue("CardName", 0).ToString().Trim();

            SAPbouiCOM.EditText ETCD = (SAPbouiCOM.EditText)oForm.Items.Item("ETCUSTMR").Specific;
            ETCD.Value = Code;

            SAPbouiCOM.EditText ETNM = (SAPbouiCOM.EditText)oForm.Items.Item("ETCUSTNM").Specific;
            ETNM.Value = Name;


            SAPbouiCOM.Item CBSHPADD = oForm.Items.Item("CBSHPADD");
            CBSHPADD.Enabled = true;

            //Shipping Type
            string address = $@"SELECT ""Address"", ""Street"" FROM ""CRD1"" where ""CardCode""='{Code}'";
            SAPbouiCOM.ComboBox CBADD= (SAPbouiCOM.ComboBox)oForm.Items.Item("CBSHPADD").Specific;
            Global.GFunc.setComboBoxValue(CBADD, address);
        }

        
    }
}
