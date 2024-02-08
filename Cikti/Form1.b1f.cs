using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using SAPbobsCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace Cikti
{
    [FormAttribute("Cikti.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>asdas
        public override void OnInitializeComponent()
        {
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_4").Specific));
            this.Folder9 = ((SAPbouiCOM.Folder)(this.GetItem("Item_20").Specific));
            this.Folder10 = ((SAPbouiCOM.Folder)(this.GetItem("Item_21").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_22").Specific));
            this.Button1.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button1_ClickBefore);
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("Item_23").Specific));
            this.EditText7 = ((SAPbouiCOM.EditText)(this.GetItem("Item_24").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("Item_25").Specific));
            this.EditText9 = ((SAPbouiCOM.EditText)(this.GetItem("Item_26").Specific));
            this.EditText10 = ((SAPbouiCOM.EditText)(this.GetItem("Item_27").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_28").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_29").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_30").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_31").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_32").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_33").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_34").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_37").Specific));
            this.Button2.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button2_ClickBefore);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }

        private void OnCustomInitialize()
        {

        }
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.Folder Folder9;
        private SAPbouiCOM.Folder Folder10;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.EditText EditText7;
        private SAPbouiCOM.EditText EditText8;
        private SAPbouiCOM.EditText EditText9;
        private SAPbouiCOM.EditText EditText10;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.StaticText StaticText10;
        private SAPbouiCOM.StaticText StaticText11;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Application oApp;
        private SAPbobsCOM.Company oCompany;
        private SAPbouiCOM.Form oForm;
        private bool bFormAktifOldu = false;

        private void Button1_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            Folder9.Select();
            
            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string _Sorgu = "SELECT TOP 10 '' AS \"Seç\" , CASE WHEN \"TransType\" = '30' THEN 'YK' WHEN \"TransType\" = '13' THEN ELSE '' END  AS \"Tipi\" ,  \"TransId\" , '' AS \"Açıklama\" FROM \"OJDT\"  ";
            oRecordset.DoQuery(_Sorgu);

            SAPbouiCOM.Form oForm = this.oApp.Forms.Item(pVal.FormUID);
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_34").Specific;

            for (int i = 1; i < oRecordset.RecordCount; i++)
            {

                oMatrix.AddRow();

                SAPbouiCOM.EditText Sayac = (SAPbouiCOM.EditText)oMatrix.Columns.Item("#").Cells.Item(i).Specific;
                SAPbouiCOM.CheckBox Sec = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific;
                SAPbouiCOM.EditText Tip = (SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(i).Specific;
                SAPbouiCOM.EditText TransId = (SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_2").Cells.Item(i).Specific;
                SAPbouiCOM.EditText Aciklama = (SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_3").Cells.Item(i).Specific;

                Sayac.Value = i.ToString();

                if (oRecordset.Fields.Item("Seç").Value.ToString() == "Y")
                {
                    Sec.Checked = true;
                }
                else
                {
                    Sec.Checked = false;
                }
                
                Tip.Value = oRecordset.Fields.Item("Tipi").Value.ToString();
                TransId.Value = oRecordset.Fields.Item("TransId").Value.ToString();

                //---------------------------------
               

                //---------------------------------
                Aciklama.Value = oRecordset.Fields.Item("Açıklama").Value.ToString();
                oRecordset.MoveNext();

            }
  
        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (!bFormAktifOldu)
                {
                    bFormAktifOldu = true;
                    oApp = (SAPbouiCOM.Application)Application.SBO_Application;
                    oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
                    oForm = (SAPbouiCOM.Form)this.UIAPIRawForm;
                    //oForm = (SAPbouiCOM.Form)oApp.Forms.Item(pVal.FormUID);
                }
            }
            catch (Exception)
            {

            }
        }

        private SAPbouiCOM.Button Button2;

        private void Button2_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            oApp.SetStatusBarMessage("başladı");


            try
            {

                SAPbobsCOM.CompanyService oCmpSrv;
                SAPbobsCOM.ReportLayoutsService oReportLayoutService;
                SAPbobsCOM.ReportLayoutPrintParams oPrintParam;
                oCmpSrv = oCompany.GetCompanyService();
                oReportLayoutService = (ReportLayoutsService)oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService);
                oPrintParam = (ReportLayoutPrintParams)oReportLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutPrintParams);
                oPrintParam.LayoutCode = "JDT20003";
                oPrintParam.DocEntry = 5;
                oReportLayoutService.Print(oPrintParam);
            }
            catch (Exception e)
            {
                oApp.MessageBox(e.ToString());
            }

         

        }
    }
}