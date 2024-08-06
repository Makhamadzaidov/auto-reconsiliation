using AutoReconciliation.Models;
using AutoReconciliation.Services;
using SAPbouiCOM.Framework;
using System;
using System.IO;
using System.Xml.Serialization;

namespace AutoReconciliation
{
    [FormAttribute("AutoReconciliation.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        private SAPbouiCOM.EditText bpCode;
        private SAPbouiCOM.EditText dateFrom;
        private SAPbouiCOM.EditText dateTo;
        private SAPbouiCOM.EditText resultEdittext;
        private SAPbouiCOM.CheckBox isGroupCheckbox;
        private SAPbouiCOM.ComboBox groupsCombobox;
        private SAPbouiCOM.Button ReconcileButton;


        SAPbouiCOM.Application oApp;
        SAPbobsCOM.Company oCom;
        TransactionService transactionService;
        ReconciliationService reconciliationService;
        ReportService reportService;
        SAPbobsCOM.Recordset oRS;
        SAPbobsCOM.BusinessPartners oBP;

        public Form1(SAPbouiCOM.Application oApp, SAPbobsCOM.Company oCom, SAPbobsCOM.Recordset oRS)
        {
            this.oApp = oApp;
            this.oCom = oCom;
            this.oRS = oRS;
            oBP = (SAPbobsCOM.BusinessPartners)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
            transactionService = new TransactionService();
            reconciliationService = new ReconciliationService(oCom);
            reportService = new ReportService();
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.GetItem("rec_btn").AffectsFormMode = false;
            this.GetItem("bp_code").AffectsFormMode = false;
            this.GetItem("from_date").AffectsFormMode = false;
            this.GetItem("to_date").AffectsFormMode = false;
            this.GetItem("Item_0").AffectsFormMode = false;
            this.GetItem("Item_9").AffectsFormMode = false;
            this.GetItem("Item_10").AffectsFormMode = false;

            this.ReconcileButton = ((SAPbouiCOM.Button)(this.GetItem("rec_btn").Specific));
            this.ReconcileButton.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Reconcile_ClickBefore);
            this.bpCode = ((SAPbouiCOM.EditText)(this.GetItem("bp_code").Specific));
            this.dateFrom = ((SAPbouiCOM.EditText)(this.GetItem("from_date").Specific));
            this.dateTo = ((SAPbouiCOM.EditText)(this.GetItem("to_date").Specific));
            this.resultEdittext = ((SAPbouiCOM.EditText)(this.GetItem("Item_0").Specific));
            this.isGroupCheckbox = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_9").Specific));
            this.groupsCombobox = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_10").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }


        private void OnCustomInitialize()
        {
        }

        private void Reconcile_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (isGroupCheckbox.Checked)
            {
                if (groupsCombobox.Value == "" || dateFrom.Value == "" || dateTo.Value == "")
                {
                    oApp.StatusBar.SetText("Заполните все поля.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return;
                }
                oApp.StatusBar.SetText("Аддон начал выверку. Это может занять некоторое время...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                oRS.DoQuery($"SELECT \"CardCode\" FROM OCRD WHERE \"CardType\" = '{groupsCombobox.Value}'");
                BusinessPartners businessPartners;
                XmlSerializer xmlReader = new XmlSerializer(typeof(BusinessPartners));
                using (StringReader sr = new StringReader(oRS.GetAsXML()))
                {
                    businessPartners = (BusinessPartners)xmlReader.Deserialize(sr);
                }

                string response;
                bool errorOccured = false;
                resultEdittext.Value = "";

                foreach (var bp in businessPartners.BO.OCRD)
                {
                    oBP.GetByKey(bp.CardCode);
                    string cardCode = bp.CardCode;
                    string linkedCardCode = oBP.LinkedBusinessPartner;
                    string xml = "";

                    SAPbobsCOM.InternalReconciliationOpenTransParams transParams;
                    SAPbobsCOM.InternalReconciliationOpenTrans openTrans;
                    transactionService.ClearTransactions();
                    try
                    {
                        if (linkedCardCode == "")
                        {
                            transParams = reconciliationService.GetTransParams(cardCode, dateFrom.Value, dateTo.Value);
                        }
                        else
                        {
                            transParams = reconciliationService.GetLinkedTransParams(cardCode, linkedCardCode, dateFrom.Value, dateTo.Value);
                        }

                        openTrans = reconciliationService.GetOpenTransactions(transParams);
                        xml = openTrans.ToXMLString();

                        foreach (SAPbobsCOM.InternalReconciliationOpenTransRow row in openTrans.InternalReconciliationOpenTransRows)
                        {
                            transactionService.AddTransaction(new Transaction(row.TransId, row.TransRowId, row.CreditOrDebit != 0, row.ReconcileAmount));
                        }
                        transactionService.AutoReconciliate();
                        var reconciliatedTransactions = transactionService.GetReconciliatedTransactions();
                        reconciliationService.Reconciliate(openTrans, reconciliatedTransactions);

                        response = resultEdittext.Value;
                        response = linkedCardCode == "" ? $"Выверка для [{cardCode}] успешно завершена.\r\n" + response : $"Выверка для [{cardCode}] и [{linkedCardCode}] успешно завершена.\r\n" + response;
                        resultEdittext.Value = response;
                    }
                    catch (Exception e)
                    {
                        errorOccured = true;
                        response = resultEdittext.Value;
                        response = $"Ошибка при выверке для [{bp.CardCode}]: {e}\r\n" + response;
                        resultEdittext.Value = response;

                        reportService.ReportError(cardCode, linkedCardCode, e, xml);
                    }
                }
                if (!errorOccured)
                {
                    oApp.StatusBar.SetText($"Автоматическая выверка успешно завершена.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                else
                {
                    oApp.StatusBar.SetText($"Произошла ошибка при выверке некторых бизнес-партнеров. Просмотрите логи. Процесс выверки успешно завершена.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }
            else
            {
                if (bpCode.Value == "" || dateFrom.Value == "" || dateTo.Value == "")
                {
                    oApp.StatusBar.SetText("Заполните все поля.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return;
                }

                oApp.StatusBar.SetText("Аддон начал выверку. Пожалуйста подожите...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                resultEdittext.Value = "";

                oBP.GetByKey(bpCode.Value);
                string cardCode = bpCode.Value;
                string linkedCardCode = oBP.LinkedBusinessPartner;
                string xml = "";
                SAPbobsCOM.InternalReconciliationOpenTransParams transParams;
                SAPbobsCOM.InternalReconciliationOpenTrans openTrans;
                try
                {
                    if (oBP.LinkedBusinessPartner == "")
                    {
                        transParams = reconciliationService.GetTransParams(cardCode, dateFrom.Value, dateTo.Value);
                    }
                    else
                    {
                        transParams = reconciliationService.GetLinkedTransParams(cardCode, linkedCardCode, dateFrom.Value, dateTo.Value);
                    }

                    openTrans = reconciliationService.GetOpenTransactions(transParams);
                    xml = openTrans.ToXMLString();
                    transactionService.ClearTransactions();
                    foreach (SAPbobsCOM.InternalReconciliationOpenTransRow row in openTrans.InternalReconciliationOpenTransRows)
                    {
                        transactionService.AddTransaction(new Transaction(row.TransId, row.TransRowId, row.CreditOrDebit != 0, row.ReconcileAmount));
                    }
                    transactionService.AutoReconciliate();
                    var reconciliatedTransactions = transactionService.GetReconciliatedTransactions();
                    reconciliationService.Reconciliate(openTrans, reconciliatedTransactions);

                    resultEdittext.Value = linkedCardCode == "" ? $"Выверка для [{cardCode}] успешно завершена.\r\n" : $"Выверка для [{cardCode}] и [{linkedCardCode}] успешно завершена.\r\n";
                    oApp.StatusBar.SetText($"Автоматическая выверка успешно завершена.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                catch (Exception e)
                {
                    resultEdittext.Value = $"Ошибка при выверке для [{cardCode}]: {e}";
                    oApp.StatusBar.SetText($"Произошла ошибка при автоматической выверке: {e}", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                    reportService.ReportError(cardCode, linkedCardCode, e, xml);
                }
            }
        }
    }
}