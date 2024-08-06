using AutoReconciliation.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReconciliation.Services
{
    class ReconciliationService
    {
        SAPbobsCOM.Company oCom;
        SAPbobsCOM.InternalReconciliationsService service;

        public ReconciliationService(SAPbobsCOM.Company oCom)
        {
            this.oCom = oCom;
            service = (SAPbobsCOM.InternalReconciliationsService)this.oCom.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.InternalReconciliationsService);
        }
        public SAPbobsCOM.InternalReconciliationsService GetService()
        {
            return service;
        }

        public SAPbobsCOM.InternalReconciliationOpenTrans GetOpenTransactions(SAPbobsCOM.InternalReconciliationOpenTransParams transactionParams)
        {
            return service.GetOpenTransactions(transactionParams);
        }

        public SAPbobsCOM.InternalReconciliationOpenTransParams GetTransParams(string cardCode, string dateFrom, string dateTo)
        {
            SAPbobsCOM.InternalReconciliationOpenTransParams transParams = (SAPbobsCOM.InternalReconciliationOpenTransParams)service.GetDataInterface(SAPbobsCOM.InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTransParams);

            transParams.ReconDate = DateTime.Today;
            transParams.DateType = SAPbobsCOM.ReconSelectDateTypeEnum.rsdtPostDate;
            transParams.FromDate = DateTime.ParseExact(dateFrom, "yyyyMMdd", null);
            transParams.ToDate = DateTime.ParseExact(dateTo, "yyyyMMdd", null);

            transParams.CardOrAccount = SAPbobsCOM.CardOrAccountEnum.coaCard;
            transParams.InternalReconciliationBPs.Add();
            transParams.InternalReconciliationBPs.Item(0).BPCode = cardCode;

            return transParams;
        }
        public SAPbobsCOM.InternalReconciliationOpenTransParams GetLinkedTransParams(string cardCode, string linkedCardCode, string dateFrom, string dateTo)
        {
            SAPbobsCOM.InternalReconciliationOpenTransParams transParams = (SAPbobsCOM.InternalReconciliationOpenTransParams)service.GetDataInterface(SAPbobsCOM.InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTransParams);

            transParams.ReconDate = DateTime.Today;
            transParams.DateType = SAPbobsCOM.ReconSelectDateTypeEnum.rsdtPostDate;
            transParams.FromDate = DateTime.ParseExact(dateFrom, "yyyyMMdd", null);
            transParams.ToDate = DateTime.ParseExact(dateTo, "yyyyMMdd", null);

            transParams.CardOrAccount = SAPbobsCOM.CardOrAccountEnum.coaCard;
            transParams.InternalReconciliationBPs.Add();
            transParams.InternalReconciliationBPs.Item(0).BPCode = cardCode;
            transParams.InternalReconciliationBPs.Add();
            transParams.InternalReconciliationBPs.Item(1).BPCode = linkedCardCode;

            return transParams;
        }
        public SAPbobsCOM.InternalReconciliationParams Reconciliate(SAPbobsCOM.InternalReconciliationOpenTrans openTransactions, Dictionary<(int,int), double> reconciliatedTransactions)
        {
            bool needUpdate = false;
            foreach (SAPbobsCOM.InternalReconciliationOpenTransRow row in openTransactions.InternalReconciliationOpenTransRows)
            {
                var id = (row.TransId, row.TransRowId);
                if (reconciliatedTransactions.ContainsKey(id))
                {
                    needUpdate = true;
                    row.Selected = SAPbobsCOM.BoYesNoEnum.tYES;
                    row.ReconcileAmount = reconciliatedTransactions[id];
                }
            }
            return needUpdate ? service.Add(openTransactions) : null;
        }
    }
}
