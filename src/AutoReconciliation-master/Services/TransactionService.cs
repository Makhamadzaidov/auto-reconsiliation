using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoReconciliation.Models;

namespace AutoReconciliation.Services
{
    class TransactionService
    {
        List<Transaction> transactions;
        List<Transaction> reconciliatedTransactions;
        Queue<Transaction> debitTransactions;
        Queue<Transaction> creditTransactions;
        public TransactionService()
        {
            transactions = new List<Transaction>();
            reconciliatedTransactions = new List<Transaction>();
            debitTransactions = new Queue<Transaction>();
            creditTransactions = new Queue<Transaction>();
        }

        public void AddTransaction(Transaction t)
        {
            transactions.Add(t);
        }
        public void RemoveTransaction(Transaction t)
        {
            transactions.Remove(t);
        }
        public void ClearTransactions()
        {
            transactions.Clear();
        }
        public void AutoReconciliate()
        {
            debitTransactions.Clear();
            creditTransactions.Clear();
            reconciliatedTransactions.Clear();
            foreach (var t in transactions)
            {
                if (t.creditOrDebit)
                {
                    debitTransactions.Enqueue(t);
                }
                else
                {
                    creditTransactions.Enqueue(t);
                }
            }
            
            if (debitTransactions.Count == 0 || creditTransactions.Count == 0)
            {
                return;
            }
           
            Transaction debitT = debitTransactions.Dequeue();
            Transaction creditT = creditTransactions.Dequeue();
            while(true) {
                var debitAmount = debitT.reconcileAmount - debitT.calculatedReconcileAmount;
                var creditAmount = creditT.reconcileAmount - creditT.calculatedReconcileAmount;
                if (debitAmount > creditAmount)
                {
                    debitT.addReconcileAmount(creditAmount);
                    creditT.addReconcileAmount(creditAmount);
                    reconciliatedTransactions.Add(creditT);
                    if (creditTransactions.Count == 0)
                    {
                        reconciliatedTransactions.Add(debitT);
                        break;
                    }
                    creditT = creditTransactions.Dequeue();
                }
                else if (debitAmount < creditAmount)
                {
                    creditT.addReconcileAmount(debitAmount);
                    debitT.addReconcileAmount(debitAmount);
                    reconciliatedTransactions.Add(debitT);
                    if (debitTransactions.Count == 0)
                    {
                        reconciliatedTransactions.Add(creditT);
                        break;
                    }
                    debitT = debitTransactions.Dequeue();
                }
                else
                {
                    debitT.addReconcileAmount(debitAmount);
                    creditT.addReconcileAmount(creditAmount);
                    reconciliatedTransactions.Add(debitT);
                    reconciliatedTransactions.Add(creditT);
                    if (debitTransactions.Count == 0 || creditTransactions.Count == 0) break;
                    debitT = debitTransactions.Dequeue();
                    creditT = creditTransactions.Dequeue();
                }
            }
        }

        public Dictionary<(int, int), double> GetReconciliatedTransactions()
        {
            Dictionary<(int, int), double> result = new Dictionary<(int, int), double>();
            foreach (var t in reconciliatedTransactions)
            {
                result.Add((t.transID, t.transRowID), t.calculatedReconcileAmount);
            }
            return result;
        }
    }
}
