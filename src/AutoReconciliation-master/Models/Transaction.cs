using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReconciliation.Models
{
    class Transaction
    {
        public int transID;
        public int transRowID;
        // True = Debit, False = Credit
        public bool creditOrDebit;
        public double reconcileAmount;
        public double calculatedReconcileAmount;
        public bool isNegative;
        public Transaction(int transID, int transRowID, bool creditOrDebit, double reconcileAmount)
        {
            this.transID = transID;
            this.transRowID = transRowID;
            this.creditOrDebit = creditOrDebit;
            this.reconcileAmount = reconcileAmount;
            if (this.reconcileAmount < 0)
            {
                this.isNegative = true;
                this.creditOrDebit = !this.creditOrDebit;
                this.reconcileAmount *= -1;
            }
        }

        public void addReconcileAmount(double amount)
        {
            if (isNegative)
            {
                calculatedReconcileAmount -= amount;
            }
            else
            {
                calculatedReconcileAmount += amount;
            }
        }
    }
}
