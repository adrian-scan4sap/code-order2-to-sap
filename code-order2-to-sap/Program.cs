using SAPbobsCOM;
using System;
using System.Runtime.InteropServices;

namespace code_order2_to_sap
{
    internal class Program
    {
        /// <summary>
        /// Main Test Method
        /// </summary>
        /// <param name="args">Not used</param>
        static void Main(string[] args)
        {
            /* Declare the company variable - the connection */
            Company company = null;

            Console.WriteLine("Connecting to SAP...");

            /* Connect returns if connection has been established as bool */
            var isConnected = Connect(ref company);

            Console.WriteLine(Environment.NewLine + "Connected; adding order now...");

            var additionResult = AddSalesOrder(company);

            Console.WriteLine(string.Format("{0}Addition to SAP is Successful = [{1}] and Message is = [{2}]", Environment.NewLine, additionResult.Item1, additionResult.Item2));

            if (additionResult.Item1)
            {
                var sapComSalesOrder = (Documents)company.GetBusinessObject(BoObjectTypes.oOrders);
                sapComSalesOrder.GetByKey(int.Parse(company.GetNewObjectKey()));

                // Post Down Payment Invoice to SAP
                var isAdditionSuccessful = AddLinkedDownPaymentInvoice(company, sapComSalesOrder);

                if (!isAdditionSuccessful)
                {
                    /* Get SAP error in case of unsuccessful addition | Down Payment Invoice */
                    Console.WriteLine(string.Format("Sales Order was added successfully but the Down Payment Invoice had an issue: {0}", company.GetLastErrorDescription()));
                }
                else
                {
                    // Post Incoming Payment to SAP
                    isAdditionSuccessful = AddLinkedIncomingPayment(company, company.GetNewObjectKey(), sapComSalesOrder.CardCode, sapComSalesOrder.DocTotal, sapComSalesOrder.DocCurrency, "PAY-123", "Payment for Sales Order " + sapComSalesOrder.DocNum);

                    if (!isAdditionSuccessful)
                    {
                        /* Get SAP error in case of unsuccessful addition | Incoming Payment */
                        Console.WriteLine(string.Format("Sales Order + Down Payment Invoice were added successfully but the Incoming Payment had an issue: {0}", company.GetLastErrorDescription()));
                    }
                }
            }

            /* Disconnect also released the held memory */
            Disconnect(ref company);

            Console.WriteLine(Environment.NewLine + "Disconnected. Press any key to exit.");
            Console.ReadKey();
        }

        #region Sales Order

        /// <summary>
        /// Attempts to add a sales order to SAP
        /// </summary>
        /// <param name="company">The SAP Connection</param>
        /// <returns>[True, ""] if sales order added, [False, "sap-error-message"] in case of error</returns>
        static Tuple<bool, string> AddSalesOrder(Company company)
        {
            /* New Order Instance */
            Documents sapComSalesOrder = (Documents)company.GetBusinessObject(BoObjectTypes.oOrders);

            /* Header Fields */
            SetOrderHeader(sapComSalesOrder);

            /* Order Lines */
            SetOrderLines(sapComSalesOrder);

            /* Address Setup */
            SetOrderAddresses(sapComSalesOrder);

            /* Add Freight Charges */
            SetOrderFreight(sapComSalesOrder);

            var operationMessage = "";

            // Post Sales Order to SAP
            var isAdditionSuccessful = sapComSalesOrder.Add() == 0;

            if (!isAdditionSuccessful)
            {
                /* Get SAP error in case of unsuccessful addition | Sales Order */
                operationMessage = company.GetLastErrorDescription();
            }

            return new Tuple<bool, string>(isAdditionSuccessful, operationMessage);
        }
                
        /// <summary>
        /// Sets the sales order header fields
        /// </summary>
        /// <param name="sapComSalesOrder">The SAP order instance to set it on</param>
        static void SetOrderHeader(Documents sapComSalesOrder)
        {
            sapComSalesOrder.CardCode = "C20000";
            sapComSalesOrder.DocDate = DateTime.Now;
            sapComSalesOrder.DocDueDate = DateTime.Now.AddDays(1);
            sapComSalesOrder.DocTotal = 34;
            sapComSalesOrder.DocCurrency = "$";
        }

        /// <summary>
        /// Sets order lines
        /// </summary>
        /// <param name="sapComSalesOrder">The SAP order instance to set them on</param>
        static void SetOrderLines(Documents sapComSalesOrder)
        {
            /* Setting two lines */
            var sapComSalesOrderLine = sapComSalesOrder.Lines;
            SetOrderLastLine(sapComSalesOrderLine, "A00001", "Different Description than SAP's", 2, 10, "$", "01", "CA");
            sapComSalesOrderLine.Add();
            SetOrderLastLine(sapComSalesOrderLine, "A00002", "Different Description than SAP's 2", 1, 5, "$", "01", "CA");
        }

        /// <summary>
        /// Sets the last line of the sales order according to the provided parameters
        /// </summary>
        /// <param name="sapComSalesOrderLine">The SAP order line instance to set it on</param>
        /// <param name="itemCode">Line: Item Code</param>
        /// <param name="description">Line: Description</param>
        /// <param name="quantity">Line: Quantity</param>
        /// <param name="unitPrice">Line: Unit Price</param>
        /// <param name="currency">Line: Currency Code</param>
        /// <param name="warehouseCode">Line: WhsCode</param>
        /// <param name="vatCode">Line: VAT Code</param>
        static void SetOrderLastLine(Document_Lines sapComSalesOrderLine, string itemCode, string description, double quantity, double unitPrice, string currency, string warehouseCode, string vatCode)
        {
            sapComSalesOrderLine.SetCurrentLine(sapComSalesOrderLine.Count - 1);
            sapComSalesOrderLine.ItemCode = itemCode;
            sapComSalesOrderLine.ItemDescription = description;            
            sapComSalesOrderLine.Quantity = quantity;
            sapComSalesOrderLine.UnitPrice = unitPrice;
            sapComSalesOrderLine.Currency = currency;
            sapComSalesOrderLine.WarehouseCode = warehouseCode;
            sapComSalesOrderLine.VatGroup = vatCode;
        }

        /// <summary>
        /// Sets the sales order addresses
        /// </summary>
        /// <param name="sapComSalesOrder">The SAP order instance to set it on</param>
        static void SetOrderAddresses(Documents sapComSalesOrder)
        {
            /* Addresses */
            sapComSalesOrder.AddressExtension.BillToStreet = "Billing";
            sapComSalesOrder.AddressExtension.BillToStreetNo = "Clockhouse Place";
            sapComSalesOrder.AddressExtension.BillToBuilding = "Bedfond Road";
            sapComSalesOrder.AddressExtension.BillToCity = "Feltham";
            sapComSalesOrder.AddressExtension.BillToCountry = "GB";
            sapComSalesOrder.AddressExtension.BillToZipCode = "TW14 8HD";

            sapComSalesOrder.AddressExtension.ShipToStreet = "Shipping";
            sapComSalesOrder.AddressExtension.ShipToStreetNo = "Clockhouse Place";
            sapComSalesOrder.AddressExtension.ShipToBuilding = "Bedfond Road";
            sapComSalesOrder.AddressExtension.ShipToCity = "Feltham";            
            sapComSalesOrder.AddressExtension.ShipToCountry = "GB";
            sapComSalesOrder.AddressExtension.ShipToZipCode = "TW14 8HD";
        }

        /// <summary>
        /// Sets Freight Charges on the SAP Document (the sales order)
        /// </summary>
        /// <param name="sapComSalesOrder">The SAP order instance to set it on</param>
        static void SetOrderFreight(Documents sapComSalesOrder)
        {
            sapComSalesOrder.Expenses.Remarks = "Manual Remark";
            sapComSalesOrder.Expenses.ExpenseCode = 1;
            sapComSalesOrder.Expenses.VatGroup = "CA";
            sapComSalesOrder.Expenses.TaxCode = "CA";
            sapComSalesOrder.Expenses.LineTotal = 4;
            sapComSalesOrder.Expenses.DistributionMethod = BoAdEpnsDistribMethods.aedm_RowTotal;
        }

        #endregion

        #region Down Payment Invoice

        /// <summary>
        /// Attempts to add a Down Payment Invoice to SAP
        /// </summary>
        /// <param name="company">The SAP Connection</param>
        /// <param name="sapComSalesOrder">The Sales Order to link the Down Payment Invoice to</param>
        /// <returns>[True] if successful, [False] otherwise</returns>
        static bool AddLinkedDownPaymentInvoice(Company company, Documents sapComSalesOrder)
        {
            Documents sapComDownPaymentDocument = (Documents)company.GetBusinessObject(BoObjectTypes.oDownPayments);
            sapComDownPaymentDocument.DownPaymentType = DownPaymentTypeEnum.dptInvoice;

            SetDownPaymentInvoiceHeader(sapComDownPaymentDocument, sapComSalesOrder);
            SetDownPaymentInvoiceLines(sapComDownPaymentDocument.Lines, sapComSalesOrder.Lines);

            var isDownPaymentDocumentAdditionSuccessful = sapComDownPaymentDocument.Add() == 0;

            return isDownPaymentDocumentAdditionSuccessful;
        }

        /// <summary>
        /// Sets the header of the Down Payment Invoice
        /// </summary>
        /// <param name="sapComDownPaymentInvoice">The Down Payment Invoice for the header to be set</param>
        /// <param name="sapComSalesOrder">The Sales Order to link the Down Payment Invoice to</param>
        static void SetDownPaymentInvoiceHeader(Documents sapComDownPaymentInvoice, Documents sapComSalesOrder)
        {
            sapComDownPaymentInvoice.CardCode = sapComSalesOrder.CardCode;
            sapComDownPaymentInvoice.DocDate = sapComSalesOrder.DocDate;
            sapComDownPaymentInvoice.DocDueDate = sapComSalesOrder.DocDueDate;
            sapComDownPaymentInvoice.DocCurrency = sapComSalesOrder.DocCurrency;
            sapComDownPaymentInvoice.DocTotal = sapComSalesOrder.DocTotal;
        }

        /// <summary>
        /// Sets the lines of the Down Payment Invoice
        /// </summary>
        /// <param name="sapComDownPaymentInvoiceLines">>The Down Payment Invoice Lines to be set</param>
        /// <param name="sapComSalesOrderLines">The Sales Order to link the Down Payment Invoice to</param>
        static void SetDownPaymentInvoiceLines(Document_Lines sapComDownPaymentInvoiceLines, Document_Lines sapComSalesOrderLines)
        {
            var newRowNeeded = false;

            for (var orderLineIndex = 0; orderLineIndex < sapComSalesOrderLines.Count; orderLineIndex++)
            {
                if (newRowNeeded) { sapComDownPaymentInvoiceLines.Add(); }
                else { newRowNeeded = true; }

                sapComSalesOrderLines.SetCurrentLine(orderLineIndex);
                sapComDownPaymentInvoiceLines.SetCurrentLine(sapComDownPaymentInvoiceLines.Count - 1);                

                sapComDownPaymentInvoiceLines.BaseEntry = sapComSalesOrderLines.DocEntry;
                sapComDownPaymentInvoiceLines.BaseType = 17;
                sapComDownPaymentInvoiceLines.BaseLine = sapComSalesOrderLines.LineNum;

                sapComDownPaymentInvoiceLines.Quantity = sapComSalesOrderLines.Quantity;
                sapComDownPaymentInvoiceLines.UnitPrice = sapComSalesOrderLines.UnitPrice;
                sapComDownPaymentInvoiceLines.PriceAfterVAT = sapComSalesOrderLines.PriceAfterVAT;
                sapComDownPaymentInvoiceLines.Currency = sapComSalesOrderLines.Currency;
                sapComDownPaymentInvoiceLines.VatGroup = sapComSalesOrderLines.VatGroup;
            }
        }

        #endregion

        #region Incoming Payment

        /// <summary>
        /// Attemts to add a linked Incoming Payment to the Down Payment Invoice
        /// </summary>
        /// <param name="company">The SAP Connection</param>
        /// <param name="downPaymentInvoiceDocEntry">The Down Payment Invoice DocEntry to have the Incoming Payment to be linkes to</param>
        /// <param name="cardCode">Customer code</param>
        /// <param name="amount">Payment amount</param>
        /// <param name="currency">Payment Currency</param>
        /// <param name="paymentReference">Payment Reference</param>
        /// <param name="notes">Incoming Payment Remarks</param>
        /// <returns>[True] if addition is successful, [False] otherwise.</returns>
        static bool AddLinkedIncomingPayment(Company company, string downPaymentInvoiceDocEntry, string cardCode, double amount, string currency, string paymentReference, string notes)
        {
            try
            {
                var incomingPayment = (Payments)company.GetBusinessObject(BoObjectTypes.oIncomingPayments);

                incomingPayment.DocType = BoRcptTypes.rCustomer;
                incomingPayment.IsPayToBank = BoYesNoEnum.tNO;
                incomingPayment.CardCode = cardCode;
                incomingPayment.DocDate = DateTime.Now;
                incomingPayment.DueDate = DateTime.Now;
                incomingPayment.DocCurrency = currency;

                incomingPayment.TransferSum = amount;
                incomingPayment.TransferAccount = "_SYS00000000081";
                incomingPayment.TransferReference = paymentReference;

                incomingPayment.Invoices.InvoiceType = BoRcptInvTypes.it_DownPayment;
                incomingPayment.Invoices.DocEntry = int.Parse(downPaymentInvoiceDocEntry);
                incomingPayment.Invoices.SumApplied = amount;
                incomingPayment.Invoices.DiscountPercent = 0;

                incomingPayment.Remarks = notes;

                return incomingPayment.Add() == 0;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Connect / Disconnect to/from SAP

        /// <summary>
        /// Connects to the provided company.
        /// </summary>
        /// <param name="company">Provide uninstantiated</param>
        /// <returns>True if connection was extablished and False if connection could not be done</returns>
        static bool Connect(ref Company company)
        {
            if (company == null)
            {
                company = new Company();
            }

            if (!company.Connected)
            {
                /* Server connection details */
                company.Server = "sql-server-name";
                company.DbServerType = BoDataServerTypes.dst_MSSQL2016;
                company.DbUserName = "sa";
                company.DbPassword = "sql-password";
                company.UseTrusted = false;

                /* SAP connection details: DB, SAP User and SAP Password */
                company.CompanyDB = "SAP-database-name";
                company.UserName = "SAP-user";
                company.Password = "SAP-password";

                /* In case the SAP license server is kept in a different location (in most cases can be left empty) */
                company.LicenseServer = "";

                var isSuccessful = company.Connect() == 0;

                return isSuccessful;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Disconnects and releases the held memory (RAM)
        /// </summary>
        /// <param name="company"></param>
        static void Disconnect(ref Company company)
        {
            if (company != null
                && company.Connected)
            {
                company.Disconnect();

                Release(ref company);
            }
        }

        /// <summary>
        /// Re-usable method for releasing COM-held memory
        /// </summary>
        /// <typeparam name="T">Type of object to be released</typeparam>
        /// <param name="obj">The instance to be released</param>
        static void Release<T>(ref T obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                }
            }
            catch (Exception) { }
            finally
            {
                obj = default(T);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        #endregion
    }
}
