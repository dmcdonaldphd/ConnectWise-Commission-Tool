const { chromium } = require('playwright');
const path = require('path');
const readline = require('readline');

const BASE_URL = 'https://na.myconnectwise.net';
const API_VERSION = 'v2025_1';

const PROJECT_ID = 1273;

function waitForEnter(message) {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  return new Promise((resolve) => {
    rl.question(message, () => {
      rl.close();
      resolve();
    });
  });
}

function formatCurrency(value) {
  if (value === null || value === undefined || Number.isNaN(Number(value))) {
    return '';
  }

  return Number(value).toLocaleString('en-US', {
    style: 'currency',
    currency: 'USD',
  });
}

function formatConnectWiseDate(value) {
  if (!value) {
    return '';
  }

  const date = new Date(Number(value));

  if (Number.isNaN(date.getTime())) {
    return String(value);
  }

  return date.toLocaleDateString('en-US');
}

function buildActionMessage(payload, payloadClassName, project) {
  return JSON.stringify({
    payload: JSON.stringify(payload),
    payloadClassName,
    project,
  });
}

async function postConnectWiseAction(
  requestContext,
  endpoint,
  payload,
  payloadClassName,
  project
) {
  const formData = new URLSearchParams();

  formData.set(
    'actionMessage',
    buildActionMessage(payload, payloadClassName, project)
  );

  formData.set('clientTimezoneOffset', '-300');
  formData.set('clientTimezoneName', 'Central Standard Time');

  const url =
    `${BASE_URL}/${API_VERSION}/services/system_io/actionprocessor/${endpoint}.rails`;

  const response = await requestContext.post(url, {
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
      'X-Requested-With': 'XMLHttpRequest',
    },
    data: formData.toString(),
  });

  const responseText = await response.text();

  if (!response.ok()) {
    throw new Error(
      `${endpoint} failed with HTTP ${response.status()}:\n${responseText}`
    );
  }

  let parsedResponse;

  try {
    parsedResponse = JSON.parse(responseText);
  } catch {
    throw new Error(
      `${endpoint} returned invalid JSON:\n${responseText.slice(0, 1000)}`
    );
  }

  if (!parsedResponse.success || !parsedResponse.data?.isSuccess) {
    throw new Error(
      `${endpoint} returned an application error:\n` +
      JSON.stringify(parsedResponse, null, 2)
    );
  }

  return parsedResponse.data.action;
}

async function getProjectInvoices(requestContext, projectId) {
  const payload = {
    requestAllCount: true,
    usePagination: true,
    multilineType: 'BillingML',
    activePage: 1,
    maxResult: 250,
    culture: 'en_US',
    multilineName: 'projectinvoicelist',
    screenId: 'invoice',
    sortDirection: '',
    sortField: '',
    multilineUserParams: [
      {
        paramKey: 'project_recid',
        paramValue: String(projectId),
      },
    ],
    columns: [
      'billing_log_recid',
      'closed',
      'invoice_number',
      'sr_service_recid',
      'company_name',
      'project_name',
      'project_description',
      'po_number',
      'invoice_amount',
      'date_invoice',
      'status_descr',
      'billing_type_desc',
      'project_phase',
      'order_header_recid',
      'owner_level_name',
      'billing_unit',
      'agr_type_desc',
      'internal_note',
      'territory',
      'gl_posted',
      'routed_to',
      'company_id',
      'gl_batchid',
      'invoice_delivery',
      'credit_amount',
      'paid_amount',
      'date_paid',
      'remit_to_name',
    ],
  };

  const action = await postConnectWiseAction(
    requestContext,
    'System/GetMultilineDataAction_projectinvoicelist',
    payload,
    'GetMultilineDataAction',
    'SystemCommon'
  );

  return action.multilineRows.map((item) => JSON.parse(item.row));
}

async function getInvoiceCalculations(requestContext, billingLogRecID) {
  const action = await postConnectWiseAction(
    requestContext,
    'Billing/GetInvoiceCalculationsPodAction',
    {
      billingLogRecID,
      podID: 'invoice_calculations',
    },
    'GetInvoiceCalculationsPodAction',
    'BillingCommon'
  );

  return action.invoiceCalculationsPodViewModel;
}

async function getInvoicePayments(requestContext, billingLogRecID) {
  const action = await postConnectWiseAction(
    requestContext,
    'Billing/GetInvoiceCreditsAndPaymentsPodAction',
    {
      billingLogRecID,
      podID: 'invoice_credits_payments',
    },
    'GetInvoiceCreditsAndPaymentsPodAction',
    'BillingCommon'
  );

  return action.invoiceCreditsAndPaymentsPodViewModel;
}

function determinePaymentStatus(invoice, payments) {
  const invoiceAmount = Number(invoice.invoice_amount || 0);
  const paidAmount = Number(invoice.paid_amount || 0);
  const balance = Number(payments.invoiceBalance || 0);

  if (balance <= 0.005 && invoiceAmount > 0) {
    return 'Paid';
  }

  if (paidAmount > 0 || balance < invoiceAmount) {
    return 'Partially Paid';
  }

  return 'Unpaid';
}


module.exports = {
    chromium,
    BASE_URL,
    getProjectInvoices,
    getInvoiceCalculations,
    getInvoicePayments,
    determinePaymentStatus
};