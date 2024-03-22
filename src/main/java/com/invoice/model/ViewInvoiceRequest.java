package com.invoice.model;

import java.util.Date;

public record ViewInvoiceRequest(String id, String invoice_link, String status) {
}
