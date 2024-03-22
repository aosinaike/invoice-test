package com.invoice.controller;

import com.invoice.model.InvoiceRequest;
import com.invoice.model.ViewInvoiceRequest;
import org.apache.commons.io.IOUtils;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.*;
import java.net.URISyntaxException;
import java.nio.file.Path;
import java.nio.file.Paths;

import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.Table;
import com.spire.doc.fields.Field;

@RestController
@RequestMapping("/api/v1/invoice")
public class InvoiceController {

    @PostMapping
    public ResponseEntity<Object> GenerateInvoiceWebHook(@RequestBody InvoiceRequest invoiceRequest)  {

        Document doc = new Document();
        //load the template file
        doc.loadFromFile("src/main/resources/Invoice-Template.docx");

        //replace text in the document
        doc.replace("#InvoiceNum", invoiceRequest.invoiceNumber(), true, true);
        doc.replace("#CompanyName", invoiceRequest.customerName(), true, true);
        doc.replace("#CompanyAddress", invoiceRequest.customerAddress(), true, true);
        doc.replace("#CityStateZip", invoiceRequest.customerCity(), true, true);
        doc.replace("#Country", invoiceRequest.customerCountry(), true, true);
        doc.replace("#Tel1", invoiceRequest.customerPhone(), true, true);
        doc.replace("#ContactPerson", invoiceRequest.customerShipping(), true, true);
        doc.replace("#ShippingAddress", invoiceRequest.shippingDetailAddress(), true, true);
        doc.replace("#Tel2", invoiceRequest.shippingDetailPhone(), true, true);

        //define purchase data
        String[][] purchaseData = {
                new String[]{"Product A", "5", "22.8"},
                new String[]{"Product B", "4", "35.3"},
                new String[]{"Product C", "2", "52.9"},
                new String[]{"Product D", "3", "25"},
        };

        //write the purchase data to the document
        writeDataToDocument(doc, purchaseData);

        //update fields
        doc.isUpdateFields(true);

        //save file in pdf format
        doc.saveToFile(invoiceRequest.id()+".pdf", FileFormat.PDF);
        return new ResponseEntity<>("http://localhost:8080/api/v1/invoice/view/"+invoiceRequest.id(), HttpStatus.OK);
    }

    @GetMapping(value = "/view/{invoice}", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
    public ResponseEntity<Object> viewInvoice(@PathVariable String invoice)  {
        String projectPath = System.getProperty("user.dir");
        File file = new File(projectPath+"\\"+invoice+".pdf");
        InputStream inputStream = null;
        try {
            inputStream = new FileInputStream(file);
            return new ResponseEntity<>(IOUtils.toByteArray(inputStream),HttpStatus.OK);
        } catch (FileNotFoundException e) {
            return new ResponseEntity<>(null,HttpStatus.NOT_FOUND);
        } catch (IOException e) {
            return new ResponseEntity<>(null,HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    private static void writeDataToDocument(Document doc, String[][] purchaseData) {
        //get the third table
        Table table = doc.getSections().get(0).getTables().get(2);
        //determine if it needs to add rows
        if (purchaseData.length > 1) {
            //add rows
            addRows(table, purchaseData.length - 1);
        }
        //fill the table cells with value
        fillTableWithData(table, purchaseData);
    }

    private static void addRows(Table table, int rowNum) {
        for (int i = 0; i < rowNum; i++) {
            //insert specific number of rows by cloning the second row
            table.getRows().insert(2 + i, table.getRows().get(1).deepClone());
            //update formulas for Total
            for (Object object : table.getRows().get(2 + i).getCells().get(3).getParagraphs().get(0).getChildObjects()
            ) {
                if (object instanceof Field) {
                    Field field = (Field) object;
                    field.setCode(String.format("=B%d*C%d\\# \"0.00\"", 3 + i,3 + i));
                }
                break;
            }
        }
        //update formula for Total Tax
        for (Object object : table.getRows().get(4 + rowNum).getCells().get(3).getParagraphs().get(0).getChildObjects()
        ) {
            if (object instanceof Field) {
                Field field = (Field) object;
                field.setCode(String.format("=D%d*0.05\\# \"0.00\"", 3 + rowNum));
            }
            break;
        }
        //update formula for Balance Due
        for (Object object : table.getRows().get(5 + rowNum).getCells().get(3).getParagraphs().get(0).getChildObjects()
        ) {
            if (object instanceof Field) {
                Field field = (Field) object;
                field.setCode(String.format("=D%d+D%d\\# \"$#,##0.00\"", 3 + rowNum, 5 + rowNum));
            }
            break;
        }
    }


    private static void fillTableWithData(Table table, String[][] data) {
        for (int r = 0; r < data.length; r++) {
            for (int c = 0; c < data[r].length; c++) {
                //fill data in cells
                table.getRows().get(r + 1).getCells().get(c).getParagraphs().get(0).setText(data[r][c]);
            }
        }
    }
}
