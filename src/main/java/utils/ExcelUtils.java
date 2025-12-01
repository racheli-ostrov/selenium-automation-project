package utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.File;
import java.util.List;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class ExcelUtils {

    // פונקציה עם השוואה בין מצופה לבפועל (4 פרמטרים)
    public static void writeCartToExcel(String path, List<pages.CartPage.CartItem> expectedItems, 
                                         List<pages.CartPage.CartItem> actualItems, String total) throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        
        // Create styles
        XSSFCellStyle headerStyle = workbook.createCellStyle();
        XSSFFont headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 12);
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        XSSFCellStyle successStyle = workbook.createCellStyle();
        successStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        successStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        XSSFCellStyle failStyle = workbook.createCellStyle();
        failStyle.setFillForegroundColor(IndexedColors.CORAL.getIndex());
        failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        // Sheet 1: Expected vs Actual Comparison
        Sheet comparisonSheet = workbook.createSheet("השוואה (Comparison)");
        Row header = comparisonSheet.createRow(0);
        String[] headers = {"פריט (Item)", "כמות מצופה (Expected Qty)", "כמות בפועל (Actual Qty)", 
                            "מחיר מצופה (Expected Price)", "מחיר בפועל (Actual Price)", "סטטוס (Status)"};
        for (int i = 0; i < headers.length; i++) {
            Cell cell = header.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
            comparisonSheet.setColumnWidth(i, 5000);
        }
        
        int rowIdx = 1;
        boolean allMatch = true;
        
        // Compare expected vs actual
        int maxItems = Math.max(expectedItems.size(), actualItems.size());
        for (int i = 0; i < maxItems; i++) {
            Row r = comparisonSheet.createRow(rowIdx++);
            
            pages.CartPage.CartItem expected = i < expectedItems.size() ? expectedItems.get(i) : null;
            pages.CartPage.CartItem actual = i < actualItems.size() ? actualItems.get(i) : null;
            
            String itemName = expected != null ? expected.name : (actual != null ? actual.name : "N/A");
            r.createCell(0).setCellValue(itemName);
            
            r.createCell(1).setCellValue(expected != null ? expected.qty : 0);
            r.createCell(2).setCellValue(actual != null ? actual.qty : 0);
            
            r.createCell(3).setCellValue(expected != null ? expected.price : "N/A");
            r.createCell(4).setCellValue(actual != null ? actual.price : "N/A");
            
            boolean itemMatches = expected != null && actual != null && expected.qty == actual.qty;
            
            Cell statusCell = r.createCell(5);
            if (expected == null) {
                statusCell.setCellValue("✗ פריט לא מצופה (Unexpected Item)");
                statusCell.setCellStyle(failStyle);
                allMatch = false;
            } else if (actual == null) {
                statusCell.setCellValue("✗ פריט חסר (Missing Item)");
                statusCell.setCellStyle(failStyle);
                allMatch = false;
            } else if (itemMatches) {
                statusCell.setCellValue("✓ תואם (Match)");
                statusCell.setCellStyle(successStyle);
            } else {
                statusCell.setCellValue("✗ אי התאמה (Mismatch)");
                statusCell.setCellStyle(failStyle);
                allMatch = false;
            }
        }
        
        // Add total row
        Row totalRow = comparisonSheet.createRow(rowIdx + 1);
        Cell totalLabelCell = totalRow.createCell(4);
        totalLabelCell.setCellValue("סה\"כ עגלה (Cart Total):");
        totalLabelCell.setCellStyle(headerStyle);
        Cell totalValueCell = totalRow.createCell(5);
        totalValueCell.setCellValue(total);
        totalValueCell.setCellStyle(headerStyle);
        
        // Sheet 2: Summary
        Sheet summarySheet = workbook.createSheet("סיכום (Summary)");
        int summaryRow = 0;
        
        Row timeRow = summarySheet.createRow(summaryRow++);
        timeRow.createCell(0).setCellValue("זמן הרצה (Run Time):");
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
        timeRow.createCell(1).setCellValue(timestamp);
        
        summaryRow++; // Empty row
        
        Row testResultRow = summarySheet.createRow(summaryRow++);
        Cell resultLabel = testResultRow.createCell(0);
        resultLabel.setCellValue("תוצאת בדיקה (Test Result):");
        resultLabel.setCellStyle(headerStyle);
        Cell resultValue = testResultRow.createCell(1);
        if (allMatch && expectedItems.size() == actualItems.size()) {
            resultValue.setCellValue("✓ הבדיקה עברה בהצלחה (PASS)");
            resultValue.setCellStyle(successStyle);
        } else {
            resultValue.setCellValue("✗ הבדיקה נכשלה (FAIL)");
            resultValue.setCellStyle(failStyle);
        }
        
        summaryRow++; // Empty row
        
        Row expectedCountRow = summarySheet.createRow(summaryRow++);
        expectedCountRow.createCell(0).setCellValue("פריטים מצופים (Expected Items):");
        expectedCountRow.createCell(1).setCellValue(expectedItems.size());
        
        Row actualCountRow = summarySheet.createRow(summaryRow++);
        actualCountRow.createCell(0).setCellValue("פריטים בפועל (Actual Items):");
        actualCountRow.createCell(1).setCellValue(actualItems.size());
        
        Row totalQtyExpected = summarySheet.createRow(summaryRow++);
        int expectedQty = expectedItems.stream().mapToInt(i -> i.qty).sum();
        totalQtyExpected.createCell(0).setCellValue("כמות כוללת מצופה (Expected Total Qty):");
        totalQtyExpected.createCell(1).setCellValue(expectedQty);
        
        Row totalQtyActual = summarySheet.createRow(summaryRow++);
        int actualQty = actualItems.stream().mapToInt(i -> i.qty).sum();
        totalQtyActual.createCell(0).setCellValue("כמות כוללת בפועל (Actual Total Qty):");
        totalQtyActual.createCell(1).setCellValue(actualQty);
        
        summarySheet.autoSizeColumn(0);
        summarySheet.autoSizeColumn(1);
        
        // Write to file
        File file = new File(path);
        file.getParentFile().mkdirs();
        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
        }
        workbook.close();
        
        System.out.println("📊 דוח Excel נשמר ב: " + path);
    }

    // Append a single test result (testName, status) into an Excel file.
    // If the file does not exist, create it and write a header row first.
    public static synchronized void appendTestResult(String path, String testName, String status) {
        File file = new File(path);
        Workbook workbook = null;
        Sheet sheet = null;
        try {
            if (file.exists()) {
                try (FileInputStream fis = new FileInputStream(file)) {
                    workbook = new XSSFWorkbook(fis);
                }
            } else {
                workbook = new XSSFWorkbook();
            }
            sheet = workbook.getSheet("Results");
            if (sheet == null) {
                sheet = workbook.createSheet("Results");
                Row header = sheet.createRow(0);
                header.createCell(0).setCellValue("Test Name");
                header.createCell(1).setCellValue("Status");
            }
            int lastRow = sheet.getLastRowNum();
            // If sheet only has header and no data yet, lastRow might be 0
            // Determine actual next row index (handle empty sheet case)
            int nextRowIdx = (sheet.getPhysicalNumberOfRows() == 0) ? 0 : lastRow + 1;
            Row r = sheet.createRow(nextRowIdx);
            r.createCell(0).setCellValue(testName);
            r.createCell(1).setCellValue(status);
            // Autosize columns for readability (lightweight for small sheet)
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);
            file.getParentFile().mkdirs();
            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                try { workbook.close(); } catch (Exception ignored) {}
            }
        }
    }

    // Class to hold detailed cart item info for Excel reporting
    public static class CartItemReport {
        public String searchQuery;
        public String productName;
        public boolean addedSuccessfully;
        public String price;
        public String quantity;
        public String rowTotal;

        public CartItemReport(String searchQuery, String productName, boolean addedSuccessfully, 
                               String price, String quantity, String rowTotal) {
            this.searchQuery = searchQuery;
            this.productName = productName;
            this.addedSuccessfully = addedSuccessfully;
            this.price = price;
            this.quantity = quantity;
            this.rowTotal = rowTotal;
        }
    }

    /**
     * Writes comprehensive cart test report to Excel.
     * Creates a new workbook each run with timestamp.
     * @param path Output Excel file path
     * @param items List of cart item reports
     * @param cartTotal Total cart value
     * @param testStatus Overall test status (PASS/FAIL)
     */
    public static void writeCartTestReport(String path, List<CartItemReport> items, 
                                            String cartTotal, String testStatus) throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        
        // Create header style
        XSSFCellStyle headerStyle = workbook.createCellStyle();
        XSSFFont headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 12);
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        // Create success/fail styles
        XSSFCellStyle successStyle = workbook.createCellStyle();
        successStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        successStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        XSSFCellStyle failStyle = workbook.createCellStyle();
        failStyle.setFillForegroundColor(IndexedColors.CORAL.getIndex());
        failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        // Main cart items sheet
        Sheet itemsSheet = workbook.createSheet("Cart Items");
        Row header = itemsSheet.createRow(0);
        String[] headers = {"חיפוש (Search Query)", "שם מוצר (Product Name)", 
                            "הוסף בהצלחה (Added Successfully)", "מחיר (Price)", 
                            "כמות (Quantity)", "סה\"כ שורה (Row Total)"};
        for (int i = 0; i < headers.length; i++) {
            Cell cell = header.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
            itemsSheet.setColumnWidth(i, 5000);
        }
        
        int rowIdx = 1;
        for (CartItemReport item : items) {
            Row r = itemsSheet.createRow(rowIdx++);
            r.createCell(0).setCellValue(item.searchQuery);
            r.createCell(1).setCellValue(item.productName);
            
            Cell statusCell = r.createCell(2);
            statusCell.setCellValue(item.addedSuccessfully ? "✓ כן" : "✗ לא");
            statusCell.setCellStyle(item.addedSuccessfully ? successStyle : failStyle);
            
            r.createCell(3).setCellValue(item.price);
            r.createCell(4).setCellValue(item.quantity);
            r.createCell(5).setCellValue(item.rowTotal);
        }
        
        // Add total row
        Row totalRow = itemsSheet.createRow(rowIdx + 1);
        Cell totalLabelCell = totalRow.createCell(4);
        totalLabelCell.setCellValue("סה\"כ עגלה (Cart Total):");
        totalLabelCell.setCellStyle(headerStyle);
        Cell totalValueCell = totalRow.createCell(5);
        totalValueCell.setCellValue(cartTotal);
        totalValueCell.setCellStyle(headerStyle);
        
        // Summary sheet
        Sheet summarySheet = workbook.createSheet("Summary");
        int summaryRow = 0;
        
        Row timeRow = summarySheet.createRow(summaryRow++);
        timeRow.createCell(0).setCellValue("זמן הרצה (Run Time):");
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
        timeRow.createCell(1).setCellValue(timestamp);
        
        Row statusRow = summarySheet.createRow(summaryRow++);
        statusRow.createCell(0).setCellValue("סטטוס טסט (Test Status):");
        Cell statusCell = statusRow.createCell(1);
        statusCell.setCellValue(testStatus);
        statusCell.setCellStyle(testStatus.contains("PASS") ? successStyle : failStyle);
        
        Row itemCountRow = summarySheet.createRow(summaryRow++);
        itemCountRow.createCell(0).setCellValue("מספר פריטים (Items Count):");
        itemCountRow.createCell(1).setCellValue(items.size());
        
        Row successCountRow = summarySheet.createRow(summaryRow++);
        successCountRow.createCell(0).setCellValue("הוספות מוצלחות (Successful Adds):");
        long successCount = items.stream().filter(i -> i.addedSuccessfully).count();
        successCountRow.createCell(1).setCellValue(successCount);
        
        Row totalRow2 = summarySheet.createRow(summaryRow++);
        totalRow2.createCell(0).setCellValue("סה\"כ עגלה (Cart Total):");
        totalRow2.createCell(1).setCellValue(cartTotal);
        
        summarySheet.autoSizeColumn(0);
        summarySheet.autoSizeColumn(1);
        
        // Write to file
        File file = new File(path);
        file.getParentFile().mkdirs();
        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
        }
        workbook.close();
        
        System.out.println("📊 דוח Excel נשמר ב: " + path);
    }
}
