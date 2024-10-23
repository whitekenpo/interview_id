package interview_ricky;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.*;
import java.util.*;

public class InsuranceTemplateConverter {
    private static final int COLUMN_A = 0;
    private static final int COLUMN_B = 1;

    public static class InsuranceEntry {
        String benefit; // 紫色背景的主要福利類別
        String coverage; // A欄非紫色背景的項目
        String category; // B欄的類別描述
        String planName; // Plan名稱
        String coverageValue; // 對應的值

        public InsuranceEntry(String benefit, String coverage,
                String category, String planName,
                String coverageValue) {
            this.benefit = benefit;
            this.coverage = coverage;
            this.category = category;
            this.planName = planName;
            this.coverageValue = coverageValue;
        }

        // Override toString method
        @Override
        public String toString() {
            return "" + benefit +
                    ", " + coverage +
                    ", " + category +
                    ", " + planName +
                    ", =" + coverageValue + "";
        }
    }

    public void convertTemplate(String inputPath) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputPath);
                Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sourceSheet = workbook.getSheetAt(0);
            List<InsuranceEntry> entries = extractEntries(sourceSheet);

            // Print each entry in a row
            for (InsuranceEntry entry : entries) {
                System.out.println(entry);
            }
        }
    }

    private List<InsuranceEntry> extractEntries(Sheet sheet) {
        List<InsuranceEntry> entries = new ArrayList<>();
        String currentBenefit = null;
        String coverage = null;
        int lastRow = sheet.getLastRowNum();

        for (int i = 1; i <= lastRow; i++) {
            Row row = sheet.getRow(i);
            if (row == null)
                continue;

            Cell cellA = row.getCell(COLUMN_A);

            // 檢查是否為紫色背景的Benefit
            if (isPurpleBackground(cellA)) {
                currentBenefit = getCellValue(cellA);
                continue;
            }

            // 處理非紫色背景的A欄項目作為Coverage
            if (cellA != null) { // A 欄不為空, update the coverage as the previous one
                coverage = getCellValue(cellA);
            }

            Cell cellB = row.getCell(COLUMN_B);
            // if( cellB == null) continue;

            String category = getCellValue(cellB);

            // will use while to do so
            for (int planIndex = 1, cellnum = 5; row.getCell(cellnum) != null; cellnum++, planIndex++) {
                // 處理Plan
                processPlan(entries, currentBenefit, coverage, category, ("Plan " + planIndex), row.getCell(cellnum));
            }

        }
        return entries;
    }

    private boolean isPurpleBackground(Cell cell) {
        if (cell == null)
            return false;

        CellStyle style = cell.getCellStyle();
        if (style == null)
            return false;

        // 檢查填充顏色
        return style.getFillForegroundColorColor() != null;
    }

    private void processPlan(List<InsuranceEntry> entries,
            String benefit, String coverage, String category,
            String planName, Cell valueCell) {
        if (valueCell != null && !isEmpty(valueCell)) {
            String value = getCellValue(valueCell);

            // add into entries
            entries.add(new InsuranceEntry(
                    benefit,
                    coverage,
                    category,
                    planName,
                    value));
        }
    }

    private String getCellValue(Cell cell) {
        if (cell == null)
            return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                double value = cell.getNumericCellValue();
                // 處理整數和小數的顯示格式
                if (value == Math.floor(value)) {
                    return String.format("%.0f", value);
                }
                // 保留到小數點後2位
                return String.format("%.2f", value);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return String.valueOf(cell.getNumericCellValue());
                } catch (Exception e) {
                    return cell.getStringCellValue();
                }
            default:
                return "";
        }
    }

    private boolean isEmpty(Cell cell) {
        return cell == null ||
                cell.getCellType() == CellType.BLANK ||
                getCellValue(cell).trim().isEmpty();
    }

    public static void main(String[] args) {
        try {
            InsuranceTemplateConverter converter = new InsuranceTemplateConverter();
            converter.convertTemplate(
                    "Q1.xlsx");
            System.out.println("Template conversion completed successfully!");
        } catch (IOException e) {
            System.err.println("Error converting template: " + e.getMessage());
            e.printStackTrace();
        }
    }
}