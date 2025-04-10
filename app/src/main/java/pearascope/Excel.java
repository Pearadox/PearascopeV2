package pearascope;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCustomFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCustomFilters;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilterColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STFilterOperator;

public class Excel {
    public static int addOutputHeader(XSSFSheet sheet, int rowIndex, String... values) {
        Row row = sheet.createRow(rowIndex);
        for (int i = 0; i < values.length; i++) {
            row.createCell(i).setCellValue(values[i]);
        }
        return row.getRowNum();
    }

    public static String getCellString(Row row, ColumnEnum column) {
        int columnIndex = ((Enum<?>) column).ordinal();
        Cell cell = row.getCell(columnIndex);
        return (cell != null ? cell.getStringCellValue() : "");
    }

    public static double getCellNumber(Row row, ColumnEnum column) {
        int columnIndex = ((Enum<?>) column).ordinal();
        Cell cell = row.getCell(columnIndex);
        return (cell != null ? cell.getNumericCellValue() : 0.0);
    }

    public static void setCellValue(Row row, ColumnEnum column, String value) {
        int columnIndex = ((Enum<?>) column).ordinal();
        Cell cell = row.getCell(columnIndex);
        if (cell != null ) cell.setCellValue(value);
    }

    public static void setCellValue(Row row, ColumnEnum column, double value) {
        int columnIndex = ((Enum<?>) column).ordinal();
        Cell cell = row.getCell(columnIndex);
        if (cell != null ) cell.setCellValue(value);
    }

    public static void setCellFormula(Row row, ColumnEnum column, String formula) {
        int columnIndex = ((Enum<?>) column).ordinal();
        Cell cell = row.getCell(columnIndex);
        if (cell != null ) cell.setCellFormula(formula);
    }

    public static ConditionalFormattingRule highlightTextContaining(SheetConditionalFormatting sheetCF, String textToMatch, byte[] backgroundColor, byte[] foregroundColor) {
        ConditionalFormattingRule rule = sheetCF.createConditionalFormattingRule("ISNUMBER(SEARCH(\"" + textToMatch + "\", G1))");
        PatternFormatting pattern = rule.createPatternFormatting();
        pattern.setFillBackgroundColor(new XSSFColor(backgroundColor));
        pattern.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
        rule.createFontFormatting().setFontColor(new XSSFColor(foregroundColor));
        return rule;
    }

    public static void convertToTable(XSSFWorkbook workbook, XSSFSheet sheet) {
        int lastRow = sheet.getLastRowNum();
        int lastCol = sheet.getRow(0).getLastCellNum() - 1;

        CellReference topLeft = new CellReference(0, 0);
        CellReference bottomRight = new CellReference(lastRow, lastCol);
        AreaReference area = new AreaReference(topLeft, bottomRight, workbook.getSpreadsheetVersion());

        XSSFTable table = sheet.createTable(area);
        table.setDisplayName("Analysis");

        CTTable ctTable = table.getCTTable();
        CTTableStyleInfo style = ctTable.addNewTableStyleInfo();
        style.setName("TableStyleMedium1");
        style.setShowRowStripes(true);
        style.setShowColumnStripes(false);        
    }
    
    public static void filterValues(XSSFSheet sheet, ColumnEnum column, String... filterValues) {
        int columnIndex = ((Enum<?>) column).ordinal();

        XSSFTable table = sheet.getTables().get(0);
        CTTable ctTable = table.getCTTable();
        CTAutoFilter filter = ctTable.addNewAutoFilter();
        CTFilterColumn filterColumn = filter.addNewFilterColumn();
        filterColumn.setColId(columnIndex);
        CTCustomFilters customFilters = filterColumn.addNewCustomFilters();

        for (String filterValue : filterValues) {
            CTCustomFilter filterCriteria = customFilters.addNewCustomFilter();
            filterCriteria.setOperator(STFilterOperator.Enum.forInt(1));
            filterCriteria.setVal(filterValue);
        }

        for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
            XSSFRow row = sheet.getRow(rowNum);
            if (row != null) {
                XSSFCell cell = row.getCell(columnIndex);
                boolean match = false;
                if (cell != null) {
                    String cellValue = cell.getStringCellValue();
                    for (String filterValue : filterValues) {
                        if (cellValue.equals(filterValue)) {
                            match = true;
                            break;
                        }
                    }
                }
                if (!match) {
                    row.getCTRow().setHidden(true);
                }
            }
        }
    }
}
