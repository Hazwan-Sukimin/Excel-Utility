package com.abs.dps.utils;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public class ExcelUtility {
    public HSSFWorkbook workbook;
    public HSSFSheet currentSheet;
    public HSSFRow currentRow;
    public int rowIndex;

    /**
     * Constructor
     */
    public ExcelUtility() {
        workbook = new HSSFWorkbook();
    }

    /**
     * Creates a new sheet in the workbook
     *
     * @param sheetName
     */
    public void createSheet(String sheetName) {
        currentSheet = workbook.createSheet(sheetName);
        rowIndex = 0;
    }

    /**
     * Creates a new row in the current sheet
     */
    public void createNewRow() {
        currentRow = currentSheet.createRow(rowIndex++);
    }

    /**
     * Creates a new cell in the current row
     *
     * @param columnIndex
     * @return
     */
    public HSSFCell createCell(int columnIndex) {
        if (currentRow == null) {
            createNewRow();
        }
        return currentRow.createCell(columnIndex);
    }

    /**
     * Sets the cell value to the specified value
     *
     * @param cell
     * @param value
     */
    public void setCellValue(HSSFCell cell, String value) {
        cell.setCellValue(value);
    }

    /**
     * Merges the specified cells
     *
     * @param firstRow
     * @param lastRow
     * @param firstCol
     * @param lastCol
     */
    public void mergeCells(int firstRow, int lastRow, int firstCol, int lastCol) {
        currentSheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));

        RegionUtil.setBorderTop(1, new CellRangeAddress(firstRow, lastRow, firstCol, lastCol), currentSheet, workbook);
        RegionUtil.setBorderBottom(1, new CellRangeAddress(firstRow, lastRow, firstCol, lastCol), currentSheet, workbook);
        RegionUtil.setBorderLeft(1, new CellRangeAddress(firstRow, lastRow, firstCol, lastCol), currentSheet, workbook);
        RegionUtil.setBorderRight(1, new CellRangeAddress(firstRow, lastRow, firstCol, lastCol), currentSheet, workbook);

    }

    public void addCell(String value, int row, int col, CellRangeAddress cellRangeAddress, HSSFCellStyle cellStyle){

        HSSFCell cell = createCell(col++);
        setCellValue(cell, value);
        cell.setCellStyle(cellStyle);

        currentSheet.addMergedRegion(cellRangeAddress);
        RegionUtil.setBorderTop(1, cellRangeAddress, currentSheet, workbook);
        RegionUtil.setBorderBottom(1, cellRangeAddress, currentSheet, workbook);
        RegionUtil.setBorderLeft(1, cellRangeAddress, currentSheet, workbook);
        RegionUtil.setBorderRight(1, cellRangeAddress, currentSheet, workbook);

        currentSheet.autoSizeColumn(col);
    }


    /**
     * Skips the specified number of rows
     *
     * @param numRowsToSkip
     */
    public void skipRow(int numRowsToSkip) {
        for (int i = 0; i < numRowsToSkip; i++) {
            createNewRow();
            int columnIndex = 0;
            for (int j = 0; j < 4; j++) {  // Replace 4 with the actual maximum column index
                HSSFCell cell = createCell(columnIndex++);
                // Set the cell value to an empty string to create an empty cell
                setCellValue(cell, "");
            }
        }
    }


    /**
     * Saves the workbook to the specified file
     *
     * @param fileName
     * @param response
     */
    public void saveWorkbook(String fileName, HttpServletResponse response) {
        fileName = fileName + ".xls";
        try {
            ByteArrayOutputStream outByteStream = new ByteArrayOutputStream();
            workbook.write(outByteStream);
            byte[] outArray = outByteStream.toByteArray();
            response.setContentType("application/vnd.ms-excel");
            response.setContentLength(outArray.length);
            response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"");
            OutputStream outStream = response.getOutputStream();
            outStream.write(outArray);
            outStream.flush();
            outStream.close();
            outByteStream.close();

        } catch (Exception e) {
            LogFunction.loginfo("export 1 " + e);
        }
    }


    /**
     * Adds a row with grey background and bold font
     *
     * @param rowData
     */
    public void addRowWithData(List<String> rowData) {
        createNewRow();
        int columnIndex = 0;
        for (String value : rowData) {
            HSSFCell cell = createCell(columnIndex++);
            setCellValue(cell, value);
        }

        // Get the current sheet and auto-size the columns
        Sheet sheet = currentSheet;
        for (int i = 0; i < rowData.size(); i++) {
            sheet.autoSizeColumn(i);
        }
    }

    /**
     * Creates a cell style with grey background and bold font
     *
     * @return
     */
    public HSSFCellStyle createGreyBackgroundCellStyle() {
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        // Create a font with bold style

        HSSFFont font = workbook.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

        // Set the font in the style
        style.setFont(font);

        return style;
    }

    /**
     * Creates a cell style with bold font
     *
     * @return
     */
    public HSSFCellStyle createFontBoldCellStyle() {
        // Initialize a new instance of HSSFCellStyle
        HSSFCellStyle style = workbook.createCellStyle();

        // Create a font with bold style
        HSSFFont font = workbook.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

        // Set the font in the style
        style.setFont(font);

        return style;
    }

    /**
     * Creates a cell style with border
     * @return
     */
    public HSSFCellStyle createBorderCellStyle() {
        // Initialize a new instance of HSSFCellStyle
        HSSFCellStyle style = workbook.createCellStyle();
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setAlignment(CellStyle.ALIGN_CENTER);

        return style;
    }

    /**
     * Adds a row with custom style
     *
     * @param rowData
     * @param cellStyle
     */
    public void addRowWithCustomStyle(List<String> rowData, HSSFCellStyle cellStyle) {
        createNewRow();
        int columnIndex = 0;
        for (String value : rowData) {
            HSSFCell cell = createCell(columnIndex++);
            setCellValue(cell, value);
            cell.setCellStyle(cellStyle);
            currentSheet.autoSizeColumn(columnIndex - 1); // Auto-size the column
        }
    }

    /**
     * Adds a row with custom style and merges the specified number of columns
     *
     * @param rowData
     * @param cellStyle
     */
    public void addRowWithCustomStyleMerge(List<String> rowData, HSSFCellStyle cellStyle) {
        addRowWithCustomStyleMerge(rowData, cellStyle, 0); // Assume 0 columns to merge by default
    }

    /**
     * Adds a row with custom style and merges the specified number of columns
     *
     * @param rowData
     * @param cellStyle
     * @param numColumnsToMerge
     */
    public void addRowWithCustomStyleMerge(List<String> rowData, HSSFCellStyle cellStyle, int numColumnsToMerge) {
        createNewRow();
        int columnIndex = 0;
        for (String value : rowData) {
            HSSFCell cell = createCell(columnIndex);
            setCellValue(cell, value);
            cell.setCellStyle(cellStyle);
            currentSheet.addMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex - 1, 0, (numColumnsToMerge - 1)));
            columnIndex++;
        }

    }

    public void addCellWithRowspan(List<String> rowData, HSSFCellStyle cellStyle, int rowspan) {
        Row row = currentSheet.createRow(rowIndex);

        for (int i = 0; i < rowData.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(rowData.get(i));
            cell.setCellStyle(cellStyle);

            if (rowspan > 1) {
                currentSheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + rowspan - 1, i, i));
                RegionUtil.setBorderTop(1, new CellRangeAddress(rowIndex, rowIndex + rowspan - 1, i, i), currentSheet, workbook);
                RegionUtil.setBorderBottom(1, new CellRangeAddress(rowIndex, rowIndex + rowspan - 1, i, i), currentSheet, workbook);
                RegionUtil.setBorderLeft(1, new CellRangeAddress(rowIndex, rowIndex + rowspan - 1, i, i), currentSheet, workbook);
                RegionUtil.setBorderRight(1, new CellRangeAddress(rowIndex, rowIndex + rowspan - 1, i, i), currentSheet, workbook);
            }
            currentSheet.autoSizeColumn(cell.getColumnIndex(), true);
        }
        rowIndex += rowspan;
    }

    public void addCellsWithRowspanJoin(List<String> values, HSSFCellStyle cellStyle, int rowspan) {
        int startColumn = currentRow.getLastCellNum() == -1 ? 0 : currentRow.getLastCellNum();

        for (int i = 0; i < values.size(); i++) {
            HSSFCell cell = createCell(startColumn + i);
            setCellValue(cell, values.get(i));
            cell.setCellStyle(cellStyle);

            if (rowspan > 1) {
                currentSheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + rowspan - 1, startColumn + i, startColumn + i));
            }
        }
    }

    public void addCellToLastCellNum(String value) {

        int lastCellNum = currentRow.getLastCellNum();
        Cell cell = currentRow.createCell(lastCellNum);
        cell.setCellValue(value);
    }

    public void mergeCellsRowInRange(int startColumn, int endColumn, int rowNum) {
        Sheet sheet = workbook.getSheetAt(0); // Assuming you're working with the first sheet

        for (int column = startColumn; column <= endColumn; column++) {
            CellRangeAddress cellRange = new CellRangeAddress(rowNum, rowNum, column, column);
            sheet.addMergedRegion(cellRange);
        }
    }

}
