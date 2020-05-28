package ru.kwazarart.java.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;

public class App {

    public static SimpleDateFormat sdf = new SimpleDateFormat("yyyy.MM.dd");

    public static void main(String[] args) throws IOException {
        FileInputStream fis = new FileInputStream("книга1.xls");

        Workbook wb = new HSSFWorkbook(fis);

        for (Row row : wb.getSheetAt(0)) {
            for (Cell cell : row) {
                CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
                System.out.print(cellRef.formatAsString());
                System.out.print(" - ");
                System.out.println(getCellText(cell));
            }
        }
        fis.close();
    }

    public static String getCellText(Cell cell) {
        String result = "";
        switch (cell.getCellType()) {
            case STRING:
                result = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    result = sdf.format(cell.getDateCellValue());
                } else {
                    result = String.valueOf(cell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                result = String.valueOf( cell.getBooleanCellValue());
                break;
            case FORMULA:
                result = cell.getCellFormula().toString();
                break;
            default:
                break;
        }
        return result;
    }
}
