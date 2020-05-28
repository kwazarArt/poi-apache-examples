package ru.kwazarart.java.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;

public class SizeApp {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("Лист 1");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("Новая ячейка");

        //sheet.setColumnWidth(0, 5000);
        //sheet.setColumnWidth(3, 3000);
        //sheet.autoSizeColumn(0);

        //row.setHeightInPoints(24);

        sheet.addMergedRegion(new CellRangeAddress(0, 5, 0,2));

        FileOutputStream fos = new FileOutputStream("размер ячейки.xls");
        wb.write(fos);
        wb.close();
        fos.close();
    }
}
