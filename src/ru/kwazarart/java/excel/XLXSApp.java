package ru.kwazarart.java.excel;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class XLXSApp {
    public static void main(String[] args) throws IOException {
        Workbook wb = new XSSFWorkbook();

        Sheet sheet0 = wb.createSheet("Издатели");
        Row row = sheet0.createRow(3);
        Cell cell = row.createCell(4);
        cell.setCellValue("O'Reilly");


        FileOutputStream fos = new FileOutputStream("workbook.xlsx");
        wb.write(fos);
        fos.close();
    }
}
