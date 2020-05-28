package ru.kwazarart.java.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.FileOutputStream;
import java.io.IOException;

public class JavaExcelApp {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook();

        Sheet sheet0 = wb.createSheet("Издатели");
        Row row = sheet0.createRow(3);
        Cell cell = row.createCell(4);
        cell.setCellValue("O'Reilly");



        Sheet sheet1 = wb.createSheet("Произведение");
        Row row1 = sheet1.createRow(0);
        Cell cell1 = row1.createCell(0);
        cell1.setCellValue("Война и мир");


        Row row2 = sheet1.createRow(1);
        Cell cell2 = row2.createCell(3);
        cell2.setCellValue("Евгений Онегин");

        Sheet sheet2 = wb.createSheet("Авторы");
        Sheet sheet3 = wb.createSheet(WorkbookUtil.createSafeSheetName("skldfhgsi*(^*962!!**")); //spec simbols

        FileOutputStream fos = new FileOutputStream("my.xls");

        wb.write(fos);
        fos.close();
    }
}
