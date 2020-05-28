package ru.kwazarart.java.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

public class StyleApp {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet0 = wb.createSheet("Лист 1");
        Row row = sheet0.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("Привет");

        CellStyle style = wb.createCellStyle();
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());  // for background color
        //style.setFillBackgroundColor(IndexedColors.GREEN.getIndex());
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setBorderBottom(BorderStyle.THIN);


        Font font = wb.createFont();
        font.setFontName("Courier New");
        font.setFontHeightInPoints((short) 15);
        font.setBold(true); // жирный
        font.setStrikeout(true); // зачеркнутый
        font.setUnderline(Font.U_SINGLE); // подчеркнутый
        font.setColor(IndexedColors.RED.getIndex());

        style.setFont(font);

        cell.setCellStyle(style);

        FileOutputStream fos = new FileOutputStream("стили.xls");
        wb.write(fos);
        wb.close();
        fos.close();
    }
}
