package ru.kwazarart.java.excel;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class DrawShapes {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("Картинки");

        HSSFPatriarch patriarch = (HSSFPatriarch) sheet.createDrawingPatriarch(); // создаем художника

        HSSFClientAnchor anchor = new HSSFClientAnchor(); // создаем холст
        anchor.setCol1(2);
        anchor.setRow1(2);
        anchor.setCol2(10);
        anchor.setRow2(10);

        HSSFSimpleShape shape = patriarch.createSimpleShape(anchor); // создаем фигуру (создается художником)
        shape.setShapeType(HSSFSimpleShape.OBJECT_TYPE_OVAL);
        shape.setLineStyleColor(255, 0, 0);
        shape.setLineWidth(HSSFSimpleShape.LINEWIDTH_ONE_PT*3);
        shape.setLineStyle(HSSFSimpleShape.LINESTYLE_DASHDOTGEL);
        shape.setFillColor(0,0,255);

        FileOutputStream fos = new FileOutputStream("abc.xls");
        wb.write(fos);
        wb.close();
        fos.close();
    }
}
