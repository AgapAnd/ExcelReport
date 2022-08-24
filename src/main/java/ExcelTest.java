import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelTest {
    public static void main(String[] args) throws IOException {
        File file = new File("C:\\Users\\PKT8\\Desktop\\javaTest.xls");
        FileInputStream fis = new FileInputStream(file);
        Workbook wbi = new HSSFWorkbook(fis);
        Sheet sheet = wbi.getSheetAt(0);
        Cell cell = sheet.getRow(5).getCell(0);
        cell.setCellValue("9999999999999999");

        fis.close();

        FileOutputStream fos = new FileOutputStream(file);
        wbi.write(fos);
        fos.close();


//
//        FileOutputStream file = new FileOutputStream("C:\\Users\\PKT8\\Desktop\\javaTest.xls");
//        Workbook wb = new HSSFWorkbook();
//        Sheet sheet = wb.createSheet("Sheet 0");
//        Cell cell = sheet.createRow(1).createCell(5);
//        cell.setCellValue("Скала");
//
//        Workbook wb1 = HSSFWorkbook(file);
//
//        wb.write(file);
//        file.close();




    }
}
