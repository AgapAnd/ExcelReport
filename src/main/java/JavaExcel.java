import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class JavaExcel {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("PriseList");


        for (int i =  0; i<5; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j<5; j++) {
                row.createCell(j).setCellValue(i * 5 + j + 1);
            }
        }

        FileOutputStream file = new FileOutputStream("C:\\Users\\PKT8\\Desktop\\остатки СПБ.xls");

        wb.write(file);
        file.close();

        FileInputStream inputFile = new FileInputStream("C:\\Users\\PKT8\\Desktop\\остатки Питер.xls");
        Workbook inputWb = new HSSFWorkbook(inputFile);
        for (int i = 0; i<3; i++) {
            int value = (int) (inputWb.getSheetAt(0).getRow(7+i).getCell(4).getNumericCellValue());
            System.out.println(value);
        }
        inputFile.close();
    }
}
