import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

public class ReportExcel {
    public static void main(String[] args) throws IOException {

        final int BEGIN_NUMBERS = 7;

        Scanner scan = new Scanner(System.in);
        System.out.println("Введите общую наценку в процентах:");
        float k = 1 + scan.nextFloat()/100;

        System.out.println("Введите наценку на спецтехнику в процентах:");
        float kSpec = 1 + scan.nextFloat()/100;


        File file = new File("C:\\Users\\PKT8\\Desktop\\остатки Питер.xls");
        FileInputStream fis = new FileInputStream(file);
        Workbook wb = new HSSFWorkbook(fis);
        Sheet sheet = wb.getSheetAt(0);

        makingPrice(sheet, BEGIN_NUMBERS, k, kSpec);

        makingHeader(sheet);

//        setCellStyle(wb, sheet);




        fis.close();

        FileOutputStream fos = new FileOutputStream(file);
        wb.write(fos);
        fos.close();
    }


    public static void makingPrice(Sheet sheet, int beginNumbers, float k, float kSpec) {
        int amountInputRows = sheet.getLastRowNum();

        for (int i = beginNumbers; i< amountInputRows+1; i++) {

            long sebes = Math.round(sheet.getRow(i).getCell(3).getNumericCellValue());
            long promo = Math.round(sheet.getRow(i).getCell(4).getNumericCellValue());

            Cell cell = sheet.getRow(i).getCell(3);
            String brend = sheet.getRow(i).getCell(0).getStringCellValue();

            if ((sebes*k)>promo || brend.equals("WAYTEKO PREMIUM") || brend.equals("Rostar"))
                cell.setCellValue(promo);

            else if (brend.equals("Deutz") || brend.equals("Liugono") || brend.equals("Lonking") || brend.equals("n/n") ||
                    brend.equals("SDLG") || brend.equals("Shanghai") || brend.equals("Shantui") || brend.equals("Steyr") ||
                    brend.equals("XCMG") || brend.equals("Yuchai") || brend.equals("ZF")) {
                cell.setCellValue(Math.round(sebes * kSpec));
            }

            else
                cell.setCellValue(Math.round(sebes * k));

            sheet.getRow(i).getCell(4).setCellValue(sheet.getRow(i).getCell(5).getStringCellValue());

            sheet.getRow(i).getCell(5).setCellValue(sheet.getRow(i).getCell(6).getStringCellValue());
            sheet.getRow(i).getCell(6).setCellValue("");
        }

    }

    public static void makingHeader(Sheet sheet) {
        sheet.getRow(5).getCell(3).setCellValue("Спец цена");
        sheet.getRow(6).getCell(3).setCellValue("");
        sheet.getRow(5).getCell(4).setCellValue("Наличие");
        sheet.getRow(6).getCell(4).setCellValue("");
        sheet.getRow(5).getCell(5).setCellValue("Код");
        sheet.getRow(6).getCell(5).setCellValue("");
        sheet.getRow(5).getCell(6).setCellValue("Заказ");
        sheet.getRow(6).getCell(6).setCellValue("");
        sheet.setColumnWidth(5,5_000);

    }

    public static void setCellStyle(Workbook wb, Sheet sheet) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);

        for (Cell cell : sheet.getRow(5)) {
            cell.setCellStyle(style);
        }
    }

}
