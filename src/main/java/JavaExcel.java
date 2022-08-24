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
        Workbook outputWb = new HSSFWorkbook();
        Sheet outputSheet = outputWb.createSheet("PriseList");
        FileOutputStream outputFile = new FileOutputStream("C:\\Users\\PKT8\\Desktop\\остатки Спб.xls");


        FileInputStream inputFile = new FileInputStream("C:\\Users\\PKT8\\Desktop\\остатки Питер.xls");
        Workbook inputWb = new HSSFWorkbook(inputFile);
        Sheet inputSheet = inputWb.getSheetAt(0);

        outputWb = inputWb;

        int amountInputRows = countInputRows(inputSheet);

//        setInputCell(inputSheet, outputSheet, amountInputRows);

        Cell inputCell = inputSheet.getRow(8).getCell(5);
        System.out.println(inputCell);
        inputFile.close();

        FileOutputStream outputStream = new FileOutputStream("C:\\Users\\PKT8\\Desktop\\остатки Питер.xls");

        Workbook wb = new HSSFWorkbook();
        Sheet sheetOut = wb.getSheetAt(0);
        Cell cellOut = sheetOut.getRow(8).getCell(5);
        cellOut.setCellValue(999);
        System.out.println(cellOut);
        wb.write(outputStream);
        outputStream.close();




//        int index = 0;
//        for (Row row : inputSheet) {
//            index++;
//            for (Cell cell : row) {
//                String value;
////                Cell cellOutput  = outputSheet.createRow(index).createCell(cell.getStringCellValue());
//                switch (cell.getCellType()) {
//
//                    case STRING:
//                        value = cell.getStringCellValue();
//                        break;
//
//                    case BLANK:
//                        value = "<BLANK>";
//                        break;
//
//                    default:
//                        value = "UNKNOWN value of type " + cell.getCellType();
//                }
//            }
//        }

//        for (int i = 7; i<amountInputRows; i++) {
//            Row inputRow = inputSheet.getRow(i);
//            Cell cellPromo = inputRow.getCell(4);
//            Cell cellSebes = inputRow.getCell(3);
//            int promo = (int) (cellPromo.getNumericCellValue());
//            int sebes = (int) (cellSebes.getNumericCellValue());
//
//            Cell cellOutput = outputSheet.createRow(i).createCell(4);
//
//            if ((sebes*1.4) > promo) {
//                cellOutput.setCellValue(Math.round(promo));
//            }
//            else {
//                cellOutput.setCellValue(Math.round(sebes*1.4));
//            }
//
//
//        }

//        for (int i = 7; i<10; i++) {
//            for (Cell cell : sheet1.getRow(i)) {
//                System.out.println(cell.getNumericCellValue());
//            }
//        }
//        outputWb.write(outputFile);
//        inputWb.write(inputFile);
//        outputFile.close();
//        inputFile.close();
    }

    public static int countInputRows(Sheet inputSheet) {
        int amountInputRows = 0;
        for (Row row: inputSheet)
            amountInputRows++;
        return amountInputRows;
    }

//    public static void setInputCell(Sheet inputSheet, Sheet outputSheet, int amountInputRows) {
//        int i = 7;
//        while (i<amountInputRows) {
//            for (int j = 0; j< 7; j++) {
//                Cell inputCell = inputSheet.getRow(i).getCell(j);
//                Cell outputCell = outputSheet.getRow(i).createCell(j).setCellValue(0);
//                switch (inputCell.getCellType()) {
//                    case STRING:
//                        outputCell.setCellValue(inputSheet.getRow(i).getCell(j).getStringCellValue());
//                        break;
//                    case NUMERIC:
//                        outputCell.setCellValue(inputSheet.getRow(i).getCell(j).getNumericCellValue());
//                        break;
//                    case BLANK:
//                        outputCell.setCellValue(inputSheet.getRow(i).getCell(j).getStringCellValue());
//                        break;
//                }
//            }
//        }
//
//    }
}
