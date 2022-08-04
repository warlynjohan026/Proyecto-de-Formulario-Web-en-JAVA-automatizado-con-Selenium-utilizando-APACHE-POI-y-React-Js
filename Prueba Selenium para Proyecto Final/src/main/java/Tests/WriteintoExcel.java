package Tests;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class WriteintoExcel {

    public void writeData(String excelPath, String sheetName, int rowNumber, int columnNumber, String data) throws IOException {

        File file = new File(excelPath);
        FileInputStream inputStream = new FileInputStream(file);
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.getSheet(sheetName);
        XSSFRow row = sheet.getRow(rowNumber);
        XSSFCell cell = row.getCell(columnNumber, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

        if(cell==null){

            row.createCell(columnNumber);
            cell.setCellValue(data);

        } else{

            cell.setCellValue(data);
        }

        FileOutputStream fio = new FileOutputStream(file);
        wb.write(fio);
        wb.close();






    }
}
