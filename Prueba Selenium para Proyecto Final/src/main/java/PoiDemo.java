import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class PoiDemo {

    public static void main(String[] args) throws IOException {

        String filePath = "/Users/warlynpina/Desktop/PruebaSelenium/hola.xlsx";
        File myFile = new File(filePath);
        Workbook workbook = new XSSFWorkbook();

        FileOutputStream fout = new FileOutputStream(myFile);
        workbook.write(fout);
        fout.close();
        System.out.println("File created");


    }
}
