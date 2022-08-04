package Tests;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.concurrent.TimeUnit;
import com.excel.lib.util.Xls_Reader;


public class AutomationProyectoFinal {


    public static void main(String[] args) throws IOException, InterruptedException {



        System.setProperty("webdriver.edge.driver","/Users/warlynpina/Desktop/College/2022-02/Desarrollo de Software con Tecnología Open Source/Prueba Selenium para Proyecto Final/src/main/resources/msedgedriver");
        WebDriver driver = new EdgeDriver();
        driver.get("http://localhost:8080/ProyectoFinal_war/");
        driver.manage().timeouts().implicitlyWait(5, TimeUnit.MINUTES );




        WebElement nom = driver.findElement(By.id("nombrepropietario"));
        WebElement numcontact = driver.findElement(By.id("numerodecontacto"));
        WebElement emails = driver.findElement(By.id("email"));
        WebElement direccions = driver.findElement(By.id("direccion"));
        WebElement numplaca = driver.findElement(By.id("numeroplaca"));
        WebElement marcas = driver.findElement(By.id("marca"));
        WebElement modelos = driver.findElement(By.id("modelo"));
        WebElement año = driver.findElement(By.id("año"));
        WebElement km = driver.findElement(By.id("k/m"));
        WebElement servicios = driver.findElement(By.id("servicio"));
        WebElement anteriors = driver.findElement(By.id("anterior"));
        WebElement fecha = driver.findElement(By.id("date"));
        WebElement comentarioss = driver.findElement(By.id("comentario"));
        WebElement cargar = driver.findElement(By.id("cargar"));





        Xls_Reader reader = new  Xls_Reader ("/Users/warlynpina/Desktop/College/2022-02/Desarrollo de Software con Tecnología Open Source/Prueba Selenium para Proyecto Final/src/main/resources/Excel_Para_Cargar.xlsx");
        String sheetName = "Final";

        int rowCount = reader.getRowCount(sheetName);

        for(int rowNum=2; rowNum<=rowCount; rowNum++){

            String nombre = reader.getCellData(sheetName, "NombrePropietario", rowNum);
            String contacto = reader.getCellData(sheetName, "NumContacto", rowNum);
            String email = reader.getCellData(sheetName, "Email", rowNum);
            String direccion = reader.getCellData(sheetName, "Direccion", rowNum);
            String placa = reader.getCellData(sheetName, "NumeroPlaca", rowNum);
            String marca = reader.getCellData(sheetName, "Marca", rowNum);
            String modelo = reader.getCellData(sheetName, "Modelo", rowNum);
            String tiempo = reader.getCellData(sheetName, "Año", rowNum);
            String millas = reader.getCellData(sheetName, "KM/MI", rowNum);
            String servicio = reader.getCellData(sheetName, "Servicio", rowNum);
            String anterior = reader.getCellData(sheetName, "MantenimientoAnterior", rowNum);
            String comentarios = reader.getCellData(sheetName, "Comentarios", rowNum);
            String Fecha = reader.getCellData(sheetName, "Fecha", rowNum);


            Thread.sleep(3000);
            driver.findElement(By.id("cargar")).click();

            synchronized (driver){
                driver.wait(2000);

            }



            nom.sendKeys(nombre);
            numcontact.sendKeys(contacto);
            emails.sendKeys(email);
            direccions.sendKeys(direccion);
            numplaca.sendKeys(placa);
            marcas.sendKeys(marca);
            modelos.sendKeys(modelo);
            año.sendKeys(tiempo);
            km.sendKeys(millas);
            servicios.sendKeys(servicio);
            anteriors.sendKeys(anterior);
            comentarioss.sendKeys(comentarios);
            fecha.sendKeys(Fecha);


        }

        driver.findElement(By.id("generar")).click();

        String filePath = "/Users/warlynpina/Desktop/College/2022-02/Desarrollo de Software con Tecnología Open Source/Prueba Selenium para Proyecto Final/Creación de Nuevo Excel.xlsx";
        File myFile = new File(filePath);
        Workbook workbook = new XSSFWorkbook();

        Sheet sheet = workbook.createSheet("Creación de Nueva Hoja");

        XSSFCell cell0 = (XSSFCell) sheet.createRow(0).createCell(0);

        cell0.setCellValue("Fecha y Hora de creacioón de la Solicitud ");

        XSSFCreationHelper creationHelper = (XSSFCreationHelper) workbook.getCreationHelper();

        CellStyle style1 = workbook.createCellStyle();
        style1.setDataFormat(creationHelper.createDataFormat().getFormat("dd-mm-yyyy hh:mm:ss"));

        XSSFCell cell1 = (XSSFCell) sheet.createRow(1).createCell(0);
        cell1.setCellValue(new Date());
        cell1.setCellStyle(style1);





        FileOutputStream fout = new FileOutputStream(myFile);
        workbook.write(fout);
        fout.close();
        System.out.println("File created");



    }

}
