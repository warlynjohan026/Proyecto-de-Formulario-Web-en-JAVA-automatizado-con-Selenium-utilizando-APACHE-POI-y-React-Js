package Tests;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.edge.EdgeDriver;

public class test {

    public static void main(String[] args){

        System.setProperty("webdriver.edge.driver","/Users/warlynpina/Desktop/PruebaSelenium/src/main/resources/msedgedriver");

        WebDriver driver = new EdgeDriver();
        driver.get("http://localhost:8080/ProyectoFinal_war/");




    }

}
