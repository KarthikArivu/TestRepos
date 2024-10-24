package Demo.Demo;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class GoogleSheetDownloader {
	
	
	public static void main(String[] args) {
		WebDriverManager.chromedriver().setup();

		WebDriver driver = new ChromeDriver();
		
	driver.get("https://docs.google.com/spreadsheets/d/1zCzoqMyECC4T4K5UbjXv2xwTd5nQSf-H/edit?gid=1910714829#gid=1910714829");
	
	driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

    // Click on the File menu
    WebElement fileMenu = driver.findElement(By.xpath("//div[text()='File']"));
    fileMenu.click();

    // Wait and click on Download option
    WebElement downloadMenu = driver.findElement(By.xpath("//span[text()='ownload']"));
    downloadMenu.click();

    // Wait and click on Excel option
    WebElement excelOption = driver.findElement(By.xpath("//span[@aria-label='Microsoft Excel (.xlsx) x']"));
    excelOption.click();
	
		
	}

}
