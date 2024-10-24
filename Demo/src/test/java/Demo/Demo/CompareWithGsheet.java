package Demo.Demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import io.github.bonigarcia.wdm.WebDriverManager;

public class CompareWithGsheet {
	
	public static String latestSheetPath=null;
	public static String existingSheetPath=null;
	
	
	public void downloadExcel(WebDriver driver) {
		

    // Click on the File menu
		 WebDriverWait wait=new WebDriverWait(driver, 20 );
    WebElement fileMenu = driver.findElement(By.xpath("//div[text()='File']"));
    wait.until(ExpectedConditions.elementToBeClickable(fileMenu));
    fileMenu.click();
   
    WebElement downloadMenu = driver.findElement(By.xpath("//span[text()='ownload']"));
    wait.until(ExpectedConditions.elementToBeClickable(downloadMenu));
    
    
    
    // Wait and click on Download option
   
    downloadMenu.click();

    // Wait and click on Excel option
    WebElement excelOption = driver.findElement(By.xpath("//span[@aria-label='Microsoft Excel (.xlsx) x']"));
    wait.until(ExpectedConditions.elementToBeClickable(excelOption));
    excelOption.click();
	}
	
	 public static Map<String, String> readExcel(String filePath, String sheetName) throws IOException {
	        Map<String, String> dataMap = new HashMap<>();
	        FileInputStream fileInputStream = new FileInputStream(filePath);
	        Workbook workbook = new XSSFWorkbook(fileInputStream);
	        Sheet sheet = workbook.getSheet(sheetName);

	        for (Row row : sheet) {
	            Cell keyCell = row.getCell(0); // Assuming key is in the first column
	            Cell valueCell = row.getCell(1); // Assuming value is in the second column
	            if (keyCell != null && valueCell != null) {
	                dataMap.put(keyCell.toString(), valueCell.toString());
	            }
	        }

	        workbook.close();
	        return dataMap;
	    }
	 public static void ExistingfileCheck(String existingFilePath) {
		 
		 
	 }
	 
	 public static  void compareExcelSheets(String latestSheetPath, String sheet1, String existingSheetPath, String sheet2)  throws IOException {
	        Map<String, String> data1 = readExcel(existingSheetPath, sheet1);
	        Map<String, String> data2 = readExcel(latestSheetPath, sheet2);

	        boolean isEqual = true;

	        for (Map.Entry<String, String> entry : data1.entrySet()) {
	            String key = entry.getKey();
	            String value1 = entry.getValue();
	            String value2 = data2.get(key);

	            if (value2 == null || !value1.equals(value2)) {
	                System.out.println("Mismatch for key: " + key);
	                System.out.println("Value in file1: " + value1);
	                System.out.println("Value in file2: " + value2);
	                isEqual = false;
	            }
	        }

	        for (Map.Entry<String, String> entry : data2.entrySet()) {
	            String key = entry.getKey();
	            if (!data1.containsKey(key)) {
	                System.out.println("Key present in file2 but not in file1: " + key);
	                isEqual = false;
	            }
	        }

	        if (isEqual) {
	            System.out.println("The sheets are identical.");
	            
	            
	          
	        } else {
	            System.out.println("The sheets have differences.");
	        }
	    }
	 
	    public static void main(String[] args) {
	    	WebDriverManager.chromedriver().setup();

			WebDriver driver = new ChromeDriver();
			
		driver.get("https://docs.google.com/spreadsheets/d/1zCzoqMyECC4T4K5UbjXv2xwTd5nQSf-H/edit?gid=1910714829#gid=1910714829");
		
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	        
	        CompareWithGsheet compareWithGsheet=new CompareWithGsheet();
	        compareWithGsheet.downloadExcel(driver);
	        latestSheetPath="C://Users//Karthik//Downloads//actiTimeScenarios(1).xlsx" ;//pass the downloaded path
	        existingSheetPath= "C://Users//Karthik//Downloads//actiTimeScenarios (1).xlsx";// pass the Existing sheet path
	        
	        if (new File(existingSheetPath).exists()) {
	        	 // Compare Excel sheets
		        try {
		        	
		            compareExcelSheets(
		                "existingSheetPath", "Sheet1",
		                "latestSheetPath", "Sheet1"
		            );
		            
		            //Delete the old file
		            new File(existingSheetPath).delete();
		            
		            //Renaming the file 
		            new File(latestSheetPath).renameTo(new File(existingSheetPath));
		            
		            
		        } catch (IOException e) {
		            e.printStackTrace();
		        }
	        	
	        }
	        else{
	        	System.out.println("We dont have the existing file will proceed further");
	        }

	       

	        driver.quit();
	    }
	

}
