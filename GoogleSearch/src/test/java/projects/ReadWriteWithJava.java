package projects;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import io.github.bonigarcia.wdm.WebDriverManager;

public class ReadWriteWithJava {

	public static void main(String[] args) throws InterruptedException {
		
		WebDriverManager.chromedriver();
		WebDriverManager.firefoxdriver();
		WebDriverManager.edgedriver();
		
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://www.google.com/");
		
		WebElement searchBox = driver.findElement(By.id("APjFqb"));
		searchBox.sendKeys("Dhaka");
		
		
		
		Thread.sleep(5000);
		
		
		List <WebElement> suggestionList = driver.findElements(By.xpath("//ul[@role='listbox']//li/descendant::div[@class='wM6W7d']"));
		//List <WebElement> suggestionList = driver.findElements(By.xpath("//ul[@role='listbox']//li/descendant::div[@class='wM6W7d']"));
		
		
		for (WebElement suggestion : suggestionList) {
			String suggestionText = suggestion.getText();
			System.out.println(suggestionText);
        }
		
		
		// Find the largest and smallest suggestion text
        List<String> suggestionTexts = new ArrayList<>();
        for (WebElement suggestion : suggestionList) {
            String suggestionText = suggestion.getText();
            suggestionTexts.add(suggestionText);
        }

        // Print largest and smallest suggestion text
        String largestSuggestion = Collections.max(suggestionTexts);
        String smallestSuggestion = Collections.min(suggestionTexts);

        System.out.println("Largest Suggestion: " + largestSuggestion);
        System.out.println("Smallest Suggestion: " + smallestSuggestion);
        
        
        //Create a New WorkBook
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("suggestions");
        
        
        //Create The Header Row
        Row headerRow =sheet.createRow(0);
        headerRow.createCell(1).setCellValue("largest Suggestion");
        headerRow.createCell(2).setCellValue("smallest Suggestion");
        
        //Create the data Row
        Row dataRow = sheet.createRow(1);
        dataRow.createCell(1).setCellValue(largestSuggestion);
        dataRow.createCell(2).setCellValue(smallestSuggestion);
        
//     // Update the row with suggestion data
//        dataRow.createCell(1).setCellValue(smallestSuggestion);
//        dataRow.createCell(1 + 1).setCellValue(largestSuggestion);
        
        //Auto-size the columns
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        
        
        //Save the workbook to an XLSX File 
        try {
        	FileOutputStream outputStream = new FileOutputStream("C:\\Users\\ShawoN\\Desktop\\suggestions.xlsx");
        	workbook.write(outputStream);
        	workbook.close();
        	outputStream.close();
        	System.out.println("Suggestions saved to suggestions.xlsx");	
        } catch (IOException e) {
        	e.printStackTrace();
        }
        

        // Quit WebDriver
        driver.quit();
      }
		
	}


