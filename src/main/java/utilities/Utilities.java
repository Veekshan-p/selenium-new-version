package utilities;

import java.io.File;
import java.io.IOException;
import java.time.Duration;
import java.util.Date;
import java.util.List;
import java.util.Properties;

import org.openqa.selenium.By;
import org.openqa.selenium.ElementClickInterceptedException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.io.FileHandler;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

public class Utilities {

	WebDriver driver;
	private WebDriverWait wait;
    private Actions actions;
    private JavascriptExecutor js;

	public Utilities(WebDriver driver) 
	{
		this.driver = driver;
		wait = new WebDriverWait(driver, Duration.ofSeconds(15));
        actions = new Actions(driver);
        js = (JavascriptExecutor) driver;
	}

	public Properties prop;
	public static final int IMPLICIT_WAIT_TIME = 20;
	public static final int PAGE_LOAD_TIME = 20;

	public static String generateEmailWithTimeStamp() 
	{
		
		Date date = new Date();
		String timestamp = date.toString().replace(" ", "_").replace(":", "_");
		return "veekshan.poshala" + timestamp + "@alcumus.com";
		
	}

	public static String captureScreenshot(WebDriver driver, String testName) 
	{

		File srcScreenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		String destinationScreenshotPath = System.getProperty("user.dir") + "\\Screenshots\\" + testName + ".png";

		try 
		{
			FileHandler.copy(srcScreenshot, new File(destinationScreenshotPath));
		} 
		catch (IOException e) 
		{
			e.printStackTrace();
		}

		return destinationScreenshotPath;
	}
	
	public void clickElement(WebElement element) throws InterruptedException 
	{
		boolean elementFound = false;

        for (int attempt = 1; attempt <= 3; attempt++) 
        {
            try 
			{
                // Check if the element is present using ExpectedConditions
            	wait.until(ExpectedConditions.visibilityOf(element));

                // Perform the action on the element
            	element.click();

                elementFound = true; // Set flag to true if element is found
                break; // Exit the loop if successful

            } 
            catch (Exception e) 
            {
               //  Handle the exception if the element is not found
               System.out.println("Element not found in attempt " + attempt + ": " + e.getMessage());
                // You can add a delay between retries if needed
               Thread.sleep(3000); // 1 second delay between retries
               element.click();
            }
        }
	}

	public void assertElementText(WebElement element, String expectedText) throws InterruptedException 
	{
		boolean elementFound = false;

        for (int attempt = 1; attempt <= 3; attempt++) 
        {
            try 
			{
                // Check if the element is present using ExpectedConditions
            	wait.until(ExpectedConditions.visibilityOf(element));

                // Perform the action on the element
            	String actualText = element.getText();
            	
      	        Assert.assertEquals(actualText, expectedText, "Element does not display '" + expectedText + "' text");

                elementFound = true; // Set flag to true if element is found
                break; // Exit the loop if successful

            } 
            catch (Exception e) 
            {
                // Handle the exception if the element is not found
                System.out.println("Element not found in attempt " + attempt + ": " + e.getMessage());
                // You can add a delay between retries if needed
                Thread.sleep(1000); // 1 second delay between retries
            }
        }
	        
	}

	public void highlightAndClickElement(WebElement element) throws InterruptedException 
	{
		wait.until(ExpectedConditions.visibilityOf(element));
		
		String originalStyle = element.getAttribute("style");
		// Highlight the element by changing its style
		js.executeScript("arguments[0].setAttribute('style', 'border: 4px solid Orange;');", element);
		// Click the highlighted element
		element.click();
        // Remove the added style to de-highlight and restore the original style
        js.executeScript("arguments[0].setAttribute('style', arguments[1]);", element, originalStyle);
	}
	
	public void highlightElement(WebElement element) throws InterruptedException 
	{
		wait.until(ExpectedConditions.visibilityOf(element));
		
		String originalStyle = element.getAttribute("style");
		// Highlight the element by changing its style
		js.executeScript("arguments[0].setAttribute('style', 'border: 4px solid Orange;');", element);
		Thread.sleep(200); // 0.3 seconds (300 milliseconds)
        // Remove the added style to de-highlight and restore the original style
        js.executeScript("arguments[0].setAttribute('style', arguments[1]);", element, originalStyle);
	}
	
	public void waitForElementTextToChange(WebElement element, String expectedText) throws InterruptedException 
	{
		
	        boolean elementFound = false;

	        for (int attempt = 1; attempt <= 3; attempt++) {
	            try 
				{
	                // Check if the element is present using ExpectedConditions
	                
	            	wait.until(ExpectedConditions.textToBePresentInElement(element, expectedText));
	            	
	                elementFound = true; // Set flag to true if element is found
	                break; // Exit the loop if successful

	            } catch (Exception e) {
	                // Handle the exception if the element is not found
	                System.out.println("Element not found in attempt " + attempt + ": " + e.getMessage());

	                // You can add a delay between retries if needed
	                Thread.sleep(1000); // 1 second delay between retries
	            }
	        }
	}
	
	public void waitForElementToBeInvisible(WebElement element) {
	       
	        wait.until(ExpectedConditions.invisibilityOf(element));
	    }
	
	public void waitForElementToLoad(WebElement element) 
	{
		wait.until(ExpectedConditions.visibilityOf(element));
	}
	
	public void type(WebElement element, String textToBeTyped) throws InterruptedException 
	{
		boolean elementFound = false;

        for (int attempt = 1; attempt <= 3; attempt++) 
        {
            try 
			{
                // Check if the element is present using ExpectedConditions
            	wait.until(ExpectedConditions.visibilityOf(element));

                // Perform the action on the element
            	element.click();
            	element.clear();
            	element.sendKeys(textToBeTyped);

                elementFound = true; // Set flag to true if element is found
                break; // Exit the loop if successful

            } 
            catch (Exception e) 
            {
                // Handle the exception if the element is not found
              //  System.out.println("Element not found in attempt " + attempt + ": " + e.getMessage());
                // You can add a delay between retries if needed
                Thread.sleep(1000); // 1 second delay between retries
            }
        }
	}
	
	public void highlightAndType(WebElement element, String textToBeTyped) 
	{
		wait.until(ExpectedConditions.elementToBeClickable(element));
		
		String originalStyle = element.getAttribute("style");
		// Highlight the element by changing its style
		js.executeScript("arguments[0].setAttribute('style', 'border: 4px solid Orange;');", element);
		// Click the highlighted element
		element.click();
		element.clear();
		element.sendKeys(textToBeTyped);
		js.executeScript("arguments[0].setAttribute('style', arguments[1]);", element, originalStyle);
	}

	public void selectOptionInDropdown(WebElement element,String dropDownOption) 
	{
		Select select = new Select(element);
		select.selectByVisibleText(dropDownOption);
	}
	
	public void clickOptionIfExists(String optionText, WebElement element) throws InterruptedException 
	{
		boolean elementFound = false;

        for (int attempt = 1; attempt <= 3; attempt++) 
        {
            try 
			{
            	wait = new WebDriverWait(driver, Duration.ofSeconds(15));
                // Check if the element is present using ExpectedConditions
            	wait.until(ExpectedConditions.visibilityOf(element));

                // Perform the action on the element
            	List<WebElement> options = element.findElements(By.xpath(".//*"));

        	    for (WebElement option : options) 
        	    {
        	        if (option.getText().equals(optionText)) 
        	        {
        	        	option.click();
        	        	//se.clickElement(option);
        	            break;
        	        }
        	    }

                elementFound = true; // Set flag to true if element is found
                break; // Exit the loop if successful

            } 
            catch (Exception e) 
            {
                // Handle the exception if the element is not found
             //   System.out.println("Element not found in attempt " + attempt + ": " + e.getMessage());
                // You can add a delay between retries if needed
                Thread.sleep(1000); // 1 second delay between retries
            }
        }
	    
	}
	
	public void dynamicWait(WebElement element) 
	{
		Wait<WebDriver> wait = new FluentWait<>(driver)
				.withTimeout(Duration.ofSeconds(20)) // 20 seconds timeout
				.pollingEvery(Duration.ofMillis(500)) // Check every 500 milliseconds
				.ignoring(NoSuchElementException.class)
                .ignoring(ElementClickInterceptedException.class)
                .ignoring(StaleElementReferenceException.class);
		
		wait.until(ExpectedConditions.visibilityOf(element));
	}
	
	public void scrollToElement(WebElement element) 
	{
	 
	    js.executeScript("arguments[0].scrollIntoView(true);", element);
	}
	
	public void scrollElementIntoView(WebElement element) 
	{
	  
	    js.executeScript("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'start' });", element);
	}
	
	public void clickElementWithJS(WebElement element) 
	{
	    js.executeScript("arguments[0].click();", element);
	    // can be used when click() method is not working with Selenium WebDriver for various reasons
	}
	
	public void typeTextWithJS(WebElement element, String text) 
	{
	    js.executeScript("arguments[0].value = arguments[1];", element, text);
	}
	
	public void doubleClickElement(WebElement element) 
	{
	    actions.doubleClick(element).build().perform();
	}
	
	public void refreshPage() 
	{
	    driver.navigate().refresh();
	}
	 public void endOfPage(WebElement element)
	 {
		 js.executeScript("window.scrollTo(0, document.body.scrollHeight)");
		 wait.until(ExpectedConditions.visibilityOf(element));
	 }

//	public static Object[][] getTestDataFromExcel(String sheetName) 
//	{
//		File excelFile = new File(System.getProperty("user.dir")+"\\src\\main\\java\\testdata\\TestData.xlsx");
//		
//		XSSFWorkbook workbook = null;
//		
//		try 
//		{
//			FileInputStream fisExcel = new FileInputStream(excelFile);
//			workbook = new XSSFWorkbook(fisExcel);
//		} 
//		catch (Throwable e) 
//		{
//			e.printStackTrace();
//		}
//
//		XSSFSheet sheet = workbook.getSheet(sheetName);
//
//		int rows = sheet.getLastRowNum();
//		int cols = sheet.getRow(0).getLastCellNum();
//
//		Object[][] data = new Object[rows][cols];
//
//		for (int i = 0; i < rows; i++) 
//		{
//
//			XSSFRow row = sheet.getRow(i + 1);
//
//			for (int j = 0; j < cols; j++) 
//			{
//
//				XSSFCell cell = row.getCell(j);
//				CellType cellType = cell.getCellType();
//
//				switch (cellType) 
//				{
//
//				case STRING:
//					data[i][j] = cell.getStringCellValue();
//					break;
//				case NUMERIC:
//					data[i][j] = Integer.toString((int) cell.getNumericCellValue());
//					break;
//				case BOOLEAN:
//					data[i][j] = cell.getBooleanCellValue();
//					break;
//
//				}
//
//			}
//
//		}
//		return data;
//	}
}
