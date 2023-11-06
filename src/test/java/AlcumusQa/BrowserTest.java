package AlcumusQa;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

public class BrowserTest {

	
	@Test
	public void getData() throws InterruptedException
	{
		System.out.println("Hello from pipeline-----------------------------------------------------------------------------");
		//System.out.getProperties().list(System.out);
		//System.out.println("Hello from pipeline-----------------------------------------------------------------------------");
		System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
		WebDriver driver =new ChromeDriver();
		driver.get("https://alcumusmain.crm4.dynamics.com/main.aspx");
		//String text =driver.findElement(By.cssSelector("h1")).getText();
		//System.out.println(text);
		//Assert.assertTrue(text.equalsIgnoreCase("RahulShettyAcademy.com Learning"));
		System.out.println("end of pipeline---------------------------------------------------------------------------------");
		driver.close();
	
		
		
	}
	
}