package AlcumusQa;

import org.testng.annotations.Test;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class FirstTest {
	
	@Test
	public void Login()
	{
		
		System.out.println("Hello from pipeline-------------------------------------------------------------");
		WebDriver driver = new ChromeDriver();
		driver.get("https://alcumusmain.crm4.dynamics.com/main.aspx");
		System.out.println("End of pipeline-------------------------------------------------------------");
		driver.close();
	}

}
