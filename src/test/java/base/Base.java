package base;

import java.io.File;
import java.io.FileInputStream;
import java.time.Duration;
import java.util.Properties;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.safari.SafariDriver;

import utilities.Utilities;



public class Base 
{

	WebDriver driver;
	public Properties prop;
	public Properties testDataProp;

	public Base() {

		prop = new Properties();
		File popertiesFile = new File(System.getProperty("user.dir") + "\\src\\main\\java\\configurations\\config.properties");

		Properties testDataProp = new Properties();
		File testDataPopertiesFile = new File(System.getProperty("user.dir") + "\\src\\main\\java\\testdata\\testdata.properties");
		
		try 
		{
			FileInputStream fisTwo = new FileInputStream(testDataPopertiesFile);
			testDataProp.load(fisTwo);
		} 
		catch (Throwable e) 
		{
			e.printStackTrace();
		}

		try 
		{
			FileInputStream fis = new FileInputStream(popertiesFile);
			prop.load(fis);
		} 
		catch (Throwable e) {
			e.printStackTrace();
		}

	}

	public WebDriver initializeBrowserAndVisitUrl(String browserName, String url) 
	{

		if (browserName.equalsIgnoreCase("Chrome")) {
			driver = new ChromeDriver();
		} else if (browserName.equalsIgnoreCase("Edge")) {
			driver = new EdgeDriver();
		} else if (browserName.equalsIgnoreCase("Safari")) {
			driver = new SafariDriver();
		}
		
		//Dimension dimension = new Dimension(1920, 1080);
		//driver.manage().window().setSize(dimension);
		
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(Utilities.IMPLICIT_WAIT_TIME));
		driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(Utilities.PAGE_LOAD_TIME));
		
//		String url = System.getProperty("mainURL", System.getProperty("uatURL"));
//        if (url == null || url.isEmpty()) {
//            throw new IllegalArgumentException("Test URL not specified");
//        }
        
		 if (url.contains("prod")) 
		 {
	            System.out.println("It's not recommended to run automation tests in production environment");
	     } 
		 else 
		 {
			 driver.get(url);
	     }
        
		//driver.get(prop.getProperty("url"));

		return driver;

	}
}
