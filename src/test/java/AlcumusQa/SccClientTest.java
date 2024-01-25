package AlcumusQa;

import org.openqa.selenium.WebDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import base.Base;
import dynamicsPageObjects.AppsPage;
import dynamicsPageObjects.LoginPage;
import dynamicsPageObjects.SccClientPO;

public class SccClientTest extends Base {

	public SccClientTest() {
		super();
	}

	public WebDriver driver;

	String email = prop.getProperty("email");
	String password = prop.getProperty("password");
	String browser = prop.getProperty("browser");
	

	@Parameters("URL")
	@BeforeMethod
	public void openBrowser(String URL) throws InterruptedException {

		driver = initializeBrowserAndVisitUrl(prop.getProperty("browser"), URL);

		LoginPage loginPage = new LoginPage(driver);

		loginPage.loginUsing(email, password);

	}

	@Test 
	public void SccClientSalesProcess() throws InterruptedException 
	{
		AppsPage appspage = new AppsPage(driver);

		appspage.clickOnSCC();
		
		SccClientPO SccClient = new SccClientPO(driver);
		
		SccClient.createAccountAndContact();
		SccClient.createALead();
		SccClient.businessProcessFlowToQualify();
		//SccClient.fileUpload();
		SccClient.qualifyToDevelop();
		SccClient.developToNegotiate();
		SccClient.negotiateToClose();

	}
	
	


	@AfterMethod
	public void closeDown() {

		// driver.quit();
	}

}
