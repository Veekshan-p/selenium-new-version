package AlcumusQa;


import org.openqa.selenium.WebDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Optional;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import base.Base;
import dynamicsPageObjects.AppsPage;
import dynamicsPageObjects.LoginPage;
import dynamicsPageObjects.SafeContractorPO;

public class SafeContractorTest  extends Base {

	public SafeContractorTest() {
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
	
	/*
	 * work in progress
	 */
	
	@Test 
	public void safeContractorSalesProcess() throws InterruptedException 
	{
		AppsPage appspage = new AppsPage(driver);

		appspage.clickOnSCC();
		
		SafeContractorPO safeContractorPO = new SafeContractorPO(driver);
		
		//safeContractorPO.createAccountAndContact();
	
		//safeContractorPO.createCampaignAndAddLeads();
		
		//safeContractorPO.QualifyTheLead();
		
	//	safeContractorPO.addProduct();
		
		safeContractorPO.addDataInExcel();
		
		//safeContractorPO.fileUpload();

	}
	

	@AfterMethod
	public void closeDown() {

		 driver.quit();
	}

}
