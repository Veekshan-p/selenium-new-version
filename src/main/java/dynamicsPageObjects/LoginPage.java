package dynamicsPageObjects;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import utilities.Utilities;


public class LoginPage {

	WebDriver driver;
	
	private Utilities utilities;
	

	public LoginPage(WebDriver driver) 
	{
		this.driver = driver;
		PageFactory.initElements(driver, this);
		utilities = new Utilities(driver);
		
	}
	
	@FindBy(xpath = "//input[@type='email']")
	private WebElement Email;

	@FindBy(xpath = "//input[@type='password']")
	private WebElement Password;

	@FindBy(xpath = "//input[@type='submit']")
	private WebElement Next;

	@FindBy(xpath = "//input[@name='DontShowAgain']")
	private WebElement DoNotShowAgain;


	// actions
	
	public void clickOnNext() throws InterruptedException 
	{
		utilities.clickElement(Next);
	}

	public void clickOnDoNotShowAgain() throws InterruptedException {

		utilities.clickElement(DoNotShowAgain);
		//DoNotShowAgain.click();
	}

	public void loginUsing(String email, String password) throws InterruptedException {
		
		utilities.type(Email, email);
		clickOnNext();
		utilities.type(Password, password);
		clickOnNext();
		Thread.sleep(2000);
		clickOnDoNotShowAgain();
		clickOnNext();
	}

}
