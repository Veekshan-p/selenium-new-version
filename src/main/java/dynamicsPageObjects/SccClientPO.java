package dynamicsPageObjects;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.testng.Assert;

import utilities.Utilities;

public class SccClientPO {
	
	WebDriver driver;
	private Utilities se;

	public SccClientPO(WebDriver driver) 
	{
		
		this.driver = driver;
		PageFactory.initElements(driver, this);
		se = new Utilities(driver);
	}
	
	@FindBy(xpath = "//div[@title='Contacts']//div[@role='presentation']")
    WebElement contactElement;
	
	@FindBy(xpath = "//input[@aria-label='First Name']")
	 WebElement contactFirstName;
	
	@FindBy(xpath = "//span[normalize-space()='Accounts']")
	 WebElement Accounts;
	 
	@FindBy(xpath = "//input[@aria-label='Account Name, Lookup']")
	 WebElement accountName;
	
	@FindBy(xpath = "//input[@aria-label='Last Name']")
	 WebElement contactLastName;
	
	@FindBy(css = "div[title='Sales']")
	WebElement Sales;
	
	@FindBy(xpath = "//div[@title='Leads']")
	private WebElement clickOnLeads;
	
	@FindBy(xpath = "//button[@aria-label='New']")
	private WebElement New;
	
	@FindBy(xpath = "//input[@aria-label='Brand, Lookup']")
	private WebElement brandValue;
	
	@FindBy(xpath = "//ul[@aria-label='Lookup results']//li[1]")
	private WebElement SCCClientResult;
	
	@FindBy(xpath = "//button[@aria-label='Try the new look']")
	private WebElement tryTheNewLook;
	
	@FindBy(xpath = "//button[contains(., 'Switch to the new look')]")
	private WebElement SwitchToNewLook;
	
	@FindBy(xpath = "//input[@aria-label='Topic']")
	WebElement topic;
	
	@FindBy(css = "div[title='Qualify'] div[role='presentation'] div[role='presentation']")
	WebElement Qualify;
	
	@FindBy(xpath = "//div[normalize-space()='Pre-Qualify']")
	WebElement Prequalify;
	
	@FindBy(css = "div[title='Develop'] div[role='presentation'] div[role='presentation']")
	WebElement Develop;
	
	@FindBy(css = "div[title='Propose'] div[role='presentation'] div[role='presentation']")
	WebElement Propose;
	
	@FindBy(css = "div[title='Negotiate'] div[role='presentation'] div[role='presentation']")
	WebElement Negotiate;
	
	@FindBy(css = "div[title='Close'] div[role='presentation'] div[role='presentation']")
	WebElement Close;
	
	@FindBy(xpath = "//input[@aria-label='Source Campaign, Lookup']")
	WebElement SourceCampaign;
	
	@FindBy(xpath = "//li[@aria-label='test campaign 3011v']")
	WebElement campaignName;
	
	@FindBy(xpath = "//input[@aria-label='Lead Type, Lookup']")
	WebElement leadType;
	
	@FindBy(xpath = "//li[@aria-label='New Business']")
	WebElement NewBusiness;
	
	@FindBy(xpath = "//button[@aria-label='Lead source']")
	WebElement leadSource;
	
	@FindBy(xpath = "//div[contains(@id, 'pa-option-set-component')]")
	WebElement leadSourceOption;
	
	@FindBy(xpath = "//button[@title='Go to entity tab to add new']")
    WebElement newAccountRecord;
	
	@FindBy(xpath = "//li[@aria-label='account']//div[@title='Accounts']")
    WebElement newAccountButton;
	
	@FindBy(xpath = "//input[@aria-label='Account Name']")
    WebElement accountNameInput;
	
	@FindBy(xpath = "//input[@aria-label='Annual Revenue']")
    WebElement annualRevenueInput;
	
	@FindBy(xpath = "(//input[@type='text' and @aria-label='Street 1'])[2]")
	WebElement streetOneInput;
	
	@FindBy(xpath = "(//input[@type='text' and @aria-label='Postcode'])[2]")
	WebElement accountPostcodeInput;
	
	@FindBy(xpath = "//button[@aria-label='Save and Close']")
	private WebElement saveAndClose;
	
	@FindBy(xpath = "//input[@aria-label='Email']")
    WebElement emailInput;
	
	@FindBy(xpath = "//input[contains(@id, 'telephone1') and contains(@title, 'Select to enter data')]")
	WebElement businessContactInput;
	
	@FindBy(xpath = "//button[@aria-label='Save (CTRL+S)']")
	WebElement save;
	
	@FindBy(xpath = "//input[@aria-label='Contact, Lookup']")
    WebElement contactLookupInput;
	
	@FindBy(xpath = "//ul[@aria-label='Lookup results']")
    WebElement lookupResultsList;
	
	@FindBy(xpath = "//button[@title='Next Stage']")
	WebElement nextStageButton;
	
	@FindBy(xpath = "//button[@aria-label='Sub Stage']")
    WebElement qualifySubStageButton;
	
	@FindBy(xpath = "//button[@title='Create a timeline record.']")
    WebElement createTimelineRecordButton;
	
	@FindBy(xpath = "//div[normalize-space()='Appointment']")
    WebElement appointmentButton;
	
	@FindBy(xpath = "//button[@aria-label='Appointment Type']")
    WebElement appointmentTypeOption;
	
	@FindBy(xpath = "//div[@role='listbox' and contains(@class, 'fui-Listbox')]/div[@role='option']")
	WebElement appointmentlistOptions;
	
	@FindBy(xpath = "//input[@aria-label='Subject']")
    WebElement subjectInput;
	
	@FindBy(xpath = "(//label[normalize-space()='No'])[1]")
    WebElement teamsMeetingNoInput;
	
	@FindBy(xpath = "//div[@id='headerFieldsFlyout']//span[contains(@class, 'Cancel-symbol')]")
    WebElement cancelOwnerSymbol;
	
	@FindBy(xpath = "//input[@aria-label='Owner, Lookup']")
    WebElement ownerLookupInput;
	
	@FindBy(xpath = "//button[@title='Go back']")
	WebElement goBackButton;
	
	@FindBy(xpath = "//button[@aria-label='Buyer Unit']")
    WebElement buyerUnitInput;
	
	@FindBy(xpath = "//input[@aria-label='Estimated Number of Contractors']")
	WebElement estimatedContractorsInput;
	
	@FindBy(xpath = "//input[@aria-label='Estimated Number of Suppliers']")
    WebElement estimatedSuppliersInput;
	
	@FindBy(xpath = "//input[@aria-label='Number of Sites']")
    WebElement numberOfSitesInput;
	
	@FindBy(xpath = "//button[@aria-label='CSP Uploaded']")
    WebElement cspUploadedOption;
	
	@FindBy(xpath = "//input[@aria-label='Appointment Booker, Lookup']")
    WebElement appointmentBookerLookupInput;
	
	@FindBy(xpath = "//button[@title='More commands for Sharepoint Document']")
    WebElement sharepointDoccumentButton;
	
	@FindBy(xpath = "//span[contains(@class,'pa-') and normalize-space()='Upload']")
    WebElement uploadFileElement;
	
	@FindBy(xpath = "//input[@aria-label='File Upload']")
    WebElement fileUpload;
	
	@FindBy(xpath = "//button[@title='OK']")
    WebElement okButton;
	
	@FindBy(xpath = "//input[@aria-label='Account Name']")
	WebElement accountFiled;
	
	@FindBy(xpath = "//input[@aria-label='Organisation Type, Lookup']")
	WebElement organizationTypeInputField;

	@FindBy(xpath = "//li[@aria-label='Limited Company']//div[@role='presentation']")
	WebElement limitedCompanyOption;
	
	@FindBy(xpath = "//input[@aria-label='Company Registration Number']")
	WebElement companyRegistrationNumberInput;
	
	@FindBy(xpath = "//input[@aria-label='Company Registration Year']")
	WebElement companyRegistrationYearInput;
	
	@FindBy(xpath = "//input[@aria-label='Website']")
	WebElement websiteInput;
	
	@FindBy(xpath = "//input[@aria-label='Street 1']")
	WebElement street1Input;
	
	@FindBy(xpath = "//input[@aria-label='Postcode']")
	WebElement postcodeInput;
	
	@FindBy(xpath = "//div[@aria-label='Contacts']")
	WebElement contactsOption;
	
	@FindBy(xpath = "//button[@title='More commands for Contact']")
	WebElement moreCommandsButton;
	
	@FindBy(xpath = "//span[contains(text(),'New Contact')]")
	WebElement newContactButton;
	
	@FindBy(xpath = "//input[contains(@aria-label,'First Name')]")
	WebElement contactFirstNameInput;
	
	@FindBy(xpath = "//input[@aria-label='Last Name']")
	WebElement contactLastNameInput;
	
	@FindBy(xpath = "//input[@aria-label='Email']")
	WebElement contactEmailInput;
	
	@FindBy(xpath = "//input[@aria-label='Mobile Phone']")
	WebElement contactMobile;
	
	@FindBy(xpath = "(//input[@aria-label='Street 1'])[2]")
	WebElement contactStreet1Input;

	@FindBy(xpath = "//input[@aria-label='ZIP/Postal Code']")
	WebElement contactPostalCodeInput;

	@FindBy(xpath = "//input[@aria-label='First Name']")
	WebElement leadFirstNameInput;

	@FindBy(xpath = "//input[@aria-label='Last Name']")
	WebElement leadLastNameInput;
	
	@FindBy(xpath = "//input[@aria-label='Job Title']")
	WebElement leadJobTitleInput;
	
	@FindBy(xpath = "//input[@aria-label='Email']")
	WebElement leadEmailInput;

	@FindBy(xpath = "//input[contains(@id, 'telephone1.fieldControl-phone-text-input')]")
	WebElement leadTelephoneInput;
	
	@FindBy(xpath = "//input[@aria-label='Sub-Sector, Lookup']")
	WebElement leadSubSectorInput;
	
	@FindBy(xpath = "//input[@aria-label='No of Sites']")
	WebElement leadNumberOfSitesInput;
	
	@FindBy(xpath = "//input[@aria-label='Required, Multiple Selection Lookup']")
	WebElement bDMEmail;
	
	@FindBy(xpath = "//button[@aria-label='Results from: 4 types of records']")
	WebElement appointmentEmailresultsButton;
	
	@FindBy(xpath = "//div[@title='Users']")
	WebElement appointmentEmailUserButton;
	
	@FindBy(xpath = "//button[@title='More Header Editable Fields']")
	WebElement ownerDropdownFieldsButton;
	
	@FindBy(xpath = "//div[@title='Optional - Enter the account, contact, lead, user, or other equipment resources that are not needed at the appointment, but can optionally attend.']")
	WebElement appointmentOptionalResources;

	@FindBy(xpath = "//span[@role='status']")
	WebElement unsavedToSaved;
	
	@FindBy(xpath = "//button[@aria-label='Qualify']")
	WebElement qualifyButton;
	
	@FindBy(xpath = "//button[@aria-label='Sub Stage']")
	WebElement subStageButton;
	
	@FindBy(xpath = "//textarea[@aria-label='Reasons For Rejection']")
	WebElement reasonsForRejectionText;
	
	@FindBy(xpath = "//button[@title='Close']")
	WebElement closeStage;

	@FindBy(xpath = "//button[@aria-label='Forecast Status']")
	WebElement forecastStatus;

	@FindBy(xpath = "(//button[contains(@class, 'fui-CalendarDayGrid__dayButton')])[42]")
	WebElement calendarDay;
	
	@FindBy(xpath = "//input[@aria-label='Date of Est. close date']")
	WebElement dateOfEstCloseDateInput;

	@FindBy(xpath = "//input[@aria-label='Date of Forecast Rack Date']")
	WebElement dateOfForecastRackDate;
	
	@FindBy(xpath = "//input[@aria-label='Actual Number of Contractors']")
	WebElement actualContractors;
	
	@FindBy(xpath = "//input[@aria-label='Actual Number of Suppliers']")
	WebElement actualNumberofSuppliers;

	@FindBy(xpath = "(//button[@aria-label='CSP Part B Uploaded'])[2]")
	WebElement cspPartBUploadedButton;

	@FindBy(xpath = "(//input[@aria-label='Next Stage Account Owner, Lookup'])[2]")
	WebElement nextStageAccountOwnerInput;
	
	@FindBy(xpath = "//span[normalize-space()='SCC Client Annual License']")
	WebElement sccClientAnnualLicense;

	@FindBy(xpath = "//label[normalize-space()='No of Suppliers']")
	WebElement noOfSuppliersLabel;

	@FindBy(xpath = "//span[@data-id='entity_name_span']")
	WebElement entityNameOpp;
	
	@FindBy(xpath = "(//input[@placeholder='Provide the Agent Contact'])[1]")
	WebElement businessPhoneInput;
	
	@FindBy(xpath = "//input[@aria-label='Turnover']")
	WebElement turnOverInput;

	@FindBy(xpath = "(//div[@title='Pre-Qualify'])[1]")
	WebElement preQualifyBPF;
	
	@FindBy(xpath = "(//div[@title='Qualify'])[1]")
	WebElement qualifyBPF;
	
	@FindBy(xpath = "(//div[@title='Develop'])[1]")
	WebElement developBPF;
	
	@FindBy(xpath = "(//div[@title='Propose'])[1]")
	WebElement proposeBPF;
	
	@FindBy(xpath = "(//div[@title='Negotiate'])[1]")
	WebElement negotiateBPF;

	@FindBy(xpath = "//span[contains(text(),'Qualify')]")
	WebElement qualifLead;
	
	@FindBy(xpath = "//label[normalize-space()='Account']")
	WebElement accountLabel;
	
	@FindBy(xpath = "//label[normalize-space()='Company']")
	WebElement companyLabel;
	
	@FindBy(xpath = "//button[@title='YES']")
	WebElement yesButton;

	@FindBy(xpath = "//label[normalize-space()='Description']")
	WebElement descriptionLabel;
	
	@FindBy(xpath = "//label[normalize-space()='Relevant to HubSpot']")
	WebElement hubspotLabel;
	
	@FindBy(xpath = "//label[normalize-space()='Account']")
	WebElement accountLabeltext;
	
	@FindBy(xpath = "//button[@title='Ignore and save']")
	WebElement ignoreAndSaveButton;
	
	/* Test Data for
	 * SCC client sales process
	 */
	
	long currentTime = System.currentTimeMillis() / 10000;
	String contactsFirstName = "contactFirst " + currentTime;
	String contactsLastName = "contactLast " + currentTime;
	String accountsName = "Account " + currentTime;
	String contactEmail = "veekshan.poshala+" + currentTime + "@alcumus.com";
	String phone = "+44" + currentTime;
	String leadFirstName = "Test " + + currentTime;
	String leadLastName = "Name " + + currentTime;
	String leadEmail = "Lead"+ currentTime + "@alcumus.com";
	String regNumber = Long.toString(currentTime % 100000000);
	String testURL = "www"+ currentTime + ".com";
	
	public void createAccountAndContact() throws InterruptedException
	{
		se.clickElement(Sales);
		se.clickElement(Accounts);
		se.clickElement(New);
	//	Thread.sleep(2000);
		se.type(accountFiled, "Test account " + currentTime);
		se.type(turnOverInput, "223344");
		se.type(organizationTypeInputField, "Limited Company");
		se.clickElement(limitedCompanyOption);
		se.type(companyRegistrationNumberInput, regNumber);
		se.type(companyRegistrationYearInput, "2001");
		se.type(numberOfSitesInput, "21");
		se.scrollElementIntoView(websiteInput);
		se.type(websiteInput, testURL);
		se.scrollElementIntoView(street1Input);
		se.type(street1Input, "Bond Street");
		se.type(postcodeInput, "BS1 BS2");
		se.scrollElementIntoView(street1Input);
		se.clickElement(save);
		se.waitForElementTextToChange(unsavedToSaved, "- Saved");
	//	Thread.sleep(2000);
		se.scrollElementIntoView(contactsOption);
		se.clickElement(contactsOption);
		se.clickElement(moreCommandsButton);
		se.clickElement(newContactButton);
		se.type(contactFirstNameInput, contactsFirstName);
		se.type(contactLastNameInput, contactsLastName);
		se.type(contactEmailInput, contactEmail);
		se.scrollElementIntoView(contactMobile);
		se.type(contactMobile, phone);
		se.clickElement(saveAndClose);
		
//		se.clickElement(contactElement);
//		se.type(contactFirstName, contactsFirstName);
//		se.type(contactLastName, contactsLastName);
//		se.clickElement(accountName);
//		se.clickElement(newAccountRecord);
//		se.clickElement(newAccountButton);
//		se.type(accountNameInput, accountsName);
//		se.type(annualRevenueInput, "100000");
//		se.scrollElementIntoView(streetOneInput);
//		se.type(streetOneInput, "Bond street one");
//		se.type(accountPostcodeInput, "W1S 2RB");
//		se.clickElement(saveAndClose);
//		Thread.sleep(2000);
//		se.type(emailInput, contactEmail);
//		se.highlightAndClickElement(save);
		
	}
	
	public void createALead() throws InterruptedException
	{
		se.clickElement(clickOnLeads);
		se.clickElement(New);
		se.type(topic, "Topic Test Data");
		se.type(leadFirstNameInput, leadFirstName);
		se.type(leadLastNameInput, leadLastName);
		se.type(leadType, "New Business");
		se.clickElement(NewBusiness);
		se.type(brandValue, "SCC Client");
		se.clickElement(SCCClientResult);
		se.type(leadJobTitleInput, "SDE");
		se.type(businessPhoneInput, "+44" + currentTime--);
		se.type(leadEmailInput, leadEmail);

		Assert.assertEquals(Prequalify.getText(), "Pre-Qualify", "Element does not display 'Pre-Qualify' text");
		se.highlightElement(Prequalify);
		
		Assert.assertEquals(Qualify.getText(), "Qualify", "Element does not display 'Qualify' text");
		se.highlightElement(Qualify);
		
		Assert.assertEquals(Develop.getText(), "Develop", "Element does not display 'Develop' text");
		se.highlightElement(Develop);
		
		Assert.assertEquals(Propose.getText(), "Propose", "Element does not display 'Propose' text");
		se.highlightElement(Propose);
		
		Assert.assertEquals(Negotiate.getText(), "Negotiate", "Element does not display 'Negotiate' text");
		se.highlightElement(Negotiate);
		
		Assert.assertEquals(Close.getText(), "Close", "Element does not display 'Close' text");
		se.highlightElement(Close);
		
	
		se.type(SourceCampaign, "SCC campaign");
		se.scrollElementIntoView(lookupResultsList);
		se.clickElement(lookupResultsList);
		se.type(contactLookupInput, contactsFirstName + " " + contactsLastName); 
		se.clickElement(lookupResultsList);
		se.waitForElementTextToChange(unsavedToSaved, "- Saved");
		Thread.sleep(5000);
		se.clickElement(leadSource);
		se.scrollElementIntoView(leadSourceOption);
		se.clickOptionIfExists("Partner", leadSourceOption);
		
		se.clickElement(buyerUnitInput);
		se.scrollElementIntoView(leadSourceOption);
		se.clickOptionIfExists("Supply Chain", leadSourceOption);
		se.type(estimatedContractorsInput, "120");
		se.type(estimatedSuppliersInput, "130");
		se.type(leadNumberOfSitesInput, "44");
		se.clickElement(cspUploadedOption);
		se.scrollElementIntoView(leadSourceOption);
		se.clickOptionIfExists("Yes", leadSourceOption);
		se.type(appointmentBookerLookupInput, "Veekshan poshala");
		se.scrollElementIntoView(lookupResultsList);
		se.clickElement(lookupResultsList);
		se.scrollElementIntoView(companyLabel);
		se.type(leadSubSectorInput, "Construction");
		se.scrollElementIntoView(lookupResultsList);
		se.clickElement(lookupResultsList);
		
		se.clickElement(save);
		
	
		
	}
	
	public void businessProcessFlowToQualify() throws InterruptedException
	{
	
		se.clickElement(preQualifyBPF);
		se.clickElement(nextStageButton);
		se.clickElement(closeStage);
		createAnAppointment();
		
		se.refreshPage();
		Thread.sleep(10000);
		
		se.clickElement(qualifyBPF);
		Thread.sleep(3000);
		se.clickElement(qualifySubStageButton);
		Thread.sleep(2000);
		se.clickOptionIfExists("Sales Qualified Lead", leadSourceOption);
		se.clickElement(closeStage);
	
		se.type(leadJobTitleInput, "SDE");
		se.type(businessPhoneInput, "+44" + currentTime--);
		se.scrollElementIntoView(companyLabel);
		se.scrollElementIntoView(turnOverInput);
		se.type(turnOverInput, "223344");
		se.clickElement(save);
		se.clickElement(yesButton);
	
	}
	
	public void createAnAppointment() throws InterruptedException
	{
		
		se.clickElement(createTimelineRecordButton);
		se.clickElement(appointmentButton);
		se.type(bDMEmail, "veekshan poshala");
		se.clickElement(appointmentEmailresultsButton);
		se.clickElement(appointmentEmailUserButton);
		se.clickElement(lookupResultsList);
		se.clickElement(appointmentOptionalResources);
		se.type(subjectInput, "Test subject");
		se.clickElement(appointmentTypeOption);
		se.clickOptionIfExists("Initial Client Meeting", leadSourceOption);
		se.clickElement(teamsMeetingNoInput);
		se.clickElement(ownerDropdownFieldsButton);
		se.clickElement(cancelOwnerSymbol);
		se.type(ownerLookupInput, "veekshan poshala");
		se.clickElement(lookupResultsList);
		se.clickElement(save);
		se.waitForElementTextToChange(unsavedToSaved, "- Saved");
		Thread.sleep(2000);
		se.clickElement(goBackButton);
		
	}
	
	public void qualifyToDevelop() throws InterruptedException

	{
		Thread.sleep(5000);
		se.clickElement(ownerDropdownFieldsButton);
		se.clickElement(cancelOwnerSymbol);
		se.type(ownerLookupInput, "veekshan poshala");
		se.clickElement(lookupResultsList);
		
		se.clickElement(save);
		se.refreshPage();
		Thread.sleep(4000);
		
		se.clickElement(qualifyBPF);
		Thread.sleep(4000);
		se.clickElement(subStageButton);
		Thread.sleep(1000);
		se.clickOptionIfExists("Sales Rejected Lead", leadSourceOption);
		se.clickElement(closeStage);
		se.scrollElementIntoView(accountLabeltext);
		Thread.sleep(1000);
		se.scrollElementIntoView(companyLabel);
		se.type(reasonsForRejectionText, "Test data for rejection");
		se.clickElement(save);
		se.refreshPage();
		Thread.sleep(4000);
		se.clickElement(ownerDropdownFieldsButton);
		Thread.sleep(3000);
		se.clickElement(ownerDropdownFieldsButton);
		se.clickElement(qualifyBPF);
		Thread.sleep(3000);
		se.clickElement(subStageButton);
		Thread.sleep(1000);
		se.clickOptionIfExists("Sales Qualified Lead", leadSourceOption);
		se.clickElement(closeStage);
		se.clickElement(save);
		Thread.sleep(3000);
		se.clickElement(qualifLead);
	
		Thread.sleep(18000);
		
		se.assertElementText(entityNameOpp, "Opportunity");
		se.highlightElement(entityNameOpp);
		
	}
	public void developToNegotiate() throws InterruptedException
	{
		se.clickElement(developBPF);
		Thread.sleep(2000);
		se.clickElement(forecastStatus);
		se.clickOptionIfExists("Pipeline", leadSourceOption);
		se.clickElement(cspPartBUploadedButton);
		se.clickOptionIfExists("Yes", leadSourceOption);
		se.clickElement(dateOfEstCloseDateInput);
		se.clickElement(calendarDay);
		se.clickElement(dateOfForecastRackDate);
		se.clickElement(calendarDay);
		se.type(actualContractors, "30");
		se.type(actualNumberofSuppliers, "45");
		se.clickElement(nextStageButton);
		se.waitForElementTextToChange(unsavedToSaved, "- Saved");
		Thread.sleep(3000);

		se.refreshPage();
		
		se.scrollElementIntoView(descriptionLabel);
		se.assertElementText(sccClientAnnualLicense, "SCC Client Annual License");
		se.highlightElement(sccClientAnnualLicense);
		
	}
		
	public void negotiateToClose() throws InterruptedException
	{	
		Thread.sleep(3000);
		se.clickElement(proposeBPF);
		Thread.sleep(1000);
		se.clickElement(nextStageButton);
		Thread.sleep(5000);
		se.clickElement(forecastStatus);
		Thread.sleep(1000);
		se.clickOptionIfExists("Probable", leadSourceOption);
		se.clickElement(nextStageButton);
		Thread.sleep(5000);
		se.type(nextStageAccountOwnerInput, "Veekshan poshala");
		se.clickElement(lookupResultsList);
		se.clickElement(forecastStatus);
		se.clickOptionIfExists("Commit", leadSourceOption);
	}
	
	
	public void fileUpload() throws InterruptedException 
	{
		driver.get("https://alcumusmain.crm4.dynamics.com/main.aspx?appid=2a66b27d-4a5c-ed11-9562-0022489ae2e0&pagetype=entityrecord&etn=lead&id=a6e8951f-cdd9-4fb3-9d7b-04a7327c06fd");
		Thread.sleep(5000);
		se.clickElement(sharepointDoccumentButton);
		Thread.sleep(1000);
		se.clickElement(uploadFileElement);
		Thread.sleep(2000);
        
        // Specify the relative path to the test document
        String filePath = System.getProperty("user.dir") + "\\src\\main\\java\\testdata\\Test.docx";
        fileUpload.sendKeys(filePath);
        se.clickElement(okButton);
        Thread.sleep(10000);
	}
	
}
