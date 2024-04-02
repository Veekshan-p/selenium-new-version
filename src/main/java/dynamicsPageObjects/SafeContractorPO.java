package dynamicsPageObjects;




import java.awt.AWTException;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashSet;
import java.util.Random;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.testng.Assert;

import utilities.Utilities;

public class SafeContractorPO {
	
	WebDriver driver;
	private Utilities se;

	public SafeContractorPO(WebDriver driver) 
	{
		
		this.driver = driver;
		PageFactory.initElements(driver, this);
		se = new Utilities(driver);
	}
	
	@FindBy(css = "div[title='Sales']")
	WebElement Sales;
	
	@FindBy(css = "div[title='Campaigns']")
	WebElement Campaigns;
	
	@FindBy(xpath = "//button[@aria-label='New']")
	private WebElement New;
	
	@FindBy(xpath = "//input[@aria-label='Name']")
	WebElement campaignNameInput;
	
	@FindBy(xpath = "//button[@aria-label='Campaign Type']")
	WebElement campaignType;
	
	@FindBy(xpath = "//div[contains(@id, 'pa-option-set-component')]")
	WebElement leadSourceOption;
	
	@FindBy(xpath = "//input[@aria-label='Client, Lookup']")
	WebElement clientLookupInput;
	
	@FindBy(xpath = "//button[@aria-label='New Account']")
	WebElement newAccountButton;
	
	@FindBy(xpath = "//input[@aria-label='Account Name']")
	WebElement accountNameInput;
	
	@FindBy(xpath = "//span[normalize-space()='Accounts']")
	 WebElement Accounts;

	@FindBy(xpath = "//input[@aria-label='Account Name']")
	WebElement accountFiled;
	
	@FindBy(xpath = "//input[@aria-label='Turnover']")
	WebElement turnOverInput;
	
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
	
	@FindBy(xpath = "//input[@aria-label='Number of Sites']")
    WebElement numberOfSitesInput;
	
	@FindBy(xpath = "//button[@aria-label='Save (CTRL+S)']")
	WebElement save;
	
	@FindBy(xpath = "//span[@role='status']")
	WebElement unsavedToSaved;
	
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
	
	@FindBy(xpath = "//button[@aria-label='Save and Close']")
	private WebElement saveAndClose;
	
	@FindBy(xpath = "//ul[@aria-label='Lookup results']")
    WebElement lookupResultsList;
	
	@FindBy(xpath = "(//button[contains(@class, 'fui-CalendarDayGrid__dayButton')])[42]")
	WebElement calendarDay;
	
	@FindBy(xpath = "//input[@aria-label='Date of Pre-note Date']")
	WebElement preNoteDateInput;
	
	@FindBy(xpath = "//li[@aria-label='SafeContractor, 27/10/2022 08:40']")
	WebElement safeContractorOption;
	
	@FindBy(xpath = "//li[@aria-label='SafeContractor, 17/11/2022 15:53']")
	private WebElement safeContractorUatOption;

	@FindBy(xpath = "//input[@aria-label='Brand, Lookup']")
	WebElement brandLookUpValue;
	
	@FindBy(css = "li[title='Leads'][aria-label='Leads'][role='tab']")
	WebElement leadsTab;
	
	@FindBy(xpath = "//button[@aria-label='New Lead. Add New Lead']")
	WebElement newLeadButton;
	
	@FindBy(xpath = "//input[@aria-label='Topic']")
	WebElement inputTopic;
	
	@FindBy(xpath = "//input[@aria-label='First Name']")
	WebElement leadFirstName;
	
	@FindBy(xpath = "//input[@aria-label='Last Name']")
	WebElement leadLastName;
	
	@FindBy(xpath = "//input[@aria-label='Lead Type, Lookup']")
	WebElement leadType;
	
	@FindBy(xpath = "//li[@aria-label='New Business']")
	WebElement NewBusiness;
	
	@FindBy(xpath = "//input[@aria-label='Job Title']")
	WebElement leadJobTitleInput;
	
	@FindBy(xpath = "(//input[@placeholder='Provide the Agent Contact'])[1]")
	WebElement businessPhoneInput;
	
	@FindBy(xpath = "//input[@aria-label='Email']")
	WebElement leadEmailInput;
	
	@FindBy(xpath = "//input[@aria-label='Contact, Lookup']")
    WebElement contactLookupInput;
	
	@FindBy(xpath = "//button[@aria-label='Lead source']")
	WebElement leadSource;
	
	@FindBy(xpath = "//button[@title='Go back']")
	WebElement goBackButton;
	
	@FindBy(xpath = "//button[@title='Ignore and save']")
	WebElement ignoreAndSaveButton;
	
	@FindBy(xpath = "(//div[@title='Pre-Qualify'])[1]")
	WebElement preQualifyBPF;
	
	@FindBy(xpath = "//button[@title='Next Stage']")
	WebElement nextStageButton;
	
	@FindBy(xpath = "//button[@title='Close']")
	WebElement closeStage;
	
	@FindBy(xpath = "//input[@aria-label='Number of Employees']")
	WebElement numberOfEmployees;
	
	@FindBy(xpath = "(//div[@title='Qualify'])[1]")
	WebElement qualifyBPF;
	
	@FindBy(xpath = "//button[@aria-label='Sub Stage']")
    WebElement qualifySubStageButton;
	
	@FindBy(xpath = "//button[@title='Refresh']")
	WebElement refreshButton;
	
	@FindBy(xpath = "//span[contains(text(),'Qualify')]")
	WebElement qualifLead;

	@FindBy(xpath = "//span[@data-id='entity_name_span']")
	WebElement entityNameOpp;
	
	@FindBy(xpath = "(//div[@title='Propose'])[1]")
	WebElement proposeBPF;
	
	@FindBy(xpath = "//li[@title='Products']")
	WebElement productsTab;
	
	@FindBy(xpath = "//button[contains(@aria-label,'Search records for Price List, Lookup field')]")
	WebElement searchPriceListButton;
	
	@FindBy(xpath = "//li[@aria-label='SafeContractor Express Band F2-5716']")
	WebElement safeContractorBandF2;

	@FindBy(xpath = "//div[contains(@title,'PQQ')]")
	WebElement pqqText;

	@FindBy(xpath = "//span[normalize-space()='SafeContractor']")
	WebElement safeContractorText;
	
	@FindBy(xpath = "//button[@title='More commands for Opportunity']")
	WebElement moreCommandsOpp;

	@FindBy(xpath = "//img[contains(@title,'Flow')]")
	WebElement flowImgBtn;

	@FindBy(xpath = "//span[contains(text(),'Generate Quotation')]")
	WebElement generateQuotation;

	@FindBy(xpath = "//button[contains(@title,'OK')]")
	WebElement okButton;

	@FindBy(xpath = "//img[contains(@src, 'Save.c3e8b193ccb83c8b1f680b864b3a3109.svg')]")
	WebElement saveOpp;

	@FindBy(xpath = "//li[contains(@title,'Summary')]")
	WebElement summaryTab;

	@FindBy(xpath = "//div[contains(text(),'Quotation')]")
	WebElement quotationText;
	
	@FindBy(xpath = "//input[@data-id='safe_referralcode.fieldControl-text-box-text']")
	WebElement referralCodeText;

	@FindBy(xpath = "//input[@value='No']")
	WebElement noOptionInput;

	@FindBy(xpath = "(//div[@role='gridcell'])[2]")
	WebElement firstLeadCell;
	
	@FindBy(xpath = "//span[normalize-space()='Accept all']")
	WebElement acceptAll;
	
	@FindBy(xpath = "//div[@title='Account Profiles']")
	WebElement accountProfiles;
	
	@FindBy(xpath = "//button[@title='More commands for Account Profile']")
	WebElement moreCommandsforAP;

	@FindBy(xpath = "//span[contains(text(),'New Account Profile')]")
	WebElement newAccountProfile;
	
	@FindBy(xpath = "//input[@aria-label='Brand, Lookup']")
	private WebElement brandValue;
	
	@FindBy(xpath = "//ul[@aria-label='Lookup results']//li[1]")
	private WebElement SCCClientResult;
	
	@FindBy(xpath = "//input[@aria-label='Profile, Lookup']")
	WebElement profileLookup;
	
	@FindBy(xpath = "//li[@aria-label='Client']")
	WebElement clientOption;

	@FindBy(xpath = "(//button[@aria-label='Copilot'])[2]")
	WebElement secondCopilotButton;
	
	@FindBy(xpath = "//button[@title='Add an attachment']")
	WebElement attachIconSpan;

	@FindBy(xpath = "//h3[@aria-label='Timeline']")
	WebElement timelineHeader;
	
	@FindBy(xpath = "//span[text()='Add note and close']")
	WebElement addNoteAndclose;
	

	
	long currentTime = System.currentTimeMillis() / 10000;
	//String campaignName = "campaign " + currentTime;
	String accountName = "Test Account " + currentTime;
	String regNumber = Long.toString(currentTime % 100000000);
	String testURL = "www"+ currentTime + ".com";
	String contactsFirstName = "contactFirst " + currentTime;
	String contactsLastName = "contactLast " + currentTime;
	String phone = "+44" + currentTime;
	String contactEmail = "veekshan.poshala+" + currentTime + "@alcumus.com";
	String LeadFirstName = "First " + currentTime;
	String leadEmail = "Lead"+ currentTime + "@alcumus.com";
	String LeadLastName = "Last " + currentTime;
	
	Date currentDate = new Date();
    SimpleDateFormat dateFormat = new SimpleDateFormat("yyMMM dd HHmm");

    String formattedDate = dateFormat.format(currentDate);
       
	String CampaignName = "Campaign "+formattedDate;
	
	private static String SITE_WHERE_SERVICE = "All sites";
	private static String SITE_MANAGER_NAME = "site manager";
	//private static String PURCHASE_ORDER_DATE = "12/2/2024";
	private static String OWNNING_USER = "veekshan.poshala@alcumus.com";
	private static String OWNER_TYPE_ID = "systemuser";
	private static String PRE_CLEANSE_MATCH_RESULT = "N/A";
	
	public void createAccountAndContact() throws InterruptedException
	{
		se.clickElement(Sales);
		se.clickElement(Accounts);
		se.clickElement(New);
		se.type(accountFiled, "Test account " + currentTime);
		se.clickElement(secondCopilotButton);
		se.clickElement(acceptAll);
		se.type(turnOverInput, "22334");
		se.type(organizationTypeInputField, "Limited Company");
		se.clickElement(limitedCompanyOption);
		se.type(companyRegistrationNumberInput, regNumber);
		se.type(companyRegistrationYearInput, "2001");
		se.type(numberOfSitesInput, "23");
		se.scrollElementIntoView(websiteInput);
		se.type(websiteInput, testURL);
		se.scrollElementIntoView(street1Input);
		se.type(street1Input, "Bond Street");
		se.type(postcodeInput, "BS1 BS2");
		se.scrollElementIntoView(street1Input);
		se.clickElement(save);
		se.waitForElementTextToChange(unsavedToSaved, "- Saved");
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
		Thread.sleep(3000);
		se.clickElement(accountProfiles);
		se.clickElement(moreCommandsforAP);
		se.clickElement(newAccountProfile);
		
		se.type(brandValue, "SCC Client");
		se.clickElement(SCCClientResult);
		
		se.type(profileLookup, "Client");
		se.clickElement(clientOption);
		
		
		
		se.clickElement(save);
		
	}
	
	public void createCampaignAndAddLeads() throws InterruptedException
	{
		se.clickElement(Campaigns);
		se.clickElement(New);
		Thread.sleep(5000);
		se.type(campaignNameInput, CampaignName);
		se.clickElement(campaignType);
		se.scrollElementIntoView(leadSourceOption);
		se.clickOptionIfExists("Client List", leadSourceOption);
		se.clickElement(clientLookupInput);
		se.type(clientLookupInput, "Test account " + currentTime); 
		Thread.sleep(5000);
		se.clickElement(lookupResultsList);
		se.clickElement(preNoteDateInput);
		se.clickElement(calendarDay);
		se.type(brandLookUpValue, "SafeContractor");
		selectSafeContractorBrandlookUpValue();
		se.clickElement(save);
		Thread.sleep(5000);
//		se.clickElement(leadsTab);
//		
//		se.clickElement(newLeadButton);
		
//		driver.get("https://alcumusmain.crm4.dynamics.com/main.aspx?appid=2a66b27d-4a5c-ed11-9562-0022489ae2e0&pagetype=entityrecord&etn=lead");
//		Thread.sleep(5000);
//		
		
//		se.type(inputTopic, "test data");
//		se.type(leadFirstName, LeadFirstName);
//		se.type(leadLastName, LeadLastName);
//		se.type(leadType, "New Business");
//		se.scrollElementIntoView(NewBusiness);
//		se.clickElement(NewBusiness);
//		se.type(leadJobTitleInput, "SDE");
//		se.type(businessPhoneInput, "+44" + currentTime);
//		se.type(leadEmailInput, "Lead"+ currentTime + "@alcumus.com");
//		
//		se.type(contactLookupInput, contactsFirstName + " " + contactsLastName); 
//		se.clickElement(lookupResultsList);
//		se.waitForElementTextToChange(unsavedToSaved, "- Saved");
//		Thread.sleep(5000);
//		se.clickElement(leadSource);
//		se.scrollElementIntoView(leadSourceOption);
//		se.clickOptionIfExists("Paid Display", leadSourceOption);
//		se.clickElement(save);
//		Thread.sleep(5000);
//		se.clickElement(goBackButton);
//		Thread.sleep(4000);
//	
//		for (int i = 0; i <= 4; i++) 
//		{
//			se.clickElement(newLeadButton);
//			se.type(inputTopic, "test data");
//			se.type(leadFirstName, LeadFirstName+i);
//			se.type(leadLastName, LeadLastName+i);
//			se.type(leadType, "New Business");
//			se.scrollElementIntoView(NewBusiness);
//			se.clickElement(NewBusiness);
//			se.type(leadJobTitleInput, "SDE");
//			se.type(businessPhoneInput, "+44" + currentTime+i);
//			se.type(leadEmailInput, "Lead"+ currentTime+i + "@alcumus.com");
//			
//			//se.type(contactLookupInput, "contactFirst 169531335 contactLast 169531335");
//			se.type(contactLookupInput, contactsFirstName + " " + contactsLastName); 
//			se.clickElement(lookupResultsList);
//			Thread.sleep(3000);
//			if(ignoreAndSaveButton.getText().equals("Ignore and save"))
//			{
//				ignoreAndSaveButton.click();
//			}
//			else 
//			{
//				System.out.println("The button text does not equal 'Ignore and save");
//			}
//			se.waitForElementTextToChange(unsavedToSaved, "- Saved");
//			Thread.sleep(5000);
//			se.clickElement(leadSource);
//			Thread.sleep(1000);
//			se.scrollElementIntoView(leadSourceOption);
//			se.clickOptionIfExists("Paid Display", leadSourceOption);
//			se.clickElement(save);
//			Thread.sleep(5000);
//			se.clickElement(goBackButton);
//			Thread.sleep(4000);
//		}
		
	}
	
	public void QualifyTheLead() throws InterruptedException
	{
		
		se.doubleClickElement(firstLeadCell);
	//	driver.get("https://alcumusmain.crm4.dynamics.com/main.aspx?appid=2a66b27d-4a5c-ed11-9562-0022489ae2e0&cmdbar=true&pagetype=entityrecord&etn=lead&id=87a28a2a-4edc-4dac-a9f6-1a6f89944722");
		
		Thread.sleep(8000);
		
		se.clickElement(preQualifyBPF);
		se.clickElement(nextStageButton);
		//se.clickElement(closeStage);
		Thread.sleep(3000);
		
		se.clickElement(refreshButton);
		Thread.sleep(5000);
		
		se.clickElement(qualifyBPF);
		Thread.sleep(3000);
		
		
		se.type(numberOfEmployees, "120");
		
		se.clickElement(qualifySubStageButton);
		Thread.sleep(2000);
		se.clickOptionIfExists("DM Contact Made", leadSourceOption);
		
		se.clickElement(save);
		se.clickElement(ignoreAndSaveButton);
		
		se.clickElement(qualifLead);
		
		Thread.sleep(21000);
		
		se.assertElementText(entityNameOpp, "Opportunity");
		se.highlightElement(entityNameOpp);
		
		se.clickElement(proposeBPF);
		Thread.sleep(1000);
		
		Assert.assertEquals(qualifySubStageButton.getText(), "Quoted Non-DM", "Element does not display 'Quoted Non-DM' text");
		se.highlightElement(qualifySubStageButton);
		
		
	}
	
	public void addProduct() throws InterruptedException
	{
		se.clickElement(productsTab);
		se.clickElement(searchPriceListButton);
		se.clickElement(safeContractorBandF2);
		
		Thread.sleep(3000);
		
		se.clickElement(saveOpp);
		
		Thread.sleep(3000);
		
		se.assertElementText(pqqText, "PQQ");
		
		se.highlightElement(pqqText);
		
		se.assertElementText(safeContractorText, "SafeContractor");
		
		
		se.highlightElement(safeContractorText);
		
		se.clickElement(moreCommandsOpp);
		se.clickElement(flowImgBtn);
		Thread.sleep(2000);
		se.clickElement(generateQuotation);
		
		
		se.clickElement(okButton);
		
		se.clickElement(summaryTab);
		Thread.sleep(5000);
		se.clickElement(refreshButton);
		Thread.sleep(2000);
		
		se.assertElementText(quotationText, "Quotation");
		se.highlightElement(quotationText);
		
		String code = referralCodeText.getAttribute("value");
		
		System.out.println(code);
		
	//	se.scrollElementIntoView(noOptionInput);
		
	//	se.type(leadJobTitleInput, code);
		
	}
	
	public void addDataInExcel() 
	{
		
		String folderPath = System.getProperty("user.dir") + "\\src\\main\\java\\testdata\\";
	    String filePath = folderPath + "Template.xlsx";
	    
	    String[] filesToKeep = {"Template.xlsx", "TestData.xlsx"};
	    deleteOtherFiles(folderPath, filesToKeep);
		  
		// String filePath = System.getProperty("user.dir") + "\\src\\main\\java\\testdata\\Template.xlsx";
		 
		 try 
		 {
	            // Open the Excel file
	            FileInputStream inputStream = new FileInputStream(new File(filePath));

	            // Create a workbook instance
	            Workbook workbook = WorkbookFactory.create(inputStream);

	            // Get the sheet by its name
	            Sheet sheet = workbook.getSheet("Contractor List Template");
	            
	            System.out.println("Sheet name: " + sheet.getSheetName());
	            
	            // Get the first sheet by index
	             sheet = workbook.getSheetAt(0);

	             // Start clearing existing data from the fifth row (index 4)
	             for (int rowIndex = 4; rowIndex <= sheet.getLastRowNum(); rowIndex++) 
	             {
	                 Row row = sheet.getRow(rowIndex);
	                 if (row != null) 
	                 {
	                     // Clear cell values in each row
	                     for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) 
	                     {
	                         Cell cell = row.getCell(cellIndex);
	                         if (cell != null) {
	                             cell.setCellValue(""); // Clear cell value
	                         }
	                     }
	                 }
	             }


	             // Define row count
	             int rowCount = 5;
	             
	             // Write company names under "Company Name" header
	             writeCompanyNames(sheet, rowCount);

	             // Fill addresses 
	             fillAddresses(sheet, rowCount);
	             
	          // Fill telephone numbers in the Excel sheet
	             fillTelephoneNumbers(sheet, rowCount);
	             
	             
	          // Fill primary contacts in the Excel sheet
	             fillPrimaryContacts(sheet, rowCount);
	             
	             fillPositions(sheet, rowCount);
	             
	             fillEmails(sheet, rowCount);
	             
	          // Fills Site where service is provided (if applicable),  site manager/director contact name,  in the Excel sheet
	             // Last Purchase Order Date
	             allSitesAndSiteManagerAndLastPurchaseDate(sheet, rowCount);
	             
	           //  fillSiteManager(sheet, rowCount);

	             // Fill last purchase order date in the Excel sheet
	          //   fillPurchaseOrderDate(sheet, rowCount);

	             // Close the workbook and input stream
	             inputStream.close();
	             

	          // Create a new file for output
	             String newFilePath = System.getProperty("user.dir") + "\\src\\main\\java\\testdata\\" + CampaignName + ".xlsx";
	             FileOutputStream outputStream = new FileOutputStream(new File(newFilePath));
	             
	          // Write the updated workbook to the new file
	             workbook.write(outputStream);
	             workbook.close();
	             outputStream.close();

	             System.out.println("Data written to Excel file successfully.");
	             
	             
	        } 
		 catch (IOException e) 
		 {
	            System.out.println("Error reading Excel file: " + e.getMessage());
	            e.printStackTrace();
	     }
	    
	}
	
	public void deleteOtherFiles(String folderPath, String[] fileNamesToKeep) 
	{
	    File folder = new File(folderPath);
	    File[] files = folder.listFiles();
	    if (files != null) {
	        for (File file : files) {
	            if (file.isFile() && file.getName().endsWith(".xlsx") && !Arrays.asList(fileNamesToKeep).contains(file.getName())) {
	                try {
	                    Files.deleteIfExists(file.toPath());
	                    System.out.println("Deleted file: " + file.getName());
	                } catch (IOException e) {
	                    System.out.println("Error deleting file: " + file.getName());
	                    e.printStackTrace();
	                }
	            }
	        }
	    }
	}
	
	public void fileUpload() throws InterruptedException
		{
			 try {
			//driver.get("https://alcumusmain.crm4.dynamics.com/main.aspx?appid=2a66b27d-4a5c-ed11-9562-0022489ae2e0&pagetype=entityrecord&etn=campaign&id=7671deff-87ea-ee11-a203-000d3a69bbf5&cmdbar=true");
		//	Thread.sleep(11000);
			se.clickElement(secondCopilotButton);
			Thread.sleep(1000);
		//	se.clickElement(uploadFileElement);
		//	Thread.sleep(2000);
	        
	        // Specify the relative path to the test document
	        String filePath = System.getProperty("user.dir") + "\\src\\main\\java\\testdata\\Template.xlsx";
	        attachIconSpan.click();
	     //   se.clickElement(okButton);
	        Robot robot = new Robot();
	        robot.delay(2000);
	        
	        StringSelection stringSelection = new StringSelection(filePath);
	        Toolkit.getDefaultToolkit().getSystemClipboard().setContents(stringSelection, null);
	        
	        robot.keyPress(KeyEvent.VK_CONTROL);
	        robot.keyPress(KeyEvent.VK_V);
	        robot.keyRelease(KeyEvent.VK_V);
	        robot.keyRelease(KeyEvent.VK_CONTROL);
	
	        // Press Enter to navigate to the file
	        robot.keyPress(KeyEvent.VK_ENTER);
	        robot.keyRelease(KeyEvent.VK_ENTER);
	
	        // Wait for file explorer to load
	        Thread.sleep(2000);
	
	        // Press Enter to select the file
	//        robot.keyPress(KeyEvent.VK_ENTER);
	//        robot.keyRelease(KeyEvent.VK_ENTER);
	        
	//        robot.keyPress(KeyEvent.VK_ALT);
	//        robot.keyPress(KeyEvent.VK_F4);
	//        robot.keyRelease(KeyEvent.VK_F4);
	//        robot.keyRelease(KeyEvent.VK_ALT);
	        Thread.sleep(5000);
	        
	        se.clickElement(addNoteAndclose);
	        
	        Thread.sleep(10000);
			 }
			 catch (AWTException | InterruptedException e)
			 {
		            e.printStackTrace();
			 }
	        Thread.sleep(10000);
		}
	
	public void selectSafeContractorBrandlookUpValue()
	{
		try 
		{
			safeContractorOption.click();
		} 
		catch (NoSuchElementException e)
		{
			safeContractorUatOption.click();
		}
	}
	
	private static void writeCompanyNames(Sheet sheet, int rowCount) 
	{

        String headerName = "Company Name"; // Header name to match
        
        int companyNameIndex = findColumnIndex(sheet, headerName);
        if (headerName.equals("Company Name")) 
        { // Header found
            // Define company names array
			String[] companyNames = { "Amazon", "Google", "HCL", "Microsoft", "Apple", "Samsung", "Facebook", "IBM",
					"Oracle", "Intel", "Adobe", "Cisco", "Nvidia", "HP", "Dell", "Accenture", "Salesforce", "Uber",
					"Netflix", "Twitter", "LinkedIn", "Tesla", "PayPal", "Sony", "Airbnb", "SpaceX", "Reddit",
					"Dropbox", "Square", "Zoom", "Alibaba", "Tencent", "Baidu", "JD.com", "Huawei", "BMW",
					"Mercedes-Benz", "Toyota", "Volkswagen", "Ford", "Nike", "Adidas", "Puma", "Under Armour", "Gucci",
					"Louis Vuitton", "Prada", "Chanel", "Rolex", "Cartier", "Coca-Cola", "PepsiCo", "Nestle",
					"McDonald's", "Starbucks", "Walmart", "Target", "Costco", "General Electric", "Siemens", "Boeing",
					"Lockheed Martin", "Raytheon", "Northrop Grumman", "Booz Allen Hamilton", "General Dynamics",
					"BAE Systems", "ExxonMobil", "Chevron", "Shell", "BP", "Total", "Glencore", "Cargill",
					"Archer Daniels Midland", "Mitsubishi", "Toyota Tsusho", "Samsung Electronics", "LG Electronics",
					"Sony", "Panasonic", "Toshiba", "Nintendo", "Canon", "Nikon", "Sony Pictures", "Warner Bros.",
					"Paramount Pictures", "Universal Pictures", "20th Century Fox", "Lionsgate", "MGM", "Netflix",
					"HBO", "Hulu", "Disney", "Amazon Studios", "AT&T", "Verizon", "Comcast", "Charter Communications",
					"Vodafone", "China Mobile", "Nippon Telegraph and Telephone", "Telefonica", "Deutsche Telekom",
					"Orange", "Alibaba", "Tencent", "Baidu", "JD.com", "Meituan Dianping", "Infosys", "Wipro",
					"Tata Consultancy Services", "HCL Technologies", "Tech Mahindra", "Accenture", "Cognizant",
					"Capgemini", "DXC Technology", "IBM", "Visa", "Mastercard", "American Express", "PayPal",
					"Discover Financial Services", "JPMorgan Chase", "Bank of America", "Wells Fargo", "Citigroup",
					"Goldman Sachs", "Morgan Stanley", "UBS", "Credit Suisse", "HSBC", "Barclays",
					"Amazon Web Services", "Microsoft Azure", "Google Cloud Platform", "IBM Cloud", "Alibaba Cloud",
					"Salesforce", "Oracle Cloud", "SAP Cloud Platform", "VMware", "Red Hat", "Airbus", "Boeing",
					"Lockheed Martin", "Raytheon", "Northrop Grumman", "General Dynamics", "BAE Systems",
					"Rolls-Royce Holdings", "Safran", "GE Aviation", "ExxonMobil", "Shell", "BP", "Chevron", "Total",
					"Saudi Aramco", "Gazprom", "Rosneft", "ExxonMobil", "Royal Dutch Shell", "Toyota", "Volkswagen",
					"General Motors", "Ford", "Honda", "BMW", "Nissan", "Mercedes-Benz", "Hyundai", "Tesla", "Apple",
					"Samsung Electronics", "Sony", "LG Electronics", "Panasonic", "Toshiba", "Hitachi", "Nokia",
					"Xiaomi", "Huawei", "eBay", "Walmart", "Flipkart", "Rakuten", "Zalando", "Otto Group", "LVMH",
					"H&M", "Nike", "Adidas", "Puma", "Under Armour", "Uniqlo", "Zara", "Gap", "Forever 21", "Burberry",
					"Louis Vuitton", "Gucci", "Prada", "Rolex", "Cartier", "Pinterest", "Snapchat", "LinkedIn",
					"Dropbox", "Slack", "Spotify", "Etsy", "Reddit", "Twitch", "Medium", "Vimeo", "Quora",
					"Stack Overflow", "Discord", "GitLab", "TikTok", "Zoom", "PayPal", "Venmo", "Stripe", "Square",
					"Robinhood", "Coinbase", "Acorns", "WeWork", "SpaceX", "Blue Origin", "Virgin Galactic",
					"Rocket Lab" };


            Random random = new Random();
            ArrayList<String> remainingCompanyNames = new ArrayList<>(Arrays.asList(companyNames));
            for (int i = 0; i < rowCount; i++) {
                if (remainingCompanyNames.isEmpty()) {
                    break; // No more unique company names left
                }
                int randomIndex = random.nextInt(remainingCompanyNames.size());
                String companyName = remainingCompanyNames.remove(randomIndex);

                // Generate a unique 6-character string
                String uniqueString = generateUniqueString();

                // Append the unique string to the company name
                String companyNameWithUniqueString = companyName + " " + uniqueString;

                Row row = sheet.createRow(i + 4); // Start from the fifth row
                Cell cell = row.createCell(companyNameIndex); // Use the index of the "Company Name" header
                cell.setCellValue(companyNameWithUniqueString);
            }
        } else {
            System.out.println("Header '" + headerName + "' not found.");
        }
    }
	
	private static String generateUniqueString() 
	{
        StringBuilder uniqueString = new StringBuilder();
        Random random = new Random();
        // Generate 6 random characters
        for (int i = 0; i < 6; i++) 
        {
            // Generate a random ASCII value between 97 and 122 (inclusive) for lowercase letters
            char c = (char) (random.nextInt(26) + 97);
            uniqueString.append(c);
        }
        return uniqueString.toString();
    }
	
	private static void fillAddresses(Sheet sheet, int rowCount) 
	{
		String[] address1 = { "1 Main Street", "2 High Street", "3 Elm Avenue", "4 Oak Road", "5 Maple Close",
				"6 Pine Road", "7 Birch Lane", "8 Cedar Avenue", "9 Willow Crescent", "10 Chestnut Drive",
				"11 Rose Terrace", "12 Elm Street", "13 Willow Walk", "14 Cherry Lane", "15 Cedar Close",
				"16 Sycamore Way", "17 Beech Grove", "18 Hawthorn Lane", "19 Ivy Lane", "20 Acorn Avenue",
				"21 Pinecrest Drive", "22 Oakwood Avenue", "23 Maple Leaf Close", "24 Birchwood Lane",
				"25 Willowbrook Road", "26 Elmwood Terrace", "27 Cedar Ridge", "28 Chestnut Grove", "29 Holly Court",
				"30 Aspen Drive", "31 Hazelwood Way", "32 Rowan Street", "33 Juniper Lane", "34 Elderberry Close",
				"35 Mulberry Lane", "36 Poplar Street", "37 Alder Avenue", "38 Laurel Lane", "39 Redwood Avenue",
				"40 Walnut Close", "41 Magnolia Drive", "42 Jasmine Terrace", "43 Orchid Way", "44 Daisy Court",
				"45 Tulip Lane", "46 Lily Lane", "47 Violet Street", "48 Rosewood Avenue", "49 Heather Court",
				"50 Bluebell Lane", "51 Primrose Drive", "52 Sunflower Street", "53 Daffodil Court", "54 Marigold Lane",
				"55 Iris Way", "56 Lavender Avenue", "57 Carnation Close", "58 Peony Lane", "59 Hyacinth Street",
				"60 Pansy Drive", "61 Petunia Terrace", "62 Dahlia Lane", "63 Aster Way", "64 Orchid Court",
				"65 Lily of the Valley Road", "66 Snowdrop Street", "67 Daisy Avenue", "68 Poppy Lane",
				"69 Buttercup Drive", "70 Foxglove Terrace", "71 Clover Court", "72 Forget-Me-Not Lane",
				"73 Honeysuckle Avenue", "74 Hollyhock Street", "75 Thistle Lane", "76 Crocus Way",
				"77 Snapdragon Court", "78 Sweet Pea Terrace", "79 Bluebell Avenue", "80 Columbine Lane",
				"81 Canterbury Road", "82 Oxford Street", "83 Cambridge Lane", "84 Brighton Avenue",
				"85 Manchester Road", "86 Glasgow Street", "87 Cardiff Lane", "88 Belfast Avenue", "89 York Street",
				"90 Newcastle Road", "91 Edinburgh Lane", "92 Bristol Road", "93 Liverpool Street", "94 Sheffield Lane",
				"95 Leeds Road", "96 Birmingham Street", "97 Manchester Lane", "98 London Road", "99 Glasgow Lane",
				"100 Cardiff Road" };
		String[] address2 = { "Apartment 1", "Flat 2A", "Unit B", "Suite 3", "Room 4", "Flat 5B", "Unit C",
				"Apartment 6", "Suite 7A", "Room 8", "Flat 9C", "Unit 10A", "Apartment 11B", "Suite 12", "Room 13",
				"Flat 14D", "Unit E", "Apartment 15B", "Suite 16C", "Room 17A", "Flat 18B", "Unit 19A", "Apartment 20C",
				"Suite 21D", "Room 22A", "Flat 23A", "Unit 24B", "Apartment 25A", "Suite 26B", "Room 27C", "Flat 28D",
				"Unit 29C", "Apartment 30B", "Suite 31A", "Room 32D", "Flat 33A", "Unit 34B", "Apartment 35C",
				"Suite 36A", "Room 37D", "Flat 38C", "Unit 39D", "Apartment 40B", "Suite 41A", "Room 42C", "Flat 43A",
				"Unit 44B", "Apartment 45D", "Suite 46B", "Room 47A", "Flat 48D", "Unit 49A", "Apartment 50C",
				"Suite 51B", "Room 52A", "Flat 53C", "Unit 54A", "Apartment 55B", "Suite 56C", "Room 57A", "Flat 58B",
				"Unit 59C", "Apartment 60D", "Suite 61A", "Room 62C", "Flat 63A", "Unit 64B", "Apartment 65D",
				"Suite 66C", "Room 67A", "Flat 68D", "Unit 69C", "Apartment 70B", "Suite 71A", "Room 72C", "Flat 73A",
				"Unit 74B", "Apartment 75D", "Suite 76C", "Room 77A", "Flat 78D", "Unit 79C", "Apartment 80B",
				"Suite 81A", "Room 82D", "Flat 83A", "Unit 84B", "Apartment 85D", "Suite 86C", "Room 87A", "Flat 88D",
				"Unit 89C", "Apartment 90B", "Suite 91A", "Room 92D", "Flat 93A", "Unit 94B", "Apartment 95D",
				"Suite 96C", "Room 97A", "Flat 98D", "Unit 99C", "Apartment 100B" };
		String[] address3 = { "", "Building C", "Floor 2", "", "Floor 3", "", "Building D", "Floor 1", "", "Floor 4",
				"Building E", "Floor 3", "", "Building F", "Floor 2", "", "Building G", "Floor 1", "", "Floor 5",
				"Building H", "Floor 4", "", "Building I", "Floor 3", "", "Building J", "Floor 2", "", "Floor 6",
				"Building K", "Floor 5", "", "Building L", "Floor 4", "", "Building M", "Floor 3", "", "Floor 7",
				"Building N", "Floor 6", "", "Building O", "Floor 5", "", "Building P", "Floor 4", "", "Floor 8",
				"Building Q", "Floor 7", "", "Building R", "Floor 6", "", "Building S", "Floor 5", "", "Floor 9",
				"Building T", "Floor 8", "", "Building U", "Floor 7", "", "Building V", "Floor 6", "", "Floor 10",
				"Building W", "Floor 9", "", "Building X", "Floor 8", "", "Building Y", "Floor 7" };

		String[] town = { "London", "Manchester", "Birmingham", "Leeds", "Glasgow", "Liverpool", "Sheffield",
				"Edinburgh", "Bristol", "Cardiff", "Oxford", "Cambridge", "Brighton", "Newcastle upon Tyne", "Belfast",
				"Nottingham", "Reading", "Leicester", "Southampton", "Coventry", "York", "Hull", "Portsmouth",
				"Plymouth", "Norwich", "Swansea", "Aberdeen", "Dundee", "Blackpool", "Bournemouth", "Middlesbrough",
				"Luton", "Wolverhampton", "Sunderland", "Huddersfield", "Walsall", "Peterborough", "Slough",
				"Gloucester", "Camden", "Canterbury", "Croydon", "Ealing", "Guildford", "Harrow", "Havering",
				"Islington", "Kingston upon Thames", "Lambeth", "Lewisham", "Merton", "Redbridge",
				"Richmond upon Thames", "Sutton", "Waltham Forest", "Wandsworth", "Westminster", "Bath", "Durham",
				"Exeter", "Lincoln", "Hereford", "St Albans", "Stoke-on-Trent", "Worcester", "Armagh", "Lisburn",
				"Newry", "Newport", "Stirling", "Inverness", "Derry", "Londonderry", "Bangor", "Craigavon", "Cookstown",
				"Downpatrick", "Dungannon", "Enniskillen", "Larne", "Limavady", "Newtownabbey", "Newtownards", "Omagh",
				"Strabane", "Aberystwyth", "Bangor", "Cardigan", "Machynlleth", "Porthmadog", "Caernarfon", "Denbigh",
				"Llangefni", "Buckingham", "Basingstoke", "Halifax", "Stockport", "Woking", "Chelmsford", "Colchester",
				"Chichester", "Lancaster", "Preston", "Carmarthen", "Haverfordwest", "Brecon", "Bridgend", "Caerphilly",
				"Merthyr Tydfil", "Pontypridd", "Rhyl", "Swansea", "Llandudno", "Abergavenny", "Neath", "Pembroke",
				"Bala", "Barmouth", "Blaenau Ffestiniog", "Holyhead", "Llanidloes", "Machynlleth", "Rhayader",
				"Tregaron", "Aberaeron", "Aberystwyth", "Cardigan", "Lampeter", "Newcastle Emlyn", "Fishguard" };

		String[] county = { "Greater London", "Greater Manchester", "West Midlands", "West Yorkshire", "Lanarkshire",
				"Merseyside", "South Yorkshire", "Midlothian", "Bristol County", "South Glamorgan", "Oxfordshire",
				"Cambridgeshire", "East Sussex", "Tyne and Wear", "Antrim", "Nottinghamshire", "Berkshire",
				"Leicestershire", "Hampshire", "West Midlands", "North Yorkshire", "East Riding of Yorkshire",
				"Hampshire", "Devon", "Norfolk", "Swansea", "Aberdeenshire", "Dundee", "Lancashire", "Dorset",
				"North Yorkshire", "Bedfordshire", "West Midlands", "Tyne and Wear", "West Yorkshire", "West Midlands",
				"Cambridgeshire", "Berkshire", "Gloucestershire", "London", "Kent", "Greater London", "Greater London",
				"Surrey", "Greater London", "Greater London", "Greater London", "Greater London", "Greater London",
				"Greater London", "Greater London", "Greater London", "Greater London", "Greater London",
				"Greater London", "Greater London", "Greater London", "Somerset", "County Durham", "Devon",
				"Lincolnshire", "Herefordshire", "Hertfordshire", "Staffordshire", "Worcestershire", "County Armagh",
				"County Antrim", "County Down", "Newport", "Stirling", "Highland", "County Londonderry",
				"County Londonderry", "County Down", "County Armagh", "County Tyrone", "County Tyrone", "County Down",
				"County Tyrone", "County Fermanagh", "County Londonderry", "County Antrim", "County Down",
				"County Fermanagh", "County Tyrone", "Ceredigion", "Gwynedd", "Ceredigion", "Gwynedd", "Gwynedd",
				"Gwynedd", "Denbighshire", "Anglesey", "Buckinghamshire", "Hampshire", "West Yorkshire",
				"Greater Manchester", "Surrey", "Essex", "Essex", "West Sussex", "Lancashire", "Lancashire",
				"Carmarthenshire", "Pembrokeshire", "Powys", "Bridgend", "Caerphilly" };

		String[] postcode = { "W1A 1AA", "M1 1AA", "B1 1AA", "LS1 1AA", "G1 1AA", "L1 1AA", "S1 1AA", "EH1 1AA",
				"BS1 1AA", "CF1 1AA", "OX1 1AA", "CB1 1AA", "BN1 1AA", "NE1 1AA", "BT1 1AA", "NG1 1AA", "RG1 1AA",
				"LE1 1AA", "SO1 1AA", "CV1 1AA", "YO1 1AA", "HU1 1AA", "PO1 1AA", "PL1 1AA", "NR1 1AA", "SA1 1AA",
				"AB1 1AA", "DD1 1AA", "FY1 1AA", "BH1 1AA", "TS1 1AA", "LU1 1AA", "WV1 1AA", "SR1 1AA", "HD1 1AA",
				"WS1 1AA", "PE1 1AA", "SL1 1AA", "GL1 1AA", "NW1 1AA", "CT1 1AA", "CR0 1AA", "W5 1AA", "GU1 1AA",
				"HA1 1AA", "RM1 1AA", "N1 1AA", "KT1 1AA", "SE1 1AA", "SW1A 1AA", "SW17 0AA", "SW19 1AA", "TW1 1AA",
				"TW12 1AA", "TW19 1AA", "UB1 1AA", "W1G 9AA", "BA1 1AA", "DH1 1AA", "EX1 1AA", "LN1 1AA", "HR1 1AA",
				"AL1 1AA", "ST1 1AA", "WR1 1AA", "BT60 1AA", "BT27 1AA", "BT35 1AA", "NP10 1AA", "FK8 1AA", "IV1 1AA",
				"BT48 1AA", "BT47 1AA", "BT20 1AA", "BT62 1AA", "BT80 1AA", "BT30 1AA", "BT71 1AA", "BT1 2AA",
				"BT78 1AA", "SY23 1AA", "LL57 1AA", "SA43 1AA", "SY20 9AA", "LL49 9AA", "LL55 4AA", "LL16 3AA",
				"HP22 5AA", "PO9 1AA", "HX1 1AA", "SK1 1AA", "GU22 7AA", "CM1 1AA", "CO1 1AA", "PO19 1AA", "LA1 1AA",
				"PR1 1AA", "SA31 1AA", "SA61 1AA", "LD3 7AA", "CF31 1AA" };

	        int addressCount = Math.min(rowCount, address1.length); // Get the minimum of rowCount and address array length

	        Random random = new Random();
	        for (int i = 0; i < rowCount; i++) {
	            Row row = sheet.getRow(i + 4);
	            if (row == null) {
	                row = sheet.createRow(i + 4);
	            }

	            int randomIndex = random.nextInt(addressCount);

	            Cell address1Cell = row.createCell(findColumnIndex(sheet, "Address Line 1"));
	            Cell address2Cell = row.createCell(findColumnIndex(sheet, "Address Line 2"));
	            Cell address3Cell = row.createCell(findColumnIndex(sheet, "Address Line 3"));
	            Cell townCell = row.createCell(findColumnIndex(sheet, "Town"));
	            Cell countyCell = row.createCell(findColumnIndex(sheet, "County"));
	            Cell postcodeCell = row.createCell(findColumnIndex(sheet, "Post Code"));

	            address1Cell.setCellValue(address1[randomIndex]);
	            address2Cell.setCellValue(address2[randomIndex]);
	            address3Cell.setCellValue(address3[randomIndex]);
	            townCell.setCellValue(town[randomIndex]);
	            countyCell.setCellValue(county[randomIndex]);
	            postcodeCell.setCellValue(postcode[randomIndex]);
	        }
	    }
	
	private static void fillTelephoneNumbers(Sheet sheet, int rowCount) 
	{
	        Random random = new Random();

	        for (int i = 0; i < rowCount; i++) {
	            Row row = sheet.getRow(i + 4);
	            if (row == null) {
	                row = sheet.createRow(i + 4);
	            }

	            Cell telCell = row.createCell(7); // Assuming "Tel" is the eighth column
	            Cell altTelCell = row.createCell(8); // Assuming "Alternative Tel" is the ninth column

	            // Generate random telephone numbers within the specified range for "Tel"
	            long minRangeTel = 2200000000L;
	            long maxRangeTel = 9999999999L;
	            long telNumber = minRangeTel + (long) (random.nextDouble() * (maxRangeTel - minRangeTel + 1));
	            String telNumberString = "+44 " + telNumber;

	            // Generate random telephone numbers within the specified range for "Alternative Tel"
	            long minRangeAltTel = 2200000000L;
	            long maxRangeAltTel = 9999999999L;
	            long altTelNumber;
	            do {
	                altTelNumber = minRangeAltTel + (long) (random.nextDouble() * (maxRangeAltTel - minRangeAltTel + 1));
	            } while (altTelNumber == telNumber); // Ensure that altTelNumber is different from telNumber
	            String altTelNumberString = "+44 " + altTelNumber;

	            // Write telephone numbers into cells
	            telCell.setCellValue(telNumberString);
	            altTelCell.setCellValue(altTelNumberString);
	        }
	    }
	   // Method to fill primary contacts in the Excel sheet
    private static void fillPrimaryContacts(Sheet sheet, int rowCount) {
		String[] firstNames = { "John", "Emma", "Michael", "Olivia", "William", "Sophia", "James", "Isabella",
				"Alexander", "Ava", "Ethan", "Charlotte", "Daniel", "Mia", "Benjamin", "Emily", "Jacob", "Abigail",
				"David", "Harper", "Elijah", "Amelia", "Matthew", "Evelyn", "Jackson", "Sofia", "Samuel", "Madison",
				"Ella", "Logan", "Avery", "Joseph", "Scarlett", "Levi", "Grace", "Gabriel", "Chloe", "Owen", "Victoria",
				"Carter", "Riley", "Lincoln", "Lily", "Wyatt", "Nora", "Jayden", "Hannah", "Dylan", "Zoe", "Luke",
				"Penelope", "Jack", "Lillian", "Isaac", "Addison", "Ryan", "Ellie", "Nathan", "Natalie", "Caleb",
				"Aubrey", "Henry", "Stella", "Muhammad", "Mila", "Andrew", "Leah", "Mason", "Willow", "Lucas",
				"Samantha", "Joshua", "Aria", "Christopher", "Audrey", "Eli", "Claire", "Isaiah", "Bella", "Oliver",
				"Skylar", "Connor", "Lucy", "Christian", "Anna", "Hunter", "Maya", "Landon", "Gianna", "Adrian",
				"Valentina", "Evan", "Alexa", "Xavier", "Hailey", "Tyler", "Nora", "Tristan", "Paige", "Julian",
				"Layla", "Brayden", "Peyton", "Savannah", "Ian", "Kylie", "Max", "Natalia", "Leo", "Hazel", "Zachary",
				"Sarah", "Christopher", "Emma", "Nicholas", "Grace", "Jonathan", "Ella", "Justin", "Avery", "Jeremy",
				"Evelyn", "Kevin", "Sophie", "Brandon", "Aria", "Jacob", "Chloe", "Bryan", "Madison", "Jared", "Lily",
				"Eric", "Zoe", "Austin", "Charlotte", "Dylan", "Aubrey", "Scott", "Mila", "Derek", "Harper", "Jesse",
				"Aaliyah", "Philip", "Hannah", "Adrian", "Nora", "Keith", "Audrey", "Dustin", "Makayla", "Gregory",
				"Penelope", "Blake", "Maya", "Ian", "Scarlett", "Nathan", "Brooklyn", "Derrick", "Stella", "Raymond",
				"Amelia", "Travis", "Ellie", "Ronald", "Valeria", "Dwayne", "Leah", "Mitchell", "Morgan", "Gregory",
				"Abigail", "Colin", "Samantha", "Clinton", "Luna", "Joel", "Jasmine", "Preston", "Julia", "Brendan",
				"Natalie", "Keith", "Lillian", "Dennis", "Alexandra" };

		String[] lastNames = { "Smith", "Johnson", "Brown", "Williams", "Jones", "Miller", "Davis", "Garcia",
				"Rodriguez", "Martinez", "Wilson", "Taylor", "Anderson", "Thomas", "White", "Harris", "Martin",
				"Thompson", "Robinson", "Clark", "Lewis", "Lee", "Walker", "Hall", "Allen", "Young", "Hernandez",
				"King", "Wright", "Lopez", "Hill", "Scott", "Green", "Adams", "Baker", "Gonzalez", "Nelson", "Carter",
				"Mitchell", "Perez", "Roberts", "Turner", "Phillips", "Campbell", "Parker", "Evans", "Edwards",
				"Collins", "Stewart", "Sanchez", "Morris", "Rogers", "Reed", "Cook", "Morgan", "Bell", "Murphy",
				"Bailey", "Rivera", "Cooper", "Richardson", "Cox", "Howard", "Ward", "Torres", "Peterson", "Gray",
				"Ramirez", "James", "Watson", "Brooks", "Sanders", "Price", "Bennett", "Wood", "Barnes", "Ross",
				"Henderson", "Coleman", "Jenkins", "Perry", "Powell", "Long", "Patterson", "Hughes", "Flores",
				"Washington", "Butler", "Simmons", "Foster", "Gonzales", "Bryant", "Reyes", "Russell", "Griffin",
				"Diaz", "Hayes", "Myers", "Ford", "Hudson", "Hamilton", "Sullivan", "Ward", "Perkins", "Wheeler",
				"Wagner", "Burns", "Berry", "Gordon", "Bates", "Chapman", "Fletcher", "Holt", "Simmons", "Rossi",
				"Page", "Gibson", "Woods", "Holmes", "Reid", "Warren", "Ferguson", "Perkins", "Banks", "Chavez",
				"Conner", "Walters", "Mccoy", "Ellis", "Gardner", "Dixon", "Glover", "Garza", "Hubbard", "Palmer",
				"Clarke", "Vargas", "Hansen", "Nguyen", "Weber", "Sims", "Lowery", "Floyd", "Haynes", "Walters",
				"Benson", "Hale", "Luna", "Mccarthy", "Dickson", "Hicks", "Cameron", "Dunn", "Zimmerman", "Mendoza",
				"Potter", "Kim", "Haynes", "Munoz", "Joyce", "Baker", "Trujillo", "Bennett", "Colon", "Reid", "Roth",
				"Huffman", "Golden", "Yoder", "Giles", "Hood", "Vega", "Montoya", "Werner", "Carr", "Gould", "Mcdonald",
				"Mays", "Rangel", "Carrillo", "Ponce", "Nava", "Koch" };

        Random random = new Random();
        Set<String> generatedNames = new HashSet<>(); 
        for (int i = 0; i < rowCount; i++) {
            Row row = sheet.getRow(i + 4);
            if (row == null) {
                row = sheet.createRow(i + 4);
            }

            Cell primaryContactCell = row.createCell(9); // Assuming "Primary contact" is the tenth column
            // Generate a unique random full name using first names and last names arrays
            String fullName;
            do {
                fullName = generateFullName(firstNames[random.nextInt(firstNames.length)], lastNames[random.nextInt(lastNames.length)]);
            } while (!generatedNames.add(fullName)); // Ensure generated name is unique

            // Write primary contact name into cell
            primaryContactCell.setCellValue(fullName);
        }
    }

    // Method to generate a full name using first name and last name
    private static String generateFullName(String firstName, String lastName) {
        return firstName + " " + lastName;
    }
    
 // Method to fill positions in the Excel sheet
    private static void fillPositions(Sheet sheet, int rowCount) 
    {
        // Array of positions
        String[] positions = {"CEO", "CTO", "CFO", "COO", "Managing Director",
                "Senior Manager", "Project Manager", "Software Engineer", "Marketing Manager", "HR Director"};

        // Find column index for "Position" header
        int positionIndex = findColumnIndex(sheet, "Position");

        // Check if the header is found
        if (positionIndex != -1) 
        {
            Random random = new Random();

            for (int i = 0; i < rowCount; i++) {
                Row row = sheet.getRow(i + 4);
                if (row == null) {
                    row = sheet.createRow(i + 4);
                }

                Cell positionCell = row.createCell(positionIndex);

                // Get a random position from the positions array
                String position = positions[random.nextInt(positions.length)];

                // Write position into cell
                positionCell.setCellValue(position);
            }
        } else {
            System.out.println("Header 'Position' not found.");
        }
    }
    
  //  private static int findColumnIndex(Sheet sheet, String header) 
    private static int findColumnIndex(Sheet sheet, String header) 
    {
        Row headerRow = sheet.getRow(3); // Assuming header is in the fourth row
        if (headerRow != null) {
            for (Cell cell : headerRow) {
                if (cell.getStringCellValue().equals(header)) {
                    return cell.getColumnIndex();
                }
            }
        }
        return -1; // Return -1 if header is not found
    }
 // Method to fill e-mails in the Excel sheet
    private static void fillEmails(Sheet sheet, int rowCount) {
        // Find column index for "e-mail" header
        int emailIndex = findColumnIndex(sheet, "e-mail");

        // Check if the header is found
        if (emailIndex != -1) {
            Random random = new Random();

            for (int i = 0; i < rowCount; i++) {
                Row row = sheet.getRow(i + 4);
                if (row == null) {
                    row = sheet.createRow(i + 4);
                }

                Cell emailCell = row.createCell(emailIndex);

                // Generate a random string of 10-15 characters
                int length = 10 + random.nextInt(6); // Random length between 10 and 15
                String randomString = generateRandomString(length);

                // Append the random string to the domain
                String email = randomString + "@gdm1byzh.mailosaur.net";

                // Write e-mail into cell
                emailCell.setCellValue(email);
            }
        } else {
            System.out.println("Header 'e-mail' not found.");
        }
    }

   public void allSitesAndSiteManagerAndLastPurchaseDate(Sheet sheet, int rowCount) 
    {
    // Find column indices for additional headers
    int siteWhereServiceIndex = findColumnIndex(sheet, "Site where service is provided (if applicable)");
    int siteManagerNameIndex = findColumnIndex(sheet, "Site Manager/Director Contact Name");
   // int purchaseOrderDateIndex = findColumnIndex(sheet, "Last Purchase Order Date");
    int campaignNameIndex = findColumnIndex(sheet, "Campaign Name");
    int ownningUserIndex = findColumnIndex(sheet, "Ownning User");
    int ownerTypeIdIndex = findColumnIndex(sheet, "Ownertypeid");
    int preCleanseIndex = findColumnIndex(sheet, "Pre-Cleanse Match Result");

    // Check if the headers are found
   // if (siteWhereServiceIndex != -1 && siteManagerNameIndex != -1 && purchaseOrderDateIndex != -1) 
    if (siteWhereServiceIndex != -1 && siteManagerNameIndex != -1 ) 
    {
        for (int i = 0; i < rowCount; i++) 
        {
            Row row = sheet.getRow(i + 4);
            if (row == null) 
            {
                row = sheet.createRow(i + 4);
            }

            // Create cells for additional data
            Cell siteWhereServiceCell = row.createCell(siteWhereServiceIndex);
            Cell siteManagerNameCell = row.createCell(siteManagerNameIndex);
       //     Cell purchaseOrderDateCell = row.createCell(purchaseOrderDateIndex);
            Cell campaignNameCell = row.createCell(campaignNameIndex);
            Cell ownningUserNameCell = row.createCell(ownningUserIndex);
            Cell ownerTypeIdNameCell = row.createCell(ownerTypeIdIndex);
            Cell preCleanseCell = row.createCell(preCleanseIndex);

            // Write data into cells
            siteWhereServiceCell.setCellValue(SITE_WHERE_SERVICE);
            siteManagerNameCell.setCellValue(SITE_MANAGER_NAME);
      //      purchaseOrderDateCell.setCellValue(PURCHASE_ORDER_DATE);
            campaignNameCell.setCellValue(CampaignName);
            ownningUserNameCell.setCellValue(OWNNING_USER);
            ownerTypeIdNameCell.setCellValue(OWNER_TYPE_ID);
            preCleanseCell.setCellValue(PRE_CLEANSE_MATCH_RESULT);
        }
    } 
    else 
    {
        System.out.println("One or more headers for additional data not found.");
    }

    }
    // Method to generate a random string of given length
    private static String generateRandomString(int length) 
    {
        StringBuilder sb = new StringBuilder();
        Random random = new Random();
        String characters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
        for (int i = 0; i < length; i++) 
        {
            sb.append(characters.charAt(random.nextInt(characters.length())));
        }
        return sb.toString();
    }

}