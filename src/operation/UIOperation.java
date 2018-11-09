package operation;


import java.util.Properties;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

public class UIOperation {

	WebDriver driver;
	public UIOperation(WebDriver driver){
		this.driver = driver;
	}
	public void perform(String strKeyword, String strObjectType, String strIdentifier, String strTestData_1) throws Exception{
		
		System.out.println("");
		switch (strKeyword.toUpperCase()) {
		case "CLICK":
			//Perform click
			driver.findElement(this.getObject(strIdentifier,strObjectType)).click();
			break;
		case "SETTEXT":
			//Set text on control
			driver.findElement(this.getObject(strIdentifier,strObjectType)).sendKeys(strTestData_1);
			break;
			
		case "GOTOURL":
			//Get url of application
			driver.get(strTestData_1);
			break;
		case "GETTEXT":
			//Get text of an element
			driver.findElement(this.getObject(strIdentifier,strObjectType)).getText();
			break;

		default:
			break;
		}
	}
	
	/*
	 * Find element BY using object type and value
	 */
	
	private By getObject(String strIdentifier, String strObjectType) throws Exception{
		//Find by xpath
		if(strObjectType.equalsIgnoreCase("XPATH")){
			
			return By.xpath(strIdentifier);
		}
		//find by class
		else if(strObjectType.equalsIgnoreCase("CLASSNAME")){
			
			return By.className(strIdentifier);
			
		}
		//find by name
		else if(strObjectType.equalsIgnoreCase("NAME")){
			
			return By.name(strIdentifier);
			
		}
		//Find by css
		else if(strObjectType.equalsIgnoreCase("CSS")){
			
			return By.cssSelector(strIdentifier);
			
		}
		//find by link
		else if(strObjectType.equalsIgnoreCase("LINK")){
			
			return By.linkText(strIdentifier);
			
		}
		//find by partial link
		else if(strObjectType.equalsIgnoreCase("PARTIALLINK")){
			
			return By.partialLinkText(strIdentifier);
			
		}else
		{
			throw new Exception("Wrong object type");
		}
	}
}
