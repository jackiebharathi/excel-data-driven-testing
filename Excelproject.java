package excelproject;

import java.io.IOException;
import java.time.Duration;
import java.util.List;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import june1.ExcelUtil;

public class Excelproject {

	public static void main(String[] args) throws IOException, InterruptedException {
	WebDriver driver = new ChromeDriver();
	driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
	WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
	driver.get("https://www.cit.com/cit-bank/resources/calculators/certificate-of-deposit-calculator");
	driver.manage().window().maximize();
	 // System.out.println("1");
	String fpath="C:\\Users\\user\\eclipse-workspace\\sele\\src\\test\\java\\TestData\\Book1.xlsx";
	  //System.out.println("2");
	int rows= ExcelUtil.getRowCount(fpath,"Sheet1");
	
	try {
	    WebElement bannerClose = driver.findElement(By.id("onetrust-close-btn-container"));
	    if (bannerClose.isDisplayed()) {
	        bannerClose.click();
	        Thread.sleep(500);
	    }
	} catch (Exception ignored) {}
	
	  //System.out.println("hello");
	for(int i=1;i<=rows;i++) {
		//read data from excel
		
		String dep= ExcelUtil.getCellData(fpath, "Sheet1",i,0);
		String rate= ExcelUtil.getCellData(fpath, "Sheet1",i,1);
		String mon= ExcelUtil.getCellData(fpath, "Sheet1",i,2);
		String comp= ExcelUtil.getCellData(fpath, "Sheet1",i,3);
		String e_val= ExcelUtil.getCellData(fpath, "Sheet1",i,4);
		
		//pass data into web
		
		driver.findElement(By.xpath("//input[@id='mat-input-0']")).clear();
		driver.findElement(By.xpath("//input[@id='mat-input-0']")).sendKeys(dep);
		
		driver.findElement(By.xpath("//input[contains(@class, 'mat-input-element') and @formcontrolname='cdLength']")).clear();
		driver.findElement(By.xpath("//input[contains(@class, 'mat-input-element') and @formcontrolname='cdLength']")).sendKeys(mon);
		
		driver.findElement(By.xpath("//input[@formcontrolname='cdRate']")).clear();
		driver.findElement(By.xpath("//input[@formcontrolname='cdRate']")).sendKeys(rate);
		WebElement dropdown = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@id='mat-select-value-1']")));
		dropdown.click();
		//Thread.sleep(1000);
		
		List<WebElement> options=driver.findElements(By.xpath("//div[@role='listbox']//span"));
		for (WebElement w : options) {
		    if (w.getText().trim().equalsIgnoreCase(comp)) {
		        w.click();
		        break;
		    }
		}		
		WebElement submitBtn = wait.until(ExpectedConditions.elementToBeClickable(By.id("CIT-chart-submit")));
		submitBtn.click();
		
		//validation
		String a_val=driver.findElement(By.xpath("//span[@id='displayTotalValue']")).getText();
		a_val = a_val.replaceAll("[$,]", "").replaceAll("\\s+", "");
		e_val = e_val.replaceAll("[$,]", "").replaceAll("\\s+", "");
		if(Double.parseDouble(a_val)==Double.parseDouble(e_val)) {
			ExcelUtil.setCellData(fpath, "Sheet1", i, 5,"Passed");
			ExcelUtil.fillGreenColor(fpath, "Sheet1", i, 5);
		}else {
			ExcelUtil.setCellData(fpath, "Sheet1", i, 5,"Failed");
			ExcelUtil.fillRedColor(fpath, "Sheet1", i, 5);
		}
		
		
	}
	driver.quit();
	}

}
