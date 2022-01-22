package code;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Launch {
	static String appUrl="https://www.google.co.in";
	static String chromedriverLoc=".\\src\\test\\chromedriver.exe";
	static String xlsLoc=".\\src\\test\\Datasheet.xls";

	static XlsHelper xls=new XlsHelper(xlsLoc);
	public static void main(String[] args) {
		System.out.println("Launching Application.");
		System.setProperty("webdriver.chrome.driver", chromedriverLoc);
		WebDriver driver=new ChromeDriver();
		driver.get(appUrl);
		System.out.println("Page Title: "+driver.getTitle());
	

		xls.getColValuesByColNum(0);
		xls.getColValuesByColNum(1);
		xls.getColValuesByColNum(2);
		xls.getColValuesByColNum(3);
		xls.getAllValuesColWise();

		
		driver.quit();
	}
}
