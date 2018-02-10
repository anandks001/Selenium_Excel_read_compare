package tests;

import java.io.File;
import java.util.HashMap;
import java.util.concurrent.TimeUnit;


import org.apache.logging.log4j.LogManager;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.PageFactory;
import org.apache.logging.log4j.Logger;
import pageObjects.DownloadPage;
import util.PropertyReader;

public class BaseTests {

	static WebDriver driver;
	static String downloadPath = System.getProperty("user.dir") + File.separator + "externalFileDownloads";
	static String downloadedFile = System.getProperty("user.dir") +
			File.separator + "externalFileDownloads" + File.separator + "Selenium Easy - Download Table Data to CSV, Excel, PDF and Print.xlsx";
	//	static String downloadedFile;
	static String actualFile = System.getProperty("user.dir") + File.separator + "externalFileDownloads"
			+ File.separator + "Actual_data.xlsx";
	DownloadPage dp = PageFactory.initElements(driver, DownloadPage.class);
	static PropertyReader prop = new PropertyReader();
	static final Logger logger = LogManager.getLogger(BaseTests.class);



	@BeforeClass
	public static void setup() throws Exception {
		System.setProperty("webdriver.chrome.driver",prop.readProperty("ChromeDriverPath"));

		HashMap<String, Object> prefs = new HashMap<String, Object>();
		prefs.put("download.default_directory", downloadPath);
		ChromeOptions options = new ChromeOptions();
		options.setExperimentalOption("prefs", prefs);
		driver = new ChromeDriver(options);

		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);

	}

	@AfterClass
	public static void teardown() throws Exception {
		try {
			System.out.println("Teardown");
			File f = new File(downloadedFile);
			f.delete();
			driver.quit();
		} catch (Exception e) {
			System.out.println(e.toString());
			throw (e);
		}
	}
}
