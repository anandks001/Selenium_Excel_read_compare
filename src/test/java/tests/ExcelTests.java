package tests;

import org.junit.runners.MethodSorters;
import org.junit.*;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import appModules.CompareExcel;
import appModules.ExcelOperations;
import appModules.MergeExcel;
import org.apache.log4j.PropertyConfigurator;

@FixMethodOrder(MethodSorters.NAME_ASCENDING)
public class ExcelTests extends BaseTests{

	@Test
	public void testADownloadFile() throws Exception {
		try {
			driver.get(prop.readProperty("url"));
			logger.trace("Open application");

			Assert.assertTrue(dp.downloadBtn.isDisplayed());
			logger.trace("verification completed");

			//			dp.ClickDownload();
			//			Thread.sleep(3000);
			//			File excel = ExcelOperations.getLatestFile(downloadPath);
			//			File f = new File(downloadPath + File.separator + excel);
			//			downloadedFile = downloadPath + File.separator + f.getName();
			//			System.out.println(f.getName());
			//			Assert.assertEquals("Excel file downloaded from website", f.getName(),
			//					"Selenium Easy - Download Table Data to CSV, Excel, PDF and Print.xlsx");
		} catch (Exception e) {
			logger.error("Verification failed");
			e.printStackTrace();
			System.out.println(e.toString());
			throw (e);
		}
	}

	//	@Test
	public void testBCompareExcelAndWeb() throws Exception {
		try {
			int rowCount = dp.numberOfRows().size();
			int colCount = dp.numberOfColumns().size();
			for (int i = 1; i <= rowCount; i++) {
				for (int j = 1; j <= colCount; j++) {
					WebElement celldata = driver
							.findElement(By.xpath("//table[@id='example']//tbody/tr[" + i + "]/td[" + j + "]"));
					String data = celldata.getText();
					String cell = ExcelOperations.readExcel(downloadedFile, i, j - 1);
					Assert.assertEquals("Excel data not match with web table values", cell, data);
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
			System.out.println(e.toString());
			throw (e);
		}
	}

	//	@Test
	public void testCCompare2Excel() throws Exception {
		try {
			System.out.println(actualFile);
			System.out.println(downloadedFile);
			Assert.assertTrue("Excel files comparsion was not successful", CompareExcel.ExcelComparison(actualFile, downloadedFile));
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println(e.toString());
			throw (e);
		}

	}

	//	@Test
	public void testDmergeExcelFiles() throws Exception {
		try {
			Assert.assertTrue("Excel files are not merged", MergeExcel.mergeExcelFiles(actualFile, downloadedFile));
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println(e.toString());
			throw (e);
		}

	}



}
