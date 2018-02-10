package pageObjects;

import java.util.List;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindAll;
import org.openqa.selenium.support.FindBy;

public class DownloadPage extends BasePage {

	@FindBy (xpath = "//a/span[text()='Excel']")
	public WebElement downloadBtn;

	@FindAll({
		@FindBy(xpath = "//table[@id='example']//tbody/tr")
	})
	public List<WebElement> rows;

	@FindAll({
		@FindBy(xpath = "//table[@id='example']/thead/tr/th")
	})
	public List<WebElement> columns;

	public void ClickDownload() {
		downloadBtn.click();
	}

	public List<WebElement> numberOfRows() {
		return rows;
	}

	public List<WebElement> numberOfColumns() {
		return columns;
	}

}
