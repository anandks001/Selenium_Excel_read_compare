package appModules;

import java.io.File;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.*;

public class ExcelOperations {

	// Get latest file downloaded in given path
	public static File getLatestFile(String dirPath) {
		File dir = new File(dirPath);
		File[] files = dir.listFiles();
		if (files == null || files.length == 0) {
			return null;
		}

		File lastModifiedFile = files[0];
		for (int i = 1; i < files.length; i++) {
			if (lastModifiedFile.lastModified() < files[i].lastModified()) {
				lastModifiedFile = files[i];
			}
		}
		return lastModifiedFile;

	}

	//Read excel cell and return string value
	public static String readExcel(String pathdir, int rowcounter, int colcounter) throws IOException {

		@SuppressWarnings("resource")
		XSSFWorkbook srcBook = new XSSFWorkbook(pathdir);
		XSSFSheet sourceSheet = srcBook.getSheetAt(0);
		XSSFRow sourceRow = sourceSheet.getRow(rowcounter);
		XSSFCell cell2 = sourceRow.getCell(colcounter);
		DataFormatter df = new DataFormatter();
		String cell1 = df.formatCellValue(cell2);
		return cell1;
	}

	//Delete a file from the directory
	public boolean isFileDeleted(String downloadPath, String fileName) {
		File dir = new File(downloadPath);
		File[] dirContents = dir.listFiles();

		for (int i = 0; i < dirContents.length; i++) {
			if (dirContents[i].getName().equals(fileName)) {
				dirContents[i].delete();
				return true;
			}
		}
		return false;
	}

}