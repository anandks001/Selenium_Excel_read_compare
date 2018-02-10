package appModules;

import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CompareExcel {

	@SuppressWarnings("resource")
	public static boolean ExcelComparison(String file1, String file2) throws FileNotFoundException {
		boolean equalWorkBook = false;
		try {
			XSSFWorkbook workbook1 = new XSSFWorkbook(file1);
			XSSFWorkbook workbook2 = new XSSFWorkbook(file2);
			XSSFSheet sheet1 = workbook1.getSheetAt(0);
			XSSFSheet sheet2 = workbook2.getSheetAt(0);
			if (compareTwoSheets(sheet1, sheet2)) {
				System.out.println("Excel Sheets contians equal data");
				equalWorkBook = true;
			} else {
				System.out.println("Excel sheets are Not Equal");
			}

		} catch (FileNotFoundException fnf) {
			fnf.printStackTrace();
			System.out.println(fnf.toString());
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println(e.toString());
		}

		return equalWorkBook;

	}

	public static boolean compareTwoSheets(XSSFSheet sheet1, XSSFSheet sheet2) {
		int firstRow1 = sheet1.getFirstRowNum();
		int lastRow1 = sheet1.getLastRowNum();
		boolean equalSheets = true;
		for (int i = firstRow1; i <= lastRow1; i++) {
			XSSFRow row1 = sheet1.getRow(i);
			XSSFRow row2 = sheet2.getRow(i);
			if (!compareTwoRows(row1, row2)) {
				equalSheets = false;
				System.out.println("Row " + i + " - Not Equal");
				break;
			} else {
				System.out.println("Row " + i + " - Equal");
			}
		}
		return equalSheets;
	}

	public static boolean compareTwoRows(XSSFRow row1, XSSFRow row2) {
		if ((row1 == null) && (row2 == null)) {
			return true;
		} else if ((row1 == null) || (row2 == null)) {
			return false;
		}

		int firstCell1 = row1.getFirstCellNum();
		int lastCell1 = row1.getLastCellNum();
		boolean equalRows = true;

		for (int i = firstCell1; i <= lastCell1; i++) {
			XSSFCell cell1 = row1.getCell(i);
			XSSFCell cell2 = row2.getCell(i);
			if (!compareTwoCells(cell1, cell2)) {
				equalRows = false;
				System.out.println("Cell " + i + " - NOt Equal");
				break;
			} else {
				System.out.println("Cell " + i + " - Equal");
			}
		}
		return equalRows;
	}

	public static boolean compareTwoCells(XSSFCell cell1, XSSFCell cell2) {
		boolean equalCells = false;
		if ((cell1 == null) && (cell2 == null)) {
			System.err.println("\n\nCells are empty\n\n");
			return true;
		} else if ((cell1 == null) || (cell2 == null)) {
			System.err.println("Anyone cell is empty");
			return false;
		} else {
			DataFormatter df = new DataFormatter();
			String cell1Data = df.formatCellValue(cell1);
			String cell2Data = df.formatCellValue(cell2);
			if (cell1Data.equals(cell2Data)) {
				System.out.println("Cell 1 " + cell1Data + " and " + cell2Data + " Equal" );
				return true;
			} else {
				System.out.println("Cell 1 " + cell1Data + " and " + cell2Data + " Not Equal" );
			}
		}
		return equalCells;
	}

}
