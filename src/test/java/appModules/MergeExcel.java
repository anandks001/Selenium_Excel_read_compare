package appModules;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MergeExcel {

	@SuppressWarnings("resource")
	public static boolean mergeExcelFiles(String file1, String file2) throws FileNotFoundException {
		boolean merge = false;
		try {
			XSSFWorkbook workbook1 = new XSSFWorkbook(file1);
			XSSFWorkbook workbook2 = new XSSFWorkbook(file2);
			XSSFSheet sheet1 = workbook1.getSheetAt(0);
			XSSFSheet sheet2 = workbook2.getSheetAt(0);
			addSheet(sheet1, sheet2);
			File mergedFile = new File(System.getProperty("user.dir") + File.separator + "externalFileDownloads"
					+ File.separator + "merged.xlsx");
			if (!mergedFile.exists()) {
				mergedFile.createNewFile();
			}
			FileOutputStream out = new FileOutputStream(mergedFile);
			workbook1.write(out);
			out.close();
			System.out.println("Files were merged succussfully");
			merge = true;

		} catch (FileNotFoundException fnf) {
			fnf.printStackTrace();
			System.out.println(fnf.toString());
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println(e.toString());
		}
		return merge;
	}

	@SuppressWarnings("deprecation")
	public static void addSheet(XSSFSheet mergedSheet, XSSFSheet sheet) {

		Map<Integer, XSSFCellStyle> styleMap = new HashMap<Integer, XSSFCellStyle>();
		int len = mergedSheet.getLastRowNum();
		for (int j = sheet.getFirstRowNum() + 1; j <= sheet.getLastRowNum(); j++) {

			XSSFRow row = sheet.getRow(j);
			XSSFRow mrow = mergedSheet.createRow(len + j + 1);

			for (int k = row.getFirstCellNum(); k < row.getLastCellNum(); k++) {
				XSSFCell cell = row.getCell(k);
				XSSFCell mcell = mrow.createCell(k);
				if (cell.getSheet().getWorkbook() == mcell.getSheet().getWorkbook()) {
					mcell.setCellStyle(cell.getCellStyle());
				} else {
					int stHashCode = cell.getCellStyle().hashCode();
					XSSFCellStyle newCellStyle = styleMap.get(stHashCode);
					if (newCellStyle == null) {
						newCellStyle = mcell.getSheet().getWorkbook().createCellStyle();
						newCellStyle.cloneStyleFrom(cell.getCellStyle());
						styleMap.put(stHashCode, newCellStyle);
					}
					mcell.setCellStyle(newCellStyle);
				}

				switch (cell.getCellType()) {
				case HSSFCell.CELL_TYPE_FORMULA:
					mcell.setCellFormula(cell.getCellFormula());
					break;
				case HSSFCell.CELL_TYPE_NUMERIC:
					mcell.setCellValue(cell.getNumericCellValue());
					break;
				case HSSFCell.CELL_TYPE_STRING:
					mcell.setCellValue(cell.getStringCellValue());
					break;
				case HSSFCell.CELL_TYPE_BLANK:
					mcell.setCellType(HSSFCell.CELL_TYPE_BLANK);
					break;
				case HSSFCell.CELL_TYPE_BOOLEAN:
					mcell.setCellValue(cell.getBooleanCellValue());
					break;
				case HSSFCell.CELL_TYPE_ERROR:
					mcell.setCellErrorValue(cell.getErrorCellValue());
					break;
				default:
					mcell.setCellValue(cell.getStringCellValue());
					break;
				}
			}
		}
	}
}
