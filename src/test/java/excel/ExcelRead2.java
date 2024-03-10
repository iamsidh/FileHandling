package excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead2 extends Base {

	public static void main(String[] args) throws Exception {

		FileInputStream fis = new FileInputStream(excelSheet());

		XSSFWorkbook wb = new XSSFWorkbook(fis);

		XSSFSheet sheet = wb.getSheet("Sheet1");

		//int rows = sheet.getLastRowNum();
		
		int rows = sheet.getPhysicalNumberOfRows();

		int cols = sheet.getRow(1).getLastCellNum();

		System.out.println("Number of Rows are :" + rows);
		System.out.println("Number of Columns are :" + cols);

		for (int r = 0; r < rows; r++) {

			XSSFRow row = sheet.getRow(r);

			for (int c = 0; c < cols; c++) {

				XSSFCell cell = row.getCell(c);

				switch (cell.getCellType()) {

				case STRING:
					System.out.print(cell.getStringCellValue());
					break;

				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;

				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;

				case FORMULA:
					System.out.print(cell.getNumericCellValue());
					break;
				default:
					System.out.println((String) cell.getStringCellValue());
					break;

				}

				System.out.print(" | ");

			}
			System.out.println();
		}

	}

}
