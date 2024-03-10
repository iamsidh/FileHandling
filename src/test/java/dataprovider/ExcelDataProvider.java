package dataprovider;

import java.io.IOException;
import java.lang.reflect.Array;
import java.util.Arrays;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;

import excelUtils.Base;

public class ExcelDataProvider extends Base {
	@DataProvider(name="LoginData")
	public static String[][] getExcelData() {

		try {
			XSSFWorkbook workbook = new XSSFWorkbook(getExcelSheet());

			XSSFSheet sheet = workbook.getSheet("Sheet1");

			int numOfRows = sheet.getLastRowNum();
			int numOfCols = sheet.getRow(0).getLastCellNum();

			String data[][] = new String[numOfRows][numOfCols];

			for (int r = 0; r < numOfRows; r++) {

				for (int c = 0; c < numOfCols; c++) {

					DataFormatter df = new DataFormatter();

					data[r][c] = df.formatCellValue(sheet.getRow(r + 1).getCell(c));

				}
			}
			Arrays.toString(data);

			/*
			 * for (String[] strings : data) {
			 * 
			 * System.out.println(Arrays.toString(strings)); }
			 */
			workbook.close();
			return data;

		} catch (IOException e) {

			e.printStackTrace();

			return null;
		}

	}

	public static void main(String[] args) {

		getExcelData();
	}

}
