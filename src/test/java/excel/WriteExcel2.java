package excel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel2 {

	public static void main(String[] args) throws Exception {

		Object data[][] = { { "ID", "Username", "Password" }, 
				{ 1, "Siddhesh", "sidh" }, { 2, "Roy", "ghana" },
				{ 3, "batista", "wwe" }

		};

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("credentials");

		int rows = data.length;
		int cols = data[0].length;

		System.out.println("Numbers of rows to be written : " + rows);
		System.out.println("Number of Columns to be written: " + cols);

		for (int r = 0; r < rows; r++) {

			XSSFRow row = sheet.createRow(r);

			for (int c = 0; c < cols; c++) {

				XSSFCell cell = row.createCell(c);
				Object cred = data[r][c];

				if (cred instanceof String) {
					cell.setCellValue((String) cred);
				} else if (cred instanceof Integer) {
					cell.setCellValue((Integer) cred);
				} else if (cred instanceof Boolean) {
					cell.setCellValue((Boolean) cred);
				} else {
					cell.setCellValue((Integer) cred);
				}

			}
		}
		String projectpath = System.getProperty("user.dir");

		String filepath = projectpath+"/Data/ExcelWrite2.xlsx";
		FileOutputStream fos = new FileOutputStream(filepath);

		try {
			workbook.write(fos);

			workbook.close();

			fos.close();

			System.out.println("ExcelWrite2.xlsx created Successfully");
		} catch (Exception e) {
			System.out.println("Error occured while writing excel sheet");
			e.printStackTrace();
		}
	}

}
