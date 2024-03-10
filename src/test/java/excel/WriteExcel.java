package excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class WriteExcel extends Base {

	public static void writeExcel(String sheetname, int RowNum, int ColNum, String data) {

		try {
			FileInputStream fis = new FileInputStream(excelSheet());

			Workbook wb = WorkbookFactory.create(fis);

			Sheet sheet = wb.getSheet(sheetname);

			Row row = sheet.getRow(RowNum);

			Cell col = row.createCell(ColNum);

			col.setCellValue(data);
			FileOutputStream fos = new FileOutputStream(excelSheet());
			wb.write(fos);

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public static void main(String[] args) {

		writeExcel("Sheet1", 1, 2, "Pass");
	}

}
