package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExelRead extends Base {

	public String readExcel(String sheetname, int RowNum, int ColNum) {

		String data = "";

		try {
			FileInputStream fis = new FileInputStream(excelSheet());
			Workbook wb = WorkbookFactory.create(fis);
			Sheet s = wb.getSheet(sheetname);

			Row row = s.getRow(RowNum);
			Cell col = row.getCell(ColNum);
			data = col.getStringCellValue();

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			System.out.println("unable to open excel sheet");
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return data;
	}

	public static void main(String[] args) throws Exception {

		ExelRead exc = new ExelRead();
		String username = exc.readExcel("Sheet1", 1, 0);

		//System.out.println("Username is :" + username);
		String password = exc.readExcel("Sheet1", 1, 1);

		//System.out.println("Password is :" + password);
		
		readAllData("Sheet1");

	}
	
	public static void readAllData(String sheetname) throws Exception {
		FileInputStream fis = new FileInputStream(excelSheet());
		Workbook wb = WorkbookFactory.create(fis);
		Sheet s = wb.getSheet(sheetname);
		
		Row row=null;
		
		int i = 1;
		while((row=s.getRow(i))!=null) {
			
			System.out.println("username is :"+row.getCell(0).getStringCellValue());
			
			System.out.println("Password is :"+row.getCell(1).getStringCellValue());
			
			i++;
		}
		
	}
}
