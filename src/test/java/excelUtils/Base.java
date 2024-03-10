package excelUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import javax.swing.plaf.synth.SynthOptionPaneUI;

public class Base {

	static String Projectpath = System.getProperty("user.dir");

	static String file = null;

	static FileInputStream fis = null;

	public static FileInputStream getExcelSheet() {

		try {
			file=Projectpath+"/Data/Test.xlsx";
			fis = new FileInputStream(file);
			return fis;

		} catch (FileNotFoundException e) {
			System.out.println("Excel file not found / unable to open");
			e.getMessage();
		}
		return fis;

	}

}
