package excel;

import java.io.File;

public class Base {

	public static File excelSheet() {

		String projectpath = System.getProperty("user.dir");

		File file = new File(projectpath + "/Data/Test.xlsx");
		return file;

	}

}
