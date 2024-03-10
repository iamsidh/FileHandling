package dataprovider;

import org.testng.annotations.*;

public class TestExcel {
	
	@Test(dataProvider="LoginData",dataProviderClass=ExcelDataProvider.class)
	public void getData(String username,String password) {
		
		System.out.println("Username is :"+username);
		System.out.println("Password is :"+password);
	}
}
