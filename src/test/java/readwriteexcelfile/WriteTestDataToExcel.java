package readwriteexcelfile;

import org.junit.Test;

import com.excelreadwrite.ExcelReaderWriter;

public class WriteTestDataToExcel {
	
	public ExcelReaderWriter excelReader = new ExcelReaderWriter();

	@Test
	public void test() {
		/** Pass parameters
		 * sheetName,
		 * Employee Id to which testdata has to insert
		 * Column Name in which testdata has to insert
		 * Testdata*/
		excelReader.writeToExcel("EmployeeDOB", "EM00512", "CheckTest", "TestData");
	}

}
