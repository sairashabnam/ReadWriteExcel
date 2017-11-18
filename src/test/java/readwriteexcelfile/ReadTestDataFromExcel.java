package readwriteexcelfile;

import java.util.Map;

import org.junit.Test;

import com.excelreadwrite.ExcelReaderWriter;

public class ReadTestDataFromExcel {
	
	public ExcelReaderWriter excelReader = new ExcelReaderWriter();
/**	Pass sheetName and EmployeeId as parameters for reading data from excel file
*/	public Map<String, String> employeeDetails = excelReader.readExcelData("EmployeeDOB", "EM00512");

	@Test
	public void test() {
		System.out.println(employeeDetails.get("EmployeeName"));
		System.out.println(employeeDetails.get("Department"));
		System.out.println(employeeDetails.get("DOB"));
	}

}
