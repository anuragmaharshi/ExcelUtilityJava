package ExcelCode;

import org.testng.annotations.Test;

import excelCode.ExcelWriter;

import org.testng.annotations.BeforeMethod;

import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.testng.annotations.AfterMethod;

public class WriterTest {
  
  @BeforeMethod
  public void beforeMethod() {
	  System.out.println("****************************************************************");
  }

  @AfterMethod
  public void afterMethod() {
	  System.out.println("_________________________________________________________________");
	  
  }
  @Test(description="//checking excel with new workbook and new sheet")
  public void test1() {
	  ExcelWriter wr = new ExcelWriter("C:\\TestData\\Test1Write.xlsx","Sheet 1");
	  wr.excelWrite(0, 0, "Hello World");
  }
  
  @Test(description="//checking excel with Existing work book and existing sheet ")
  public void test2() {
	  ExcelWriter wr = new ExcelWriter("C:\\TestData\\Test1Write.xlsx","Sheet 1");
	  wr.excelWrite(1, 0, "Hello World New Cell");
  }
  
  @Test(description="//checking excel with Existing work book and new sheet ")
  public void test3() {
	  ExcelWriter wr = new ExcelWriter("C:\\TestData\\Test1Write.xlsx","Sheet 2");
	  wr.excelWrite(1, 0, "Hello World New Sheet");
  }
  
  @Test(description="//Writing in excel in Existing work book, string value to set")
  public void test4() {
	  ExcelWriter wr = new ExcelWriter("C:\\TestData\\Test1Write.xlsx","Sheet 2");
	  wr.excelWrite(0, 0, "Hello World New Sheet");
  }
  
  @Test(description="//Writing in excel in Existing work book, Double value to set")
  public void test5() {
	  ExcelWriter wr = new ExcelWriter("C:\\TestData\\Test1Write.xlsx","Sheet 2");
	  wr.excelWrite(1, 0, 12.23);
  }
  
  @Test(description="//Writing in excel in Existing work book, Boolean value to set")
  public void test6() {
	  ExcelWriter wr = new ExcelWriter("C:\\TestData\\Test1Write.xlsx","Sheet 2");
	  wr.excelWrite(2, 0, true);
  }
  
  @Test(description="//Writing in excel in Existing work book, date value to set")
  public void test7() {
	  ExcelWriter wr = new ExcelWriter("C:\\TestData\\Test1Write.xlsx","Sheet 2");
	  Date date = new Date();
	  wr.excelWrite(3, 0, date);
  }
  
  @Test(description="//Writing in excel in Existing work book, new map, in new sheet")
  public void test8() {
	  ExcelWriter wr = new ExcelWriter("C:\\TestData\\Test1Write.xlsx","Sheet 3");
	  Map<String,String> data = new HashMap<String,String>();
	  //Currently we need to pass integer no as string. Its type casting is not done
	  //As this is a new sheet, there is no column header. It will add keys as column header
	  data.put("Test_id", "123");
	  data.put("Test_Desc", "Test Description");
	  data.put("Status", "Pass");
	  data.put("Date", "10-10-2017");
	  wr.excelWrite(1, data);
  }
  
  @Test(description="//Writing in excel in Existing work book, new map, in old sheet.Run the test after creating sheet 4 with headers")
  public void test9() {
	  ExcelWriter wr = new ExcelWriter("C:\\TestData\\Test1Write.xlsx","Sheet 4");
	  Map<String,String> data = new HashMap<String,String>();
	  //Currently we need to pass integer no as string. Its type casting is not done
	 
	  data.put("Test_id", "123");
	  data.put("Test_Desc", "Test Description");
	  data.put("Status", "Pass");
	  data.put("Date", "10-10-2017");
	  //it will override the old data in the specified row no.
	  wr.excelWrite(1, data);
  }
  
  @Test(description="//Writing in excel in Existing work book, new map, in old sheet, in last row.Run the test after creating sheet 4 with headers")
  public void test10() {
	  ExcelWriter wr = new ExcelWriter("C:\\TestData\\Test1Write.xlsx","Sheet 4");
	  Map<String,String> data = new HashMap<String,String>();
	  //Currently we need to pass integer no as string. Its type casting is not done
	 
	  data.put("Test_id", "124");
	  data.put("Test_Desc", "Test Description 2");
	  data.put("Status", "Pass");
	  data.put("Date", "10-10-2017");
	  wr.excelWrite(data);
  }
  
  @Test(description="//Writing in excel in Existing work book, new list of map, in new sheet.")
  public void test11() {
	  ExcelWriter wr = new ExcelWriter("C:\\TestData\\Test1Write.xlsx","Sheet 5");
	  Map<String,String> data = new HashMap<String,String>();
	 
	  data.put("Test_id", "125");
	  data.put("Test_Desc", "Test Description 3");
	  data.put("Status", "Pass");
	  data.put("Date", "10-10-2017");
	  
	  Map<String,String> data2 = new HashMap<String,String>();
	  //Currently we need to pass integer no as string. Its type casting is not done
	 
	  data2.put("Test_id", "126");
	  data2.put("Test_Desc", "Test Description 4");
	  data2.put("Status", "Pass");
	  data2.put("Date", "10-10-2017");
	  
	  List<Map<String,String>> dataAll = new ArrayList<Map<String, String>> ();
	  dataAll.add(data);
	  dataAll.add(data2);
	  wr.excelWrite(dataAll);
  }
  
  @Test(description="//Writing in excel in Existing work book, new list of map, in old sheet. Re run this sheet after creating sheet 4")
  public void test12() {
	  ExcelWriter wr = new ExcelWriter("C:\\TestData\\Test1Write.xlsx","Sheet 4");
	  Map<String,String> data = new HashMap<String,String>();
	 
	  data.put("Test_id", "125");
	  data.put("Test_Desc", "Test Description 3");
	  data.put("Status", "Pass");
	  data.put("Date", "10-10-2017");
	  
	  Map<String,String> data2 = new HashMap<String,String>();
	  //Currently we need to pass integer no as string. Its type casting is not done
	 
	  data2.put("Test_id", "126");
	  data2.put("Test_Desc", "Test Description 4");
	  data2.put("Status", "Pass");
	  data2.put("Date", "10-10-2017");
	  
	  List<Map<String,String>> dataAll = new ArrayList<Map<String, String>> ();
	  dataAll.add(data);
	  dataAll.add(data2);
	  wr.excelWrite(dataAll);
  }
}
