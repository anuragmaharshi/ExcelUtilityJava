package ExcelCode;

import org.testng.annotations.Test;

import excelCode.ExcelReader;

import org.testng.annotations.BeforeMethod;

import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.testng.annotations.AfterMethod;

public class NewTest {
  @BeforeMethod
  public void beforeMethod() {
	  System.out.println("*******************************************************************************");
  }

  @AfterMethod
  public void afterMethod() {
	  System.out.println("##############################################################################");
  }

  @Test(description="//checking excel with invalid path")
  public void Test1() {    
		ExcelReader excObj = new ExcelReader("C:\\TestData\\Test1.xlsx",1);
		excObj.printExcelData();
		excObj.closeAll();
  }

  @Test (description="//checking excel with valid path and invalid sheet no")
  public void Test2() {
		ExcelReader excObj1 = new ExcelReader("C:\\TestData\\Test1.xlsx",3);
		excObj1.printExcelData();
		excObj1.closeAll();
  }
  
  @Test (description="//checking excel with valid path and invalid sheet name")
  public void Test3() {
		ExcelReader excObj1 = new ExcelReader("C:\\TestData\\Test1.xlsx","hello");
		excObj1.printExcelData();
		excObj1.closeAll();
  }
  
  @Test (description="//checking excel with valid path and valid sheet name with spaces")
  public void Test4() {
		ExcelReader excObj1 = new ExcelReader("C:\\TestData\\Test1.xlsx","Sheet new");
		excObj1.printExcelData();
		excObj1.closeAll();
  }
  
  @Test (description="//checking excel with valid path and valid sheet no")
  public void Test5() {
		ExcelReader excObj1 = new ExcelReader("C:\\TestData\\Test1.xlsx",0);
		excObj1.printExcelData();
		excObj1.closeAll();
  }
  @Test (description="//checking get cell data funtion")
  public void Test6() {
	    ExcelReader excObj1 = new ExcelReader("C:\\TestData\\Test1.xlsx",0);
		System.out.println(excObj1.getCellData(-1, -1));
		System.out.println(excObj1.getCellData(0, -1));
		System.out.println(excObj1.getCellData(-1, 0));
		System.out.println(excObj1.getCellData(0, 0));
		System.out.println(excObj1.getCellData(1, 1));
		System.out.println(excObj1.getCellData(1, 3));
		System.out.println(excObj1.getCellData(3, 1));
		System.out.println(excObj1.getCellData(3, 3));
		System.out.println(excObj1.getCellData(3, 10));
		System.out.println(excObj1.getCellData(30, 10));
		excObj1.closeAll();
  }
  
  @Test (description="//checking get column data funtion")
  public void Test7() {
	  ExcelReader excObj1 = new ExcelReader("C:\\TestData\\Test1.xlsx",0);
	  List<String> data = excObj1.getColumnData(0);
	  System.out.println("printing 0th column data");
	  for(String str : data){
		  System.out.println(str);
	  }
	  System.out.println("printing -1 column data");
	  data = excObj1.getColumnData(-1);
	  for(String str : data){
		  System.out.println(str);
	  }
	  
	  System.out.println("printing 1 column data");
	  data = excObj1.getColumnData(1);
	  for(String str : data){
		  System.out.println(str);
	  }
	  
	  System.out.println("printing 3 column data");
	  data = excObj1.getColumnData(3);
	  for(String str : data){
		  System.out.println(str);
	  }
	  
	  System.out.println("printing 10 column data");
	  data = excObj1.getColumnData(10);
	  for(String str : data){
		  System.out.println(str);
	  }
	  excObj1.closeAll();
	  
  }

  @Test (description="//checking get row data funtion")
  public void Test8() {
	  ExcelReader excObj1 = new ExcelReader("C:\\TestData\\Test1.xlsx",0);
	  List<String> data = excObj1.getRowData(-1);
	  System.out.println("printing -1 row data");
	  for(String str : data){
		  System.out.println(str);
	  }
	  
	  data = excObj1.getRowData(0);
	  System.out.println("printing 0 row data");
	  for(String str : data){
		  System.out.println(str);
	  }
	  
	  data = excObj1.getRowData(1);
	  System.out.println("printing 1 row data");
	  for(String str : data){
		  System.out.println(str);
	  }
	  
	  data = excObj1.getRowData(10);
	  System.out.println("printing 10 row data");
	  for(String str : data){
		  System.out.println(str);
	  }
  }

  @Test (description="//checking getRowNumForData funtion")
  public void Test9() {
	  ExcelReader excObj1 = new ExcelReader("C:\\TestData\\Test1.xlsx",0);
	  System.out.println(excObj1.getRowNumForData("Anurag", -1));
	  System.out.println(excObj1.getRowNumForData("Anurag", 1));
	  System.out.println(excObj1.getRowNumForData("Anurag", 0));
	  System.out.println(excObj1.getRowNumForData("Maha", 1));
	  System.out.println(excObj1.getRowNumForData("Maha", 0));
  }
  
  @Test (description="//checking Get Row data in Map funtion")
  public void Test10() {
	  System.out.println("checking row data in map");
	  ExcelReader excObj1 = new ExcelReader("C:\\TestData\\Test1.xlsx",0);
	  //checking with invalid row no
	  Map<String, String> map = excObj1.getRowDataInMap(-1);
	  Set<String> key= map.keySet();
	  Iterator<String> iter = key.iterator();
	  while(iter.hasNext()){
		  System.out.println(map.getOrDefault(iter.next(), ""));
	  }
	  
	  map = excObj1.getRowDataInMap(1);
	  key= map.keySet();
	  iter = key.iterator();
	  while(iter.hasNext()){
		  String key1 =iter.next();
		  System.out.println("key is "+ key1 + " value = "+map.getOrDefault(key1, ""));
	  }
	  System.out.println("value for key Name is "+ map.getOrDefault("Name", "") );
	  System.out.println("value for key name is "+ map.getOrDefault("name2", "") );
	  map = excObj1.getRowDataInMap(0);
	  key= map.keySet();
	  iter = key.iterator();
	  while(iter.hasNext()){
		  String key1 =iter.next();
		  System.out.println("key is "+ key1 + " value = "+map.getOrDefault(key1, ""));
	  }
	  
	  map = excObj1.getRowDataInMap(5);
	  key= map.keySet();
	  iter = key.iterator();
	  while(iter.hasNext()){
		  String key1 =iter.next();
		  System.out.println("key is "+ key1 + " value = "+map.getOrDefault(key1, ""));
	  }
	  System.out.println("value for key Name is "+ map.getOrDefault("Name", "") );
	  System.out.println("value for key name is "+ map.getOrDefault("name2", "") );
	  
  }
  

}
