package ClCompare.comparision_xcl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hpsf.Array;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App {
		static Boolean check = false;
		
		//Change column number whatever you want to take data
		public static int columnNumForFirst = 0;
		public static int columnNumForSecond = 0;
		public static int columnNum2ForFirst = 1;
		public static int columnNum2ForSecond = 1;

		public static void main(String args[]) throws IOException {

			try {
				//ArrayList arr3 = new ArrayList();
				List<String> arr3 = new ArrayList<String>();
				ArrayList<String> arr4 = new ArrayList<String>(50000);

				String file1 ="C:\\Users\\Dell\\Desktop\\Docs\\Looker\\Google - 26K.xlsx";

				String file2 = "C:\\Users\\Dell\\Desktop\\Docs\\Looker\\Looker - 26K.xlsx";

				Map<String, List> arr1 = StoreArraysToHashMap(file1,columnNumForFirst);
				Map<String, List> arr2 = StoreArraysToHashMap(file2,columnNumForFirst);

				
					Set<String> keysarr1 = arr1.keySet();
					Set<String> keysarr2 = arr2.keySet();
					Collection<List> valuesarr1 = arr1.values();
					Collection<List> valuesarr2 = arr2.values();
					Iterator i = keysarr1.iterator();
					Object[] ka1 =keysarr1.toArray();
					Object[] ka2 =keysarr2.toArray();
					int k =0;
					int l =0;
					for (int m=0;m<ka1.length;m++)
					{
						for (int n=0;n<ka2.length;n++)
						{
							if(ka1[m].equals(ka2[n]))
							{
								if(arr1.get(ka1[m]).equals(arr2.get(ka2[n])))
								{
									k++;
								}
							else {
								arr3.add("For Key -"+(String) (ka1[m])+" Data in Google Sheet data is - "+arr1.get(ka1[m]).toString()+" And in Looker Sheet data is - "+arr2.get(ka2[n]).toString());
							}}else {
								l=(arr1.size()-(k+2));
							}
						}
					}
				System.out.println("Total number of mismatched records -" + l);
				System.out.println("Total number of matched records -" + k);
				
				WriteDataToExcel(arr3);
				  
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
			

		}

		//Store data to HashMap
		private static  Map<String, List> StoreArraysToHashMap(String path, int columnNum) throws IOException {
			FileInputStream file = new FileInputStream(new File(path));
			ArrayList arr = new ArrayList();
			// Get the workbook instance for XLSX file
			XSSFWorkbook workbook1 = new XSSFWorkbook(file);

			// Get only first sheet from the workbook
			XSSFSheet sheet1 = workbook1.getSheetAt(0);

			
			// Get iterator to all the rows in current sheet1ZZ
			Iterator<Row> rowIterator1 = sheet1.iterator();
			//System.out.println(sheet1.getRow(0).getPhysicalNumberOfCells());
			//System.out.println("++++"+sheet1.getRow(0).getCell(1).getStringCellValue());
			
			Map<String, List> myMap = new HashMap<>();
			Row row = rowIterator1.next();
			// For each row, iterate through all the columns
			Iterator<Cell> cellIterator1 = row.cellIterator();
			Cell cell = cellIterator1.next();
			if(sheet1.getRow(0).getCell(0).getStringCellValue().equals("Comparison column") && sheet1.getRow(0).getCell(1).getStringCellValue().equals("AdWords AdGroup Total Impressions"))
			{
				for(int i =0; i<1;i++)
				{
					for(int j=1;j<sheet1.getLastRowNum();j++)
					{
						ArrayList ar = new ArrayList<String>();
						List subArr = new ArrayList<String>();
						switch (sheet1.getRow(j).getCell(i).getCellType())
						{
						case NUMERIC:
							//System.out.print(cell.getNumericCellValue());
							ar.add(sheet1.getRow(j).getCell(i).getStringCellValue());
							break;
						case STRING:
							ar.add(sheet1.getRow(j).getCell(i).getStringCellValue());
							DataFormatter formatter = new DataFormatter();
							String val = formatter.formatCellValue(sheet1.getRow(j).getCell(++i));
							ar.add(val);
							i--;
							//System.out.print(cell.getStringCellValue());
							break;
						case BOOLEAN:
							ar.add(sheet1.getRow(j).getCell(i).getStringCellValue());
							//System.out.print(cell.getBooleanCellValue());
							break;	
						}
						
						ar.add("f");
						subArr= ar.subList(1, 2);
						myMap.put( ar.get(0).toString(), subArr);
						//System.out.println("@@@@"+myMap);
						
					}
				}
				//for(Map.Entry m : myMap.entrySet()){    
				 //   System.out.println(m.getKey()+" "+m.getValue());    
				//}
				//System.out.println("--"+myMap.size());
				file.close();
			}
			return myMap;
		}
		
		public static void  WriteDataToExcel(List<String> arr3) throws IOException {
			 //Create blank workbook
		      XSSFWorkbook workbook = new XSSFWorkbook();
		      
		      //Create a blank sheets
		      XSSFSheet spreadsheet = workbook.createSheet( "Employee Info ");

		      //Create row object
		      XSSFRow row;

		      //Iterate over data and write to sheet
		      int rowid = 0;
		      
		      for (int i =0; i<arr3.size(); i++) {
		    	  int cellid = 0;
		         row = spreadsheet.createRow(rowid++);
		            Cell cell = row.createCell(cellid++);
		            cell.setCellValue((String) arr3.get(i));
		      }
		      //Write the workbook in file system
		      FileOutputStream out = new FileOutputStream(
		         new File("C:\\Users\\Dell\\Desktop\\Docs\\Looker\\Writesheet1.xlsx"));
		      
		      workbook.write(out);
		      out.close();
		      System.out.println("Writesheet.xlsx written successfully");
		      //System.out.println("arr3 list values, here arr1 has some values which arr2 DOES NOT have : " + arr3);
			   }
		
}