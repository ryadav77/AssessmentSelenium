package Selenium;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.By.ByClassName;
import org.openqa.selenium.By.ByName;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;


public class Assessment {
	   public static String vSearch;     //For Searching the Element
	   public static String vExecute;
	   public static int xlRows;         //For number of Rows from the Dataset
	   public static int xlCols;         //For number of Coloums from the Dataset
	   public static String xData[][];   //For Storing 2D Array Data.
	   
	   public static void main(String[] args) throws Exception {
		
		
	    
	    
	        xlRead("C:\\Users\\ryadav77\\Desktop\\YahooDDF.xls"); //Reading from the path of EXCEL File
	        for(int i=1;i<xlRows;i++) //Iterating over all the Rows
	        { 
	    	if(xData[i][1].equals("Y"))//Filter for allowing only the Companies we want to Execvute
	    	{
	    	vSearch=xData[i][0];       //Reading the company name to search
	    	System.setProperty("webdriver.chrome.driver", "C:\\Selenium\\chromedriver.exe");
			WebDriver myDriver= new ChromeDriver(); //Choosing the Chrome Driver
			myDriver.manage().window().maximize();
		    myDriver.get("https://in.yahoo.com/?p=us");
		    
		    myDriver.findElement(By.name("p")).sendKeys(vSearch);
		    Thread.sleep(2000);
		    myDriver.findElement(By.xpath("/html/body/div[2]/div/div/div/div/div/div/div[2]/div/div/form/table/tbody/tr/td[2]/button/i")).click();
		    Thread.sleep(5000);
		    String title=myDriver.getTitle();
		    xData[i][2]=title;
		  
		   
		    myDriver.close();
	    
	    }
	    	  xlwrite("C:\\Users\\ryadav77\\Desktop\\YahooDDF.xls", xData); //Writing the Data into Excel File
	    }
	    
	    
	   
	
	}
	  public static void xlRead(String sPath) throws Exception//Function to Read from our DataBase
		{
			File myFile=new File(sPath);
			FileInputStream myStream=new FileInputStream(myFile);
			HSSFWorkbook myworkbook=new HSSFWorkbook(myStream);
			HSSFSheet mySheet=myworkbook.getSheetAt(0);
			xlRows=mySheet.getLastRowNum()+1;
			xlCols=mySheet.getRow(0).getLastCellNum();
			xData=new String[xlRows][xlCols];
			for(int i=0;i<xlRows;i++)
			{
				HSSFRow row=mySheet.getRow(i);
				for(short j=0;j<xlCols;j++)
				{
					HSSFCell cell=row.getCell(j);
					String value=cellToString(cell);
					xData[i][j]=value;
					System.out.print("-"+xData[i][j]);
				}
				System.out.println();
			}
		}
			public static String cellToString(HSSFCell cell)
			{
				int type=cell.getCellType();
				Object result;
				switch(type)
				{
				case HSSFCell.CELL_TYPE_NUMERIC:
				result=cell.getNumericCellValue();
				break;
				case HSSFCell.CELL_TYPE_STRING:
				result=cell.getStringCellValue();
				break;
				case HSSFCell.CELL_TYPE_FORMULA:
				throw new RuntimeException("We cannot evaluate formula");
				case HSSFCell.CELL_TYPE_BLANK:
				result="-";
				case HSSFCell.CELL_TYPE_BOOLEAN:
				result=cell.getBooleanCellValue();
				case HSSFCell.CELL_TYPE_ERROR:
				result="This cell has some error";
				default:
				throw new RuntimeException("We do not support this cell type");
				}
				return result.toString();
				
			}
			
			public static void xlwrite(String xlpath1,String[][] xData) throws Exception//Function to Write back to the Database
			{
				System.out.println("Inside XL Write");
				File myFile1=new File(xlpath1);
				FileOutputStream fout=new FileOutputStream(myFile1);
				HSSFWorkbook wb=new HSSFWorkbook();
				HSSFSheet mySheet1=wb.createSheet("TestResults");
				for(int i=0;i<xlRows;i++)
				{
					HSSFRow row1=mySheet1.createRow(i);
					for(short j=0;j<xlCols;j++)
					{
						HSSFCell cell1=row1.createCell(j);
						cell1.setCellType(HSSFCell.CELL_TYPE_STRING);
						cell1.setCellValue(xData[i][j]);
					}
				}
				wb.write(fout);
				fout.flush();
				fout.close();
			}

	
	
}
	