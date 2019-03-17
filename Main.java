package com.mcas;
import java.sql.*;
import java.util.ArrayList;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class connection {
	public static void main(String[] args) throws Exception {
		
		File src = new File("C:\\Users\\U1167500\\Documents\\MCAS_Excel\\input_data.xlsx");
		
		FileInputStream fis = new FileInputStream(src);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet1= wb.getSheet("Sheet1");
		
		//to get the last row number 
		int num = sheet1.getLastRowNum();
		//System.out.println(num);
		
		if(num==0){
			System.out.println("Please enter the data");
		}
		//creating arrayList
		ArrayList<Integer> data = new ArrayList<Integer>();
		for(int i=1; i<=num; i++){
			//System.out.println("i is "+i);
			String fName = sheet1.getRow(i).getCell(0).getStringCellValue();
			String lName = sheet1.getRow(i).getCell(1).getStringCellValue();
			//System.out.println(fName+lName);
			// writing data into excel
			
			data = getData(fName,lName);
			for(int j : data){
				sheet1.getRow(i).createCell(2).setCellValue("p"+j);
				//System.out.println("P"+j);
			}
		}
		//setting output stream for writing
		FileOutputStream fout = new FileOutputStream(src);
		
		wb.write(fout);
		
		wb.close();
		
		System.out.println("completed");
		
		
		//System.out.println("start");
		/*
		String firstName[]={"Supreet","Jagjeet"};
		String lastName[] ={"Sidhu","Dhillon"};
		ArrayList<Integer> data = new ArrayList<Integer>();
		
		
		if(firstName.length==0 & lastName.length==0){
			System.out.println("Please Enter the first name and last name");
		}
		else if (firstName.length == lastName.length){
			int len = firstName.length;
			for(int i=0; i<len; i++){
				data = getData(firstName[i],lastName[i]);
				for(int j : data){
					System.out.println("P"+j);
				}
				
				
			}
		}else if (firstName.length > lastName.length){
			System.out.println("please provide all last names");
		}else if(firstName.length < lastName.length){
			System.out.println("please provide all first names");
		}
		// printing the element in array with p at the front
		*/		
	}
	public static Connection getConnection(){
		
		//defining variables for connection parameters
		String driver = "oracle.jdbc.driver.OracleDriver";
		String url = "jdbc:oracle:thin:@(DESCRIPTION = (FAILOVER = ON) (LOAD_BALANCE = YES)"
				+ "(ADDRESS = (PROTOCOL = TCP)(HOST = USDFW21DB54V-sca.mrshmc.com)(PORT = 1521))"
				+ "(CONNECT_DATA = (SERVER = DEDICATED) (SERVICE_NAME = oltp149)       "
				+ "(FAILOVER_MODE = (TYPE = SELECT) (METHOD = BASIC) (RETRIES = 10) (DELAY = 5))))";
		String userName = "p37602";
		String passWord = "Cold_2019";
		
		try{
			Class.forName(driver);
			Connection con = DriverManager.getConnection(url,userName,passWord);
			//System.out.println("connected");
			return con;
		}catch(Exception e){System.out.println(e);}
		
		return null;
	}
	
	public static ArrayList<Integer> getData(String First_Name, String Last_Name) throws Exception{
		try{
			//make connection by calling connection function
			Connection con = getConnection();
			//String query = "select EMP_NO from CORP.EMPLOYEE where EMP_LAST_NM = '"+Last_Name+"' AND EMP_FIRST_NM = '"+First_Name+"'";
			String query = "select EMP_NO from CORP.EMPLOYEE where EMP_LAST_NM = LTRIM(RTRIM('"+Last_Name+"')) AND EMP_FIRST_NM = LTRIM(RTRIM('"+First_Name+"'))";;
			//using preparedstatement to create a query
			PreparedStatement stmt = con.prepareStatement(query);
			//execute query and saved result into ResultSet type object
			ResultSet result = stmt.executeQuery();
			//instaniate an ArrayList to save result into arrayList
			ArrayList<Integer> array = new ArrayList<Integer>();
			//while 
			//while(result.next()==false){
				//array.add(0000);
			//}
			while(result.next()){
				//System.out.println("I am in array");
				//System.out.println(result.getInt(1));
				if (result != null) {
				array.add(result.getInt(1));
			}
					
			}
			//System.out.println("Data is stored in array");
			return array;
			
		}catch(Exception e){System.out.println(e);}
		return null;
			
	}

}

