package com.douwan.admin.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
public class JxlUtil {

	 public static List<String> xls2List(File file){
		          List<String> result=new ArrayList<String>();
		          try{
		        	  
		        	  //以路径读取
		        	  //InputStream inp = new FileInputStream(path); 
		        	  
		        	  //以文件读取
		             FileInputStream fis = new FileInputStream(file); 
		              Workbook rwb = Workbook.getWorkbook(fis);
		             Sheet[] sheet = rwb.getSheets();   
		              for (int i = 0; i < sheet.length; i++) {   
		                 Sheet rs = rwb.getSheet(i);   
		                  for (int j = 0; j < rs.getRows(); j++) {   
		                     Cell[] cells = rs.getRow(j);   
		                    for(int k=0;k<cells.length;k++) {
		                    	result.add(cells[k].getContents());
		                    }
		                 }   
		              } 
		              //路径读取
		              //inp.close(); 
		              
		              //文件读取
		              fis.close();
		        }catch(Exception e){
		              e.printStackTrace();
		         }
		         return result;
		      }
	 
	 public static void main(String[] args){
        
		 //以路径读取
          // List<String> result = xls2List("E:/111.xls");
           
		 //以文件读取
           File file = new File("E:/111.xls");
           List<String> result = xls2List(file);
           for (int i = 0; i < result.size(); i++) {
        	   System.out.println(result.get(i));
		}
      }
}
