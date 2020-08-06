package excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.stream.Collectors;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;


/*用于检查排期会文档
1、检查是否有重复项目
2、统计项目总数
3、统计每个开发的开发量*/

public class Main {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
			
	//	String fileName= "/Users/mac/Documents/企业网银排期计划表2020.8-9月-会前v1.xls";
		String fileName= "/Users/apple/Desktop/企业网银排期计划表2020.8-9月-会后V1.0.xls";
		InputStream input = new FileInputStream(fileName);  
        POIFSFileSystem fs = new POIFSFileSystem(input);  
        HSSFWorkbook wb = new HSSFWorkbook(fs);  
        HSSFSheet sheet = wb.getSheetAt(0);  
        Iterator rows = sheet.rowIterator();
        
        List<String > list = new ArrayList<String>(); 
        String info[][]= new String[1000][2];
        int index=0;
        while (rows.hasNext()) {  
            HSSFRow row = (HSSFRow) rows.next();  
            //row 行的意思
           // System.out.println("Row #" + row.getRowNum());  
            // Iterate over each cell in the row and print out the cell"s  
            // content  
            
            Iterator cells = row.cellIterator(); 
          
            while (cells.hasNext()) {
        	//System.out.println("=="+cells.toString());
        	//cell  单元格
            HSSFCell cell = (HSSFCell) cells.next(); 
            if(!cell.toString().equals("")&&!(cell.toString().trim().length()== 0))
            {
             //   System.out.println("cell = "+cell.getColumnIndex()+ " "+cell.toString());
                
                
                if(cell.getColumnIndex()==7) {
                	//人天数
                	info[index][0]=cell.toString();
                }
                if(cell.getColumnIndex()==8){
                	//开发人员信息
                	info[index][1]=cell.toString();
                	
                }
                
            }
            //row=7 本次排入人天数目
            
            //row=8 开发人员
            //统计项目相关
            //row=2 项目
            if(cell.getColumnIndex()==2&&cell.toString()!=""&&!cell.toString().equals("项目名称")) {
            	  if(list.contains(cell.toString()) ) {
            		  System.err.println("******重复信息："+"行数:"+row.getRowNum()+" 列数:" + (cell.getColumnIndex()+1)+" "+cell.toString());
            	  }else {
            		//  sumProject++;
            		  list.add(cell.toString());
            	  }
            	 // System.out.println("Cell #" + cell.getColumnIndex()+" "+cell.toString());  
            }
         }	
        index++;
       // System.out.println("项目总数："+sumProject);
	}
        
        int indexSum=0;
        Iterator isList = list.iterator();
        while(isList.hasNext()) {
        	System.out.println("项目"+(++indexSum)+"名称:"+isList.next());
        }
        System.err.println("项目总数:"+list.size()); 
        
        //统计每个开发工作量
        System.out.println("============统计每个开发工作量==========");
        
        Map<String,Double> map = new HashMap<String,Double>();
        
        for(int i =0 ;i<1000;i++) {
        	if(info[i][0]==""||info[i][1]==""||info[i][0]==null||info[i][1]==null||info[i][0].endsWith("本")) {
        		continue;
        	}else {
        		if(map.containsKey(info[i][1])) {
        			map.put(info[i][1], map.get(info[i][1])+Double.valueOf(info[i][0]));//如果已经存在就相加
        		}else {
        			map.put(info[i][1], Double.valueOf(info[i][0]));//如果map中不存在就存入
        		}
        		
        	//	System.out.println(info[i][0]+" "+info[i][1]);
        	}
        	
        }
        List<Entry<String, Double>> list1 = map.entrySet().stream()
        	      .sorted((entry1, entry2) -> entry2.getValue().compareTo(entry1.getValue())) //降序
        	      .collect(Collectors.toList());
        
        Iterator<Entry<String, Double>> it=list1.iterator();//.entrySet().iterator();
        while(it.hasNext()){ 
        	Entry<String, Double> entry=it.next();
        	if(entry.getValue()< 20) {
        		System.out.println(Math.round(entry.getValue())+" "+entry.getKey());
        	}else {
        		System.out.println(Math.round(entry.getValue())+" "+entry.getKey());
        	}
        }
		/*
		 * 
		 * Iterator<Entry<String, Double>> it=map.entrySet().iterator();
		 * while(it.hasNext()){ Entry<String, Double> entry=it.next();
		 * System.out.println("key="+entry.getKey()+","+"value="+entry.getValue()); }
		 */

}
}

