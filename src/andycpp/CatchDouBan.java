package andycpp;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Writer;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class CatchDouBan {  
	  
    private CatchDouBan() {  
  
    }  
    
    //存放读取出来的数据
    List<HashMap<String,String>> list = new ArrayList<HashMap<String,String>>();
     
    /** 
     * 每页显示记录条数 
     */  
    public static final int NUM = 20;  
  
    /** 
     * 拼接分页 
     */  
    public static final String START = "&start=";  
    
  
    private static final CatchDouBan instance = new CatchDouBan();  
  
    public static CatchDouBan getInstance() {  
        return instance;  
    }  
    
    public void createExcel() {
    	try {  
            WritableWorkbook book  =  Workbook.createWorkbook( new  File( "D:\\前100本评分最高的书.xls" ));
            WritableSheet sheet  =  book.createSheet("第一页 ", 0 );
            Label num =  new Label(0, 0,"序号");
            sheet.addCell(num);
            Label title = new Label(1, 0,"书名");
            sheet.addCell(title);
            Label sorce = new Label(2, 0,"评分");
            sheet.addCell(sorce);
            Label people = new Label(3, 0, "评价人数");
            sheet.addCell(people);
            Label author = new Label(4, 0, "作者");
            sheet.addCell(author);
            Label press = new Label(5, 0,"出版社");
            sheet.addCell(press);
            Label date = new Label(6, 0,"出版日期");
            sheet.addCell(date);
            Label price = new Label(7, 0, "价格");
            sheet.addCell(price);
            
            for(int i=0; i<(list.size()>100?100:list.size());i++) {
            	 Label lNum = new Label(0, i+1, i+1+"");
                 sheet.addCell(lNum);
                 Label lName = new Label(1, i+1, list.get(i).get("sTitle"));
                 sheet.addCell(lName);
                 Label lSorce = new Label(2, i+1, list.get(i).get("eleSorce"));
                 sheet.addCell(lSorce);
                 Label lPeople = new Label(3, i+1, list.get(i).get("sPeople"));
                 sheet.addCell(lPeople);
                 Label lAuthor = new Label(4, i+1, list.get(i).get("sAuthor"));
                 sheet.addCell(lAuthor);
                 Label lPress = new Label(5, i+1, list.get(i).get("sPress"));
                 sheet.addCell(lPress);
                 Label lDate = new Label(6, i+1, list.get(i).get("sDate"));
                 sheet.addCell(lDate);
                 Label lPrice = new Label(7, i+1, list.get(i).get("sPrice"));
                 sheet.addCell(lPrice);
            }
            
            
            book.write();
            book.close();  
  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
    }
      
    public void getDoubanReview(String urlold){  
        try {           
        	
            
            int i = 0;
        	int j = 0;
            while(j < 100){
                String url = urlold + START + String.valueOf(i * NUM);  
                System.out.println(url);  
                Connection connection = Jsoup.connect(url);  
                Document document = connection.get();  
                Elements subject = document.select("li.subject-item"); 
                if(subject.size() == 0) {
                	break;
                }
                Iterator<Element> ulIter = subject.iterator();  
                while (ulIter.hasNext()) {  
                	HashMap<String,String> map = new HashMap<String,String>();
                    Element element = ulIter.next();  
                    Elements eleInfo = element.select("div.info"); 
                    Element eleTitle = eleInfo.select("a").first();            
                    String sTitle = eleTitle.html().replaceAll("<[^>]*>", "");   //书名
                    String eleSorce = eleInfo.select(".rating_nums").size() != 0 ?eleInfo.select(".rating_nums").first().text() : "";   //评分
                    String sPeople = eleInfo.select(".pl").text();  //评分人数
                    String regEx="[^0-9]";   
                    Pattern p = Pattern.compile(regEx);   
                    Matcher m = p.matcher(sPeople);   
                    sPeople = m.replaceAll("").trim();
                    
                   if(Integer.parseInt(sPeople) <1000) {
                    	continue;
                    }
                    
                    String[] array = eleInfo.select(".pub").text().split("/");
                    String sAuthor = array[0];             //作者
                    String sPrice = array[array.length-1];  //价格
                    String sDate = array[array.length-2];  //出版日期
                    String sPress = array[array.length-3];  //出版社
                    
                    j++;

                    map.put("mNum", j+"");
                    map.put("sTitle", sTitle);
                    map.put("eleSorce", eleSorce);
                    map.put("sPeople", sPeople);
                    map.put("sAuthor", sAuthor);
                    map.put("sPress", sPress);
                    map.put("sDate", sDate);
                    map.put("sPrice", sPrice);
                    
                    //去重
                    Boolean flag = true;
                    for(int k=0;k<list.size();k++) {
                    	if(sTitle.equals(list.get(k).get("sTitle"))) {
                    		flag = false;
                    	}
                    }
                    if(flag) {
                    	list.add(map);                    
                    }                   
                    //排序
                    Collections.sort(list, new Comparator<HashMap<String,String>>(){  
                    	  
                        /*   
                         * 返回负数表示：o1 大于o2，  
                         * 返回0 表示：o1和o2相等，  
                         * 返回正数表示：o1小于o2。  
                         */  

						@Override
						public int compare(HashMap<String, String> o1,
								HashMap<String, String> o2) {
							//按照评分进行降序排列  
							if(Double.parseDouble(o1.get("eleSorce")) > Double.parseDouble(o2.get("eleSorce"))) {
								return -1;
							}else if (Double.parseDouble(o1.get("eleSorce")) == Double.parseDouble(o2.get("eleSorce"))) {
								return 0;
							} else
							return 1;
						}  
                    });                  
                }
                i++;
            }    
  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
    }  
    

    
    public static void main(String[] args) {
    	CatchDouBan ju = CatchDouBan.getInstance();  
        ju.getDoubanReview("https://book.douban.com/tag/互联网?type=S"); 
        ju.getDoubanReview("https://book.douban.com/tag/编程?type=S"); 
        ju.getDoubanReview("https://book.douban.com/tag/算法?type=S"); 
        
        ju.createExcel();
  	}
      
}  