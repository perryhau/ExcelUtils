import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableImage;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
 
public class ExcelUtils {  
 
    /**读取Excel文件的内容  
     * @param file  待读取的文件  
     * @return  
     */ 
    public static String readExcel(File file){  
        StringBuffer sb = new StringBuffer();  
          
        Workbook wb = null;  
        try {  
            //构造Workbook（工作薄）对象  
            wb=Workbook.getWorkbook(file);  
        } catch (BiffException e) {  
            e.printStackTrace();  
        } catch (IOException e) {  
            e.printStackTrace();  
        }  
          
        if(wb==null)  
            return null;  
          
        //获得了Workbook对象之后，就可以通过它得到Sheet（工作表）对象了  
        Sheet[] sheet = wb.getSheets();
          
        if(sheet!=null&&sheet.length>0){  
            //对每个工作表进行循环  
            for(int i=0;i<sheet.length;i++){  
                //得到当前工作表的行数  
                int rowNum = sheet[i].getRows();  
                for(int j=0;j<rowNum;j++){  
                    //得到当前行的所有单元格  
                    Cell[] cells = sheet[i].getRow(j);  
                    if(cells!=null&&cells.length>0){  
                        //对每个单元格进行循环  
                        for(int k=0;k<cells.length;k++){  
                            //读取当前单元格的值  
                            String cellValue = cells[k].getContents();  
                            sb.append(cellValue+"\t");  
                        }  
                    }  
                    sb.append("\r\n");  
                }  
                sb.append("\r\n");  
            }  
        }  
        //最后关闭资源，释放内存  
        wb.close();  
        System.out.println("313");  
        return sb.toString();  
    }  
    /**生成一个Excel文件  
     * @param fileName  要生成的Excel文件名  
     * @throws WriteException 
     * @throws RowsExceededException 
     */ 
    public static void writeExcel(String fileName) throws RowsExceededException, WriteException{  
        WritableWorkbook wwb = null;  
        File file = new File(fileName);
        try {  
            //首先要使用Workbook类的工厂方法创建一个可写入的工作薄(Workbook)对象  
            wwb = Workbook.createWorkbook(file);  
        } catch (IOException e) {  
            e.printStackTrace();  
        }
        
        Workbook wb = null;  
        try {  
            //构造Workbook（工作薄）对象  
            wb = Workbook.getWorkbook(file);  
        } catch (BiffException e) {  
            e.printStackTrace();  
        } catch (IOException e) {  
            e.printStackTrace();  
        }  
          
        if(wb==null || wwb == null){
            return;
        }
          
        //获得了Workbook对象之后，就可以通过它得到Sheet（工作表）对象了  
        Sheet[] sheet = wb.getSheets();
        
        if (sheet == null || sheet.length < 2) {
			return;
		}
        
        Sheet sourceDataSheet = sheet[0];
        
        int rows = sourceDataSheet.getRows(); 
        for (int i = 0; i < rows; i++) {
        	Cell[] cells = sourceDataSheet.getRow(i);
        	
        	//TODO copy sheet from 1
        	wwb.copySheet(1, cells[1].getContents(), i + 2);
        	
        	String deastination = "BARRANQUILLA";
        	String consgnee = "TO WHOM IT";
        	String orderNoString = "13YCWG123";
        	String productName = "Cold - Rolled Steel Coil";
        	String specification = "ASTM A-424 TYPPE II";
        	String size = cells[2].getContents() + "0" + " X " + cells[3].getContents() + " X C";
        	
        	String grossWgt = Double.valueOf(cells[6].getContents()) * 100 + 110 + "KG";
			String netWgt = Double.valueOf(cells[6].getContents()) * 100 + "KG";
			
			String identification = cells[1].getContents();
			String coilNo = cells[8].getContents();
			
			String date = "20130215";
			WritableSheet  writeSheet = wwb.getSheet(i + 2);
			
			// Create a cell format for Times 16, bold and italic 
			WritableFont deastinationFont = new WritableFont(WritableFont.TIMES, 16, WritableFont.BOLD, true); 
			WritableCellFormat deastinationformat = new WritableCellFormat (deastinationFont); 
			Label deastinationLabel = new Label(1, 1, deastination, deastinationformat);
			writeSheet.addCell(deastinationLabel);
			
			WritableFont consgneeFront = new WritableFont(WritableFont.TIMES, 16, WritableFont.BOLD, true); 
            WritableCellFormat consgneeformat = new WritableCellFormat (consgneeFront); 
            Label consgneeLabel = new Label(1, 2, consgnee, consgneeformat);
            writeSheet.addCell(consgneeLabel);
		}
        
        try {  
            //从内存中写入文件中  
            wwb.write();  
            //关闭资源，释放内存  
            wwb.close();  
        } catch (IOException e) {  
            e.printStackTrace();  
        } catch (WriteException e) {  
            e.printStackTrace();  
        }
    }   
    /**搜索某一个文件中是否包含某个关键字  
     * @param file  待搜索的文件  
     * @param keyWord  要搜索的关键字  
     * @return  
     */ 
    public static boolean searchKeyWord(File file,String keyWord){  
        boolean res = false;  
          
        Workbook wb = null;  
        try {  
            //构造Workbook（工作薄）对象  
            wb=Workbook.getWorkbook(file);  
        } catch (BiffException e) {  
            return res;  
        } catch (IOException e) {  
            return res;  
        }  
          
        if(wb==null)  
            return res;  
          
        //获得了Workbook对象之后，就可以通过它得到Sheet（工作表）对象了  
        Sheet[] sheet = wb.getSheets();  
          
        boolean breakSheet = false;  
          
        if(sheet!=null&&sheet.length>0){  
            //对每个工作表进行循环  
            for(int i=0;i<sheet.length;i++){  
                if(breakSheet)  
                    break;  
                  
                //得到当前工作表的行数  
                int rowNum = sheet[i].getRows();  
                  
                boolean breakRow = false;  
                  
                for(int j=0;j<rowNum;j++){  
                    if(breakRow)  
                        break;  
                    //得到当前行的所有单元格  
                    Cell[] cells = sheet[i].getRow(j);  
                    if(cells!=null&&cells.length>0){  
                        boolean breakCell = false;  
                        //对每个单元格进行循环  
                        for(int k=0;k<cells.length;k++){  
                            if(breakCell)  
                                break;  
                            //读取当前单元格的值  
                            String cellValue = cells[k].getContents();  
                            if(cellValue==null)  
                                continue;  
                            if(cellValue.contains(keyWord)){  
                                res = true;  
                                breakCell = true;  
                                breakRow = true;  
                                breakSheet = true;  
                            }  
                        }  
                    }  
                }  
            }  
        }  
        //最后关闭资源，释放内存  
        wb.close();  
        System.out.print(res);  
        return res;  
    }  
    /**往Excel中插入图片  
     * @param dataSheet  待插入的工作表  
     * @param col 图片从该列开始  
     * @param row 图片从该行开始  
     * @param width 图片所占的列数  
     * @param height 图片所占的行数  
     * @param imgFile 要插入的图片文件  
     */ 
    public static void insertImg(WritableSheet dataSheet, double col, double row, double width,  
    		double height, File imgFile){  
        WritableImage img = new WritableImage(col, row, width, height, imgFile);  
        dataSheet.addImage(img);  
    }   
      
      
    public static void main(String[] args) {  
           
        
        String filePath = "/home/jamie/Downloads/test.xls";
        
        try {
            writeExcel(filePath);
        } catch (RowsExceededException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }
        
//        try {
//        	String filePath = "/Users/jamiemo/Documents/test.xls";
//        	String picPathString = "/Users/jamiemo/Documents/a.png";
//            Workbook wb = Workbook.getWorkbook(new File(filePath));
//            WritableWorkbook wwb = Workbook.createWorkbook(new File(filePath), wb);
//            WritableSheet[] sheets = wwb.getSheets();
//            for (WritableSheet sheet : sheets) {
//            	File imgFile = new File(picPathString);
//            	insertImg(sheet, 5.5, 7.15, 3.8, 1.8,imgFile);
//            	insertImg(sheet, 5.5, 17.15, 3.8, 1.8,imgFile);
//				
//			} 
//            wwb.write();
//            wwb.close();  
//        } catch (IOException e) {
//            e.printStackTrace();
//        } catch (WriteException e) {
//            e.printStackTrace();
//        } catch (BiffException e) {
//			e.printStackTrace();
//		}
    }
}
