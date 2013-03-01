import java.io.File;
import java.io.IOException;

import com.sun.medialib.mlib.Image;
import com.sun.tools.javac.util.Log;

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
 
      
    /**生成一个Excel文件  
     * @param fileName  要生成的Excel文件名  
     * @throws WriteException 
     * @throws RowsExceededException 
     */ 
    public static void writeExcel(String fileName) throws RowsExceededException, WriteException{
        
        Workbook wb = null;
        WritableWorkbook wwb = null;
        try {
            wb = Workbook.getWorkbook(new File(fileName));
            wwb = Workbook.createWorkbook(new File(fileName), wb);
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        WritableSheet[] sheets = wwb.getSheets();
        if (sheets == null || sheets.length != 2) {
			return;
		}
        
        Sheet sourceDataSheet = sheets[0];
        
        int rows = sourceDataSheet.getRows(); 
        for (int i = 1; i < rows; i++) {
            Cell[] cells = sourceDataSheet.getRow(i);
        	wwb.copySheet("template", cells[0].getContents(), i + 2);
        	
        	String size = cells[1].getContents() + "0" + " X " + cells[2].getContents() + " X " + cells[3].getContents();
        	
        	String grossWgt = (int)(Double.valueOf(cells[4].getContents()) * 1000) + 110 + "KG";
        	String netWgt = (int)(Double.valueOf(cells[4].getContents()) * 1000) + "KG";
			
			String identification = cells[0].getContents();
			String coilNo = cells[6].getContents();
			
			String date = cells[7].getContents();
			
			WritableSheet  writeSheet = wwb.getSheet(i + 1);
			
			Label sizeLabel1 = new Label(1 , 5, size); 
			writeSheet.addCell(sizeLabel1); 
			Label sizeLabel2 = new Label(1 , 15, size); 
			writeSheet.addCell(sizeLabel2);
			
			
			Label grossWgtLabel1 = new Label(6, 5, grossWgt);
			writeSheet.addCell(grossWgtLabel1); 
			Label grossWgtLabel2 = new Label(6, 15, grossWgt);
			writeSheet.addCell(grossWgtLabel2); 
			
			Label netWgtLabel1 = new Label(9, 5, netWgt);
            writeSheet.addCell(netWgtLabel1); 
            Label netWgtLabel2 = new Label(9, 15, netWgt);
            writeSheet.addCell(netWgtLabel2); 
            
            Label identificationLabel1 = new Label(1, 6, identification);
            writeSheet.addCell(identificationLabel1); 
            Label identificationLabel2 = new Label(1, 16, identification);
            writeSheet.addCell(identificationLabel2); 
            
            Label coilLabel1 = new Label(6, 6, coilNo);
            writeSheet.addCell(coilLabel1); 
            Label coilLabel2 = new Label(6, 16, coilNo);
            writeSheet.addCell(coilLabel2); 
            
            Label dateLabel1 = new Label(4, 8, date);
            writeSheet.addCell(dateLabel1); 
            Label dateLabel2 = new Label(4, 18, date);
            writeSheet.addCell(dateLabel2); 
            
            String barCodePicPath = "/Users/Jamie/Documents/barcode.png";
            File barCodeImgFile = new File(barCodePicPath);
            insertImg(writeSheet, 5.5, 7.15, 3.8, 1.8,barCodeImgFile);
            insertImg(writeSheet, 5.5, 17.15, 3.8, 1.8,barCodeImgFile);
            
            String headPicPath = "/Users/Jamie/Documents/head.png";
            File headImgFile = new File(headPicPath);
            insertImg(writeSheet, 0.2, 0.1, 9, 0.8,headImgFile);
            insertImg(writeSheet, 0.2, 10.1, 9, 0.8,headImgFile);

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
        String filePath = "/Users/Jamie/Documents/label.xls";
        
        try {
            writeExcel(filePath);
        } catch (RowsExceededException e) {
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
}
