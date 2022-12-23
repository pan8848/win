package com.teamsart.spring.controller.win;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Date;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import javax.servlet.http.HttpServletRequest;
import javax.sql.rowset.JoinRowSet;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.eval.forked.ForkedEvaluator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;

import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.RequestMapping;

import com.teamsart.spring.controller.SpringController;

@Controller
@RequestMapping(value = "/win")
public class WinController extends SpringController {

	public final static String INDEXTARGET = ".win.index";
	public final static String INFILE = "c://win/wout.xlsx";
	
@RequestMapping (value = "/index")
public String index (HttpServletRequest req,ModelMap model) throws FileNotFoundException, InvalidFormatException {
    LocalTime starttime = LocalTime.now();
    LocalTime singletime;
    //totol cnt
    int cnt=0;
    //
    int singlecnt=0;
    try {
        FileInputStream inf = new FileInputStream("C:\\win\\W2.xlsm");
        //取得活頁薄
        XSSFWorkbook wb = new XSSFWorkbook(inf);
        //取得工作表
        XSSFSheet sheet = wb.getSheetAt(0);
        //自動更新物件

        FormulaEvaluator formatwb =  new XSSFFormulaEvaluator((XSSFWorkbook) wb);
        // 取最後一行的行數
        int rowlen = sheet.getLastRowNum();
        ///起始行
        int rowstart = 1;
        //number的個數
        int q=39;


        //設定1N

        XSSFRow rown1;
        XSSFRow rown2;
        XSSFCell celln1; //n1
        XSSFCell celln2; //n2
        CellValue cellvaluen1;
        Date s1 = new Date();
        
        //寫入檔設定
        XSSFWorkbook wbadd = wbWrite();
        XSSFSheet sheetadd = wbadd.getSheet(String.valueOf(rowlen));
        //XSSFRow rowadd = sheetadd.createRow((short)sheetadd.getLastRowNum());
        cnt = (short)sheetadd.getLastRowNum();
        System.out.println("cnt start:"+cnt);
        
        
/*         //title
        String[] title = {"","n1","n2","n3","n4","nsolongshow","1star","2star","3star","nsumshow"};
        XSSFRow toprow = sheetadd.createRow(0);
        XSSFCell topcell=null;
        for(int i=1;i<=title.length-1;i++) {
            topcell = toprow.createCell(i);
            topcell.setCellValue(title[i]);
        }
        */
        for(int i=5;i<=5;i++) {
        	 	
            
            rown1= sheet.getRow(1);
            celln1 = rown1.getCell(16);
            celln1.setCellValue(i<10 ? "0"+i:String.valueOf(i));
            //寫入行值


            //刷新公式更新值
            for(int icnt =rowstart ; icnt<= rowlen ; icnt++) {
                rown1=sheet.getRow(icnt);
                celln1=rown1.getCell(16);
                // formatwb.evaluateFormulaCell(celln1);
                formatwb.notifyUpdateCell(celln1);
                //   System.out.println(celln1+"      "+formatwb.evaluate(celln1).formatAsString());
            }

            for(int j=i+1;j<=q;j++) {

                sheet.getRow(1).getCell(17).setCellValue( j<10 ? "0"+j : String.valueOf(j));

                for(int jcnt=rowstart ; jcnt<=rowlen ; jcnt++ ){
                    rown1=sheet.getRow(jcnt);
                    celln1=rown1.getCell(17);
                    // formatwb.evaluateFormulaCell(celln1);
                    formatwb.notifyUpdateCell(celln1);
                    //  System.out.println(celln1+"      "+formatwb.evaluate(celln1).formatAsString());
                }
                for(int k=j+1;k<=q;k++) {

                    sheet.getRow(1).getCell(18).setCellValue( k<10 ? "0"+k : String.valueOf(k));

                    for(int kcnt=rowstart; kcnt <= rowlen; kcnt++){
                        rown1=sheet.getRow(kcnt);
                        celln1=rown1.getCell(18);
                        // formatwb.evaluateFormulaCell(celln1);
                        formatwb.notifyUpdateCell(celln1);
                        //   System.out.println(celln1+"      "+formatwb.evaluate(celln1).formatAsString());
                    }
                    for(int v=k+1;v<=q;v++) {

                        sheet.getRow(1).getCell(19).setCellValue( v<10 ? "0"+v : String.valueOf(v));

                        for(int kcnt=rowstart; kcnt <= rowlen; kcnt++){
                            rown1=sheet.getRow(kcnt);
                            celln1=rown1.getCell(19);
                            // formatwb.evaluateFormulaCell(celln1);
                            formatwb.notifyUpdateCell(celln1);
                            //   System.out.println(celln1+"      "+formatwb.evaluate(celln1).formatAsString());
                        }
                        //寫入值.讀取最後的ROW值
                        XSSFRow nrow = sheetadd.createRow((int) sheetadd.getLastRowNum() + 1);

                        XSSFCell ncell = nrow.createCell(1);
                        //cnt
                        nrow.createCell(0).setCellValue(cnt);
                        //n1
                        ncell.setCellValue(sheet.getRow(1).getCell(16).getStringCellValue());
                        //n2
                        ncell = nrow.createCell(2);
                        ncell.setCellValue(sheet.getRow(1).getCell(17).getStringCellValue());
                        //n3
                        ncell = nrow.createCell(3);
                        ncell.setCellValue(sheet.getRow(1).getCell(18).getStringCellValue());
                        //n4
                        ncell = nrow.createCell(4);
                        ncell.setCellValue(sheet.getRow(1).getCell(19).getStringCellValue());
                        //nlongshowcnt
                        ncell = nrow.createCell(5);
                        ncell.setCellValue(formatwb.evaluate(sheet.getRow(2).getCell(25)).getNumberValue());
                        //s1
                        ncell = nrow.createCell(6);
                        ncell.setCellValue(formatwb.evaluate(sheet.getRow(35).getCell(25)).getNumberValue());
                        //s2
                        ncell = nrow.createCell(7);
                        ncell.setCellValue(formatwb.evaluate(sheet.getRow(36).getCell(25)).getNumberValue());
                        //s3
                        ncell = nrow.createCell(8);
                        ncell.setCellValue(formatwb.evaluate(sheet.getRow(37).getCell(25)).getNumberValue());
                        cnt++;
                        singlecnt++;
                        if((singlecnt % 1000 ==0)) {
                            savaExcel(wbadd);
                            wbadd = wbWrite();
                            sheetadd = wbadd.getSheet(String.valueOf(rowlen));
                            singletime = LocalTime.now();
                            System.out.println( " singlecnt :" +singlecnt + " starttime: " + starttime + " to single time : "+ singletime);
                            
                            }
                    }
                    System.out.println( " i : " + i +" j :" + j +" k :"+k+" cnt:"+cnt + " singlecnt :" +singlecnt );
                }
                
                System.out.println("cnt:"+cnt + " singlecnt :" +singlecnt);

            }
            if((i==5)) {
            savaExcel(wbadd);
            }
        }

        inf.close();
        System.out.println("YES WIN!");
        System.out.println("starttime: " + starttime + " to : "+ new Date());
    } catch (Exception e) {
        // TODO Auto-generated catch block
        e.printStackTrace();
    }finally {


    }

	model.addAttribute("cnt",singlecnt);
	return INDEXTARGET;
}

protected static void savaExcel(XSSFWorkbook wb){
    FileOutputStream fileOut = null;
    try {
        fileOut = new FileOutputStream(INFILE);
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
    } catch (FileNotFoundException e) {
        e.printStackTrace();
    } catch (IOException e) {
        e.printStackTrace();
    } finally {
    	if(fileOut!=null) {
    		try {
				fileOut.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
    	}
    }
}

protected static XSSFWorkbook wbWrite(){
XSSFWorkbook wb = null;
FileInputStream fis = null;
File f = new File(INFILE);

    try {
        if (f!=null) {
            fis = new FileInputStream(f);
            wb = new XSSFWorkbook(fis);
        }
    } catch (Exception e) {
        return null;
    }finally{
        if(fis!=null){
            try {
                fis.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
    return wb;

}
public static XSSFWorkbook getWorkBook(String filePath) {
	XSSFWorkbook workbook =  null;
	try {
		File fileXlsxPath = new File(filePath);
		fileXlsxPath.createNewFile();
		FileOutputStream outs = new FileOutputStream(fileXlsxPath);
		//BufferedOutputStream outPutStream = new BufferedOutputStream(FileUtils.openOutputStream(fileXlsxPath));
		workbook = new XSSFWorkbook();
		workbook.createSheet("測試");
		workbook.write(outs);
		
		outs.close();
	} catch (Exception e) {
		e.printStackTrace();
	}
	return workbook;
}



private static void updatecnt(XSSFWorkbook wb,XSSFSheet sheet,int i,int rowstart) {
//	XSSFRow r=s.getRow(row);
//	XSSFCell c =null;
FormulaEvaluator formatwb = null;
	formatwb = new XSSFFormulaEvaluator((XSSFWorkbook)wb);
//	for(int i = r.getFirstCellNum();i<r.getLastCellNum();i++)
//		c=r.getCell(i);
//	if(c.getCellType()==Cell.CELL_TYPE_FORMULA)
//		eval.evaluateFormulaCell(c);

	XSSFRow rown1;
	XSSFCell celln1;
	CellValue cellvaluen1;
	 for(int j =rowstart;j<=10;j++) {
		 rown1=sheet.getRow(j);
		 celln1=rown1.getCell(16);
		// formatwb.evaluateFormulaCell(celln1);
		 formatwb.notifyUpdateCell(celln1);
		  	System.out.println(celln1+"      "+formatwb.evaluate(celln1).formatAsString());
	}
	 
}

}
package com.teamsart.spring.controller.win;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Date;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import javax.servlet.http.HttpServletRequest;
import javax.sql.rowset.JoinRowSet;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.eval.forked.ForkedEvaluator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;

import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.RequestMapping;

import com.teamsart.spring.controller.SpringController;

@Controller
@RequestMapping(value = "/win")
public class WinController extends SpringController {

	public final static String INDEXTARGET = ".win.index";
	public final static String INFILE = "c://win/wout.xlsx";
	
@RequestMapping (value = "/index")
public String index (HttpServletRequest req,ModelMap model) throws FileNotFoundException, InvalidFormatException {
    LocalTime starttime = LocalTime.now();
    LocalTime singletime;
    //totol cnt
    int cnt=0;
    //
    int singlecnt=0;
    try {
        FileInputStream inf = new FileInputStream("C:\\win\\W2.xlsm");
        //取得活頁薄
        XSSFWorkbook wb = new XSSFWorkbook(inf);
        //取得工作表
        XSSFSheet sheet = wb.getSheetAt(0);
        //自動更新物件

        FormulaEvaluator formatwb =  new XSSFFormulaEvaluator((XSSFWorkbook) wb);
        // 取最後一行的行數
        int rowlen = sheet.getLastRowNum();
        ///起始行
        int rowstart = 1;
        //number的個數
        int q=39;


        //設定1N

        XSSFRow rown1;
        XSSFRow rown2;
        XSSFCell celln1; //n1
        XSSFCell celln2; //n2
        CellValue cellvaluen1;
        Date s1 = new Date();
        
        //寫入檔設定
        XSSFWorkbook wbadd = wbWrite();
        XSSFSheet sheetadd = wbadd.getSheet(String.valueOf(rowlen));
        //XSSFRow rowadd = sheetadd.createRow((short)sheetadd.getLastRowNum());
        cnt = (short)sheetadd.getLastRowNum();
        System.out.println("cnt start:"+cnt);
        
        
/*         //title
        String[] title = {"","n1","n2","n3","n4","nsolongshow","1star","2star","3star","nsumshow"};
        XSSFRow toprow = sheetadd.createRow(0);
        XSSFCell topcell=null;
        for(int i=1;i<=title.length-1;i++) {
            topcell = toprow.createCell(i);
            topcell.setCellValue(title[i]);
        }
        */
        for(int i=5;i<=5;i++) {
        	 	
            
            rown1= sheet.getRow(1);
            celln1 = rown1.getCell(16);
            celln1.setCellValue(i<10 ? "0"+i:String.valueOf(i));
            //寫入行值


            //刷新公式更新值
            for(int icnt =rowstart ; icnt<= rowlen ; icnt++) {
                rown1=sheet.getRow(icnt);
                celln1=rown1.getCell(16);
                // formatwb.evaluateFormulaCell(celln1);
                formatwb.notifyUpdateCell(celln1);
                //   System.out.println(celln1+"      "+formatwb.evaluate(celln1).formatAsString());
            }

            for(int j=i+1;j<=q;j++) {

                sheet.getRow(1).getCell(17).setCellValue( j<10 ? "0"+j : String.valueOf(j));

                for(int jcnt=rowstart ; jcnt<=rowlen ; jcnt++ ){
                    rown1=sheet.getRow(jcnt);
                    celln1=rown1.getCell(17);
                    // formatwb.evaluateFormulaCell(celln1);
                    formatwb.notifyUpdateCell(celln1);
                    //  System.out.println(celln1+"      "+formatwb.evaluate(celln1).formatAsString());
                }
                for(int k=j+1;k<=q;k++) {

                    sheet.getRow(1).getCell(18).setCellValue( k<10 ? "0"+k : String.valueOf(k));

                    for(int kcnt=rowstart; kcnt <= rowlen; kcnt++){
                        rown1=sheet.getRow(kcnt);
                        celln1=rown1.getCell(18);
                        // formatwb.evaluateFormulaCell(celln1);
                        formatwb.notifyUpdateCell(celln1);
                        //   System.out.println(celln1+"      "+formatwb.evaluate(celln1).formatAsString());
                    }
                    for(int v=k+1;v<=q;v++) {

                        sheet.getRow(1).getCell(19).setCellValue( v<10 ? "0"+v : String.valueOf(v));

                        for(int kcnt=rowstart; kcnt <= rowlen; kcnt++){
                            rown1=sheet.getRow(kcnt);
                            celln1=rown1.getCell(19);
                            // formatwb.evaluateFormulaCell(celln1);
                            formatwb.notifyUpdateCell(celln1);
                            //   System.out.println(celln1+"      "+formatwb.evaluate(celln1).formatAsString());
                        }
                        //寫入值.讀取最後的ROW值
                        XSSFRow nrow = sheetadd.createRow((int) sheetadd.getLastRowNum() + 1);

                        XSSFCell ncell = nrow.createCell(1);
                        //cnt
                        nrow.createCell(0).setCellValue(cnt);
                        //n1
                        ncell.setCellValue(sheet.getRow(1).getCell(16).getStringCellValue());
                        //n2
                        ncell = nrow.createCell(2);
                        ncell.setCellValue(sheet.getRow(1).getCell(17).getStringCellValue());
                        //n3
                        ncell = nrow.createCell(3);
                        ncell.setCellValue(sheet.getRow(1).getCell(18).getStringCellValue());
                        //n4
                        ncell = nrow.createCell(4);
                        ncell.setCellValue(sheet.getRow(1).getCell(19).getStringCellValue());
                        //nlongshowcnt
                        ncell = nrow.createCell(5);
                        ncell.setCellValue(formatwb.evaluate(sheet.getRow(2).getCell(25)).getNumberValue());
                        //s1
                        ncell = nrow.createCell(6);
                        ncell.setCellValue(formatwb.evaluate(sheet.getRow(35).getCell(25)).getNumberValue());
                        //s2
                        ncell = nrow.createCell(7);
                        ncell.setCellValue(formatwb.evaluate(sheet.getRow(36).getCell(25)).getNumberValue());
                        //s3
                        ncell = nrow.createCell(8);
                        ncell.setCellValue(formatwb.evaluate(sheet.getRow(37).getCell(25)).getNumberValue());
                        cnt++;
                        singlecnt++;
                        if((singlecnt % 1000 ==0)) {
                            savaExcel(wbadd);
                            wbadd = wbWrite();
                            sheetadd = wbadd.getSheet(String.valueOf(rowlen));
                            singletime = LocalTime.now();
                            System.out.println( " singlecnt :" +singlecnt + " starttime: " + starttime + " to single time : "+ singletime);
                            
                            }
                    }
                    System.out.println( " i : " + i +" j :" + j +" k :"+k+" cnt:"+cnt + " singlecnt :" +singlecnt );
                }
                
                System.out.println("cnt:"+cnt + " singlecnt :" +singlecnt);

            }
            if((i==5)) {
            savaExcel(wbadd);
            }
        }

        inf.close();
        System.out.println("YES WIN!");
        System.out.println("starttime: " + starttime + " to : "+ new Date());
    } catch (Exception e) {
        // TODO Auto-generated catch block
        e.printStackTrace();
    }finally {


    }

	model.addAttribute("cnt",singlecnt);
	return INDEXTARGET;
}

protected static void savaExcel(XSSFWorkbook wb){
    FileOutputStream fileOut = null;
    try {
        fileOut = new FileOutputStream(INFILE);
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
    } catch (FileNotFoundException e) {
        e.printStackTrace();
    } catch (IOException e) {
        e.printStackTrace();
    } finally {
    	if(fileOut!=null) {
    		try {
				fileOut.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
    	}
    }
}

protected static XSSFWorkbook wbWrite(){
XSSFWorkbook wb = null;
FileInputStream fis = null;
File f = new File(INFILE);

    try {
        if (f!=null) {
            fis = new FileInputStream(f);
            wb = new XSSFWorkbook(fis);
        }
    } catch (Exception e) {
        return null;
    }finally{
        if(fis!=null){
            try {
                fis.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
    return wb;

}
public static XSSFWorkbook getWorkBook(String filePath) {
	XSSFWorkbook workbook =  null;
	try {
		File fileXlsxPath = new File(filePath);
		fileXlsxPath.createNewFile();
		FileOutputStream outs = new FileOutputStream(fileXlsxPath);
		//BufferedOutputStream outPutStream = new BufferedOutputStream(FileUtils.openOutputStream(fileXlsxPath));
		workbook = new XSSFWorkbook();
		workbook.createSheet("測試");
		workbook.write(outs);
		
		outs.close();
	} catch (Exception e) {
		e.printStackTrace();
	}
	return workbook;
}



private static void updatecnt(XSSFWorkbook wb,XSSFSheet sheet,int i,int rowstart) {
//	XSSFRow r=s.getRow(row);
//	XSSFCell c =null;
FormulaEvaluator formatwb = null;
	formatwb = new XSSFFormulaEvaluator((XSSFWorkbook)wb);
//	for(int i = r.getFirstCellNum();i<r.getLastCellNum();i++)
//		c=r.getCell(i);
//	if(c.getCellType()==Cell.CELL_TYPE_FORMULA)
//		eval.evaluateFormulaCell(c);

	XSSFRow rown1;
	XSSFCell celln1;
	CellValue cellvaluen1;
	 for(int j =rowstart;j<=10;j++) {
		 rown1=sheet.getRow(j);
		 celln1=rown1.getCell(16);
		// formatwb.evaluateFormulaCell(celln1);
		 formatwb.notifyUpdateCell(celln1);
		  	System.out.println(celln1+"      "+formatwb.evaluate(celln1).formatAsString());
	}
	 
}

}
