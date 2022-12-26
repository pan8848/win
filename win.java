import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.time.Duration;
import java.time.LocalTime;

public class winver4 {

    public final static String INDEXTARGET = ".win.index";
    public final static String INFILE = "c://win/Wout.xlsx";

    public static void main(String[] args) {

        LocalTime starttime = LocalTime.now();
        //totol cnt
        int cnt=0;
        //單次寫入次數
        int singlecnt=1000;

        try {
            //獲取資料流
            FileInputStream inf = new FileInputStream("C://win//W2.xlsx");

            //取得活頁薄
            XSSFWorkbook wb = new XSSFWorkbook(inf);

            //取得工作表
            XSSFSheet sheet = wb.getSheetAt(0);

            //自動更新物件
            FormulaEvaluator formatwb =  new XSSFFormulaEvaluator((XSSFWorkbook) wb);

            // 取最後一行的行數
            int rowlen = sheet.getLastRowNum();
            //String sheetname= String.valueOf(sheet.getLastRowNum());

            ///起始行
            int rowstart = 1;

            //number的個數
            int q=39;

            //start number
            int istart = 1;

            //設定控制number個數
            int qcnt=5;

            //寫入檔設定
            XSSFWorkbook wbadd = wbWrite();
            XSSFSheet sheetadd = null;

            if(wbadd.getSheet(String.valueOf(qcnt+"star-"+rowlen)) != null){
                sheetadd = wbadd.getSheet(String.valueOf(qcnt+"star-"+rowlen));
            } else {
                sheetadd = wbadd.createSheet(String.valueOf(qcnt+"star-"+rowlen));
                //title
                String[] title = {"cnt","n1","n2","n3","n4","n5","nsolongshow","1star","2star","3star","nsumshow","nownoshowcnt"};
                XSSFRow toprow = sheetadd.createRow(0);
                XSSFCell topcell=null;
                for(int i=0;i<=title.length-1;i++) {
                    topcell = toprow.createCell(i);
                    topcell.setCellValue(title[i]);
                }
            }
            //XSSFRow rowadd = sheetadd.createRow((short)sheetadd.getLastRowNum());
            cnt = sheetadd.getLastRowNum()+1;
            System.out.println("cnt start:"+cnt);




            for(int i=istart;i<=q;i++) {

                sheet.getRow(1).getCell(16).setCellValue(i<10 ? "0"+i:String.valueOf(i));
                //寫入行值


                //刷新公式更新值
                //  for(int icnt =rowstart ; icnt<= rowlen ; icnt++) {
                //rown1=sheet.getRow(icnt);
                //celln1=rown1.getCell(16);
                // formatwb.evaluateFormulaCell(celln1);
                //        formatwb.notifyUpdateCell(sheet.getRow(icnt).getCell(16));
                //   System.out.println(celln1+"      "+formatwb.evaluate(celln1).formatAsString());
                //     }

                for(int j=i+1;j<=q;j++) {

                    sheet.getRow(1).getCell(17).setCellValue( j<10 ? "0"+j : String.valueOf(j));

                    for(int jcnt=rowstart ; jcnt<=rowlen ; jcnt++ ){
                        formatwb.notifyUpdateCell(sheet.getRow(jcnt).getCell(16));
                        formatwb.notifyUpdateCell(sheet.getRow(jcnt).getCell(17));
                        //  System.out.println(celln1+"      "+formatwb.evaluate(celln1).formatAsString());
                    }
                    for(int k=j+1;k<=q;k++) {

                        if(qcnt > 2) {

                            sheet.getRow(1).getCell(18).setCellValue(k < 10 ? "0" + k : String.valueOf(k));

                            for (int kcnt = rowstart; kcnt <= rowlen; kcnt++) {
                                formatwb.notifyUpdateCell(sheet.getRow(kcnt).getCell(18));
                                //   System.out.println(celln1+"      "+formatwb.evaluate(celln1).formatAsString());
                            }
                        }

                        for(int v=k+1;v<=q;v++) {

                            if(qcnt > 3 ) {

                                sheet.getRow(1).getCell(19).setCellValue(v < 10 ? "0" + v : String.valueOf(v));

                                for (int kcnt = rowstart; kcnt <= rowlen; kcnt++) {
                                    formatwb.notifyUpdateCell(sheet.getRow(kcnt).getCell(19));
                                    //   System.out.println(celln1+"      "+formatwb.evaluate(celln1).formatAsString());
                                }
                            }
                            for (int m = v+1; m <q ; m++) {
                                if(qcnt > 4 ) {

                                    sheet.getRow(1).getCell(20).setCellValue(m < 10 ? "0" + m : String.valueOf(m));

                                    for (int kcnt = rowstart; kcnt <= rowlen; kcnt++) {
                                        formatwb.notifyUpdateCell(sheet.getRow(kcnt).getCell(20));
                                        //   System.out.println(celln1+"      "+formatwb.evaluate(celln1).formatAsString());
                                    }
                                }else{
                                    m=99;
                                    if (qcnt<4) {
                                        v = 99;
                                        if (qcnt < 3)
                                            k = 99;
                                    }
                                }

                                //寫入值.讀取最後的ROW值
                                writeData(sheetadd, sheet, formatwb, cnt, qcnt);

                                cnt++;
                                System.out.println("k: " + k + "  v: " + v + "  m: " + m  + " cnt :" + cnt + " to single time : " + LocalTime.now());
                                if ((cnt % singlecnt == 0)) {
                                    savaExcel(wbadd);
                                    wbadd = wbWrite();
                                    sheetadd = wbadd.getSheet(String.valueOf(qcnt + "star-" + rowlen));

                                    System.out.println(" save" + singlecnt + " recoed... in: " + (cnt / singlecnt) + " times ,Total singlecnt :" + singlecnt + " starttime: " + starttime + " to single time : " + Duration.between(starttime, LocalTime.now()).toMinutes() + " Minutes");
                                    //清空所有更新值
                                    formatwb.clearAllCachedResultValues();
                                    //重新更新公式值
                                    for (int jcnt = rowstart; jcnt <= rowlen; jcnt++) {
                                        // formatwb.evaluateFormulaCell(celln1); '用在單次'
                                        formatwb.notifyUpdateCell(sheet.getRow(jcnt).getCell(16));
                                        formatwb.notifyUpdateCell(sheet.getRow(jcnt).getCell(17));
                                        formatwb.notifyUpdateCell(sheet.getRow(jcnt).getCell(18));
                                        formatwb.notifyUpdateCell(sheet.getRow(jcnt).getCell(19));
                                        //  System.out.println(celln1+"      "+formatwb.evaluate(celln1).formatAsString());
                                    }

                                }
                            }
                        }
                        System.out.println( " i : " + i +" j :" + j +" k :"+k+" cnt:"+cnt );
                    }

                    System.out.println("change j cnt:"+cnt);

                }
                if((i==q)) {
                    savaExcel(wbadd);
                }

            }

            inf.close();
            System.out.println("YES WIN!");
            System.out.println("starttime: " + starttime + " to : "+ Duration.between(starttime,LocalTime.now()).toMinutes()+ " Minutes");
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }finally {


        }

    }
    protected static void writeData(XSSFSheet sheetadd , XSSFSheet sheet,FormulaEvaluator formatwb,int cnt,int qcnt){

        XSSFRow nrow = sheetadd.createRow((int) sheetadd.getLastRowNum() + 1);

        //XSSFCell ncell = nrow.createCell(1);
        //cnt
        nrow.createCell(0).setCellValue(cnt);

        for (int i = 1; i <= qcnt; i++) {

            //n1~n-qcnt,Excel Q行
            nrow.createCell(i).setCellValue(sheet.getRow(1).getCell(15+i).getStringCellValue());
        }
        //nlongshowcnt
        //set cell
        int startCell = 6;
        //ncell = nrow.createCell(5);
        nrow.createCell(startCell).setCellValue(formatwb.evaluate(sheet.getRow(2).getCell(25)).getNumberValue());
        //s1
        //ncell = nrow.createCell(6);
        nrow.createCell(startCell+1).setCellValue(formatwb.evaluate(sheet.getRow(35).getCell(25)).getNumberValue());
        //s2
        //ncell = nrow.createCell(7);
        nrow.createCell(startCell+2).setCellValue(formatwb.evaluate(sheet.getRow(36).getCell(25)).getNumberValue());
        //s3
        //ncell = nrow.createCell(8);
        nrow.createCell(startCell+3).setCellValue(formatwb.evaluate(sheet.getRow(37).getCell(25)).getNumberValue());
        //nsumshow
        nrow.createCell(startCell+4).setCellValue(formatwb.evaluate(sheet.getRow(1).getCell(25)).getNumberValue());
        //nownoshowcnt
        nrow.createCell(startCell+5).setCellValue(formatwb.evaluate(sheet.getRow(2).getCell(26)).formatAsString());

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
        System.out.println(f.exists());
        try {
            if(!f.exists())
                f.createNewFile();
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
        System.out.println("wb:"+wb);
        return wb;

    }
}


