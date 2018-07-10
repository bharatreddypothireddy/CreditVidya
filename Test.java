import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;


import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import static java.lang.Float.*;

public class Test {
    public static DataFormatter dfor = new DataFormatter();
    public static void main(String[] args) throws IOException, InvalidFormatException {
        Workbook work =  WorkbookFactory.create(new File("./src/preethi.xlsx"));
        Sheet sheet = work.getSheetAt(0);
        Sheet deleted = work.createSheet("deleted");
        int f = sheet.getPhysicalNumberOfRows();
        int q =1;
        for(int i =1;i<f-2;i++){
            int j = i+1;
            String val = dfor.formatCellValue(sheet.getRow(i).getCell(0));
            String nex = dfor.formatCellValue(sheet.getRow(j).getCell(0));
            int one = Integer.parseInt(dfor.formatCellValue(sheet.getRow(i).getCell(2)));
            int two = Integer.parseInt(dfor.formatCellValue(sheet.getRow(j).getCell(2)));
            if  (val.equals("debit")) {
               while(((two - one) < 3) && ((two - one) > -1)){
                  if(nex.equals("credit")){
                      String abc = dfor.formatCellValue(sheet.getRow(i).getCell(1)).replaceAll(",","");
                      float i1 = parseFloat(abc);
                      String xyz = dfor.formatCellValue(sheet.getRow(j).getCell(1)).replaceAll(",","");
                      float j1 = parseFloat(xyz);
                        if(((j1 - i1) < 60) && ((j1 - i1) > -60)){
                          int y =i;
                          for(int p =0;p<2;p++) {
                            Row newer = deleted.createRow(q);
                            Row rower = sheet.getRow(y);
                            int a = 0;
                            for(Cell old :rower){
                                newer.createCell(a).setCellValue(dfor.formatCellValue(old));
                                a++;
                            }
                            y=j;
                            q++;
                        }
                        sheet.removeRow(sheet.getRow(i));
                        sheet.removeRow(sheet.getRow(j));
                        if(j<f-1) {
                            sheet.shiftRows(j+1,sheet.getLastRowNum(),-1);
                        }
                        sheet.shiftRows(i+1,sheet.getLastRowNum(),-1);
                        f=f-2;
                        i=i-1;
                        break;
                    }}
                  j=j+1;
                  if(j<f) {
                     two = Integer.parseInt(dfor.formatCellValue(sheet.getRow(j).getCell(2)));
                     nex = dfor.formatCellValue(sheet.getRow(j).getCell(0));
                             }
                  else {
                      break;
                        }}}
           else if(val.equals("credit")){
               while(((two - one) < 3) && ((two - one) > -1)){
                   if(nex.equals("debit")){
                       String abc = dfor.formatCellValue(sheet.getRow(i).getCell(1)).replaceAll(",","");
                       float i1 = parseFloat(abc);
                       String xyz = dfor.formatCellValue(sheet.getRow(j).getCell(1)).replaceAll(",","");
                       float j1 = parseFloat(xyz);if(((j1 - i1) < 60) && ((j1 - i1) > -60)){
                           int y =i;
                           for(int p =0;p<2;p++) {
                               Row newer = deleted.createRow(q);
                               Row rower = sheet.getRow(y);
                               int a = 0;
                               for(Cell old :rower){
                                   newer.createCell(a).setCellValue(dfor.formatCellValue(old));
                                   a++;
                               }
                               y=j;
                               q++;
                           }
                           sheet.removeRow(sheet.getRow(i));
                           sheet.removeRow(sheet.getRow(j));
                           if(j<f-1)
                               sheet.shiftRows(j+1,sheet.getLastRowNum(),-1);
                           sheet.shiftRows(i+1,sheet.getLastRowNum(),-1);
                           f=f-2;
                           i=i-1;
                           break;
                       }}
                   j=j+1;
                   if(j<f) {
                       two = Integer.parseInt(dfor.formatCellValue(sheet.getRow(j).getCell(2)));
                       nex = dfor.formatCellValue(sheet.getRow(j).getCell(0));
                   }
                   else {
                       break;
                   }}}}
        FileOutputStream pathers = new FileOutputStream("./src/outputpreethi.xlsx");
        work.write(pathers);
        pathers.close();
    }}
