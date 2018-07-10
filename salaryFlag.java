import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


public class salaryFlag {
    public static DataFormatter dfor = new DataFormatter();
    public static void main(String[] args) throws IOException, InvalidFormatException {
        Properties prop = new Properties();
        String propFileName = "config.properties";
        InputStream inputStream = new FileInputStream(propFileName);
        prop.load(inputStream);
        System.out.println(inputStream);
        String SalaryMessagesFileLocation = prop.getProperty("SalaryMessagesFileLocation");
        String SalaryMessagesCopyFileLocation = prop.getProperty("SalaryMessagesCopyFileLocation");
        String SalaryPatternFileLocation = prop.getProperty("SalaryPatternFileLocation");
        File file = new File(SalaryMessagesCopyFileLocation);
        if (file.exists()) {
            file.delete();
        }
        Workbook Message = WorkbookFactory.create(new File(SalaryMessagesFileLocation));
        Sheet sheet1 = Message.getSheetAt(0);
        int rows = sheet1.getPhysicalNumberOfRows();
        Workbook Pattern = WorkbookFactory.create(new File(SalaryPatternFileLocation));
        Sheet sheet2 = Pattern.getSheetAt(0);
        Workbook Flag = new XSSFWorkbook();
        Sheet flag = Flag.createSheet("salaryFlags");
        int rows_pattern = sheet2.getPhysicalNumberOfRows();
        for(int i = 1; i < rows; ++i) {
            String message = dfor.formatCellValue(sheet1.getRow(i).getCell(1));
            System.out.println(message);
            Row create = flag.createRow(i);
            create.createCell(0).setCellValue(message);
            for(int j = 1; j < rows_pattern; ++j) {
              String pattern = dfor.formatCellValue(sheet2.getRow(j).getCell(0));
              Pattern text = java.util.regex.Pattern.compile(pattern);
              Matcher matcher = text.matcher(message);
                if (matcher.find()) {
                    create.createCell(1).setCellValue(1);
                    break;
            }
            create.createCell(1).setCellValue(0);
        }
        }
        FileOutputStream out = new FileOutputStream("./src/messagescopy.xlsx");
        Flag.write(out);
        out.close();
        Flag.close();
        Message.close();
        Pattern.close();
    }}


