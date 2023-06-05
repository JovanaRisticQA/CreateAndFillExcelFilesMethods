package createFakerExcelTable;

import com.github.javafaker.Faker;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

public class MethodsToPopulateExcelFileWithFaker {
    public static void main(String[] args) {

        String filePath = "MyDummyDocument.xlsx";
        createFakerExcelTableWithTwoForLoops(filePath);
        //createFakerTableWithTwoDimensionalObjectArray(filePath);

    }

    public static void createFakerExcelTableWithTwoForLoops(String filePath) {

        //The try-with-resources statement automatically closes all the resources at the end of the statement
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             FileOutputStream outputStream = new FileOutputStream(filePath)) {

            XSSFSheet sheet = workbook.createSheet("FakerData");
            XSSFRow row;

            //Create an Excel table with 100 rows and 10 columns
            for (int i = 0; i < 100; i++) {
                row = sheet.createRow(i);
                for (int j = 0; j < 10; j++) {
                    row.createCell(j);
                }
            }
            //Go through each row and populate each cell with various Faker data types
            Faker faker = new Faker();
            for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                sheet.getRow(i).getCell(0).setCellValue(faker.idNumber().valid());
                sheet.getRow(i).getCell(1).setCellValue(faker.name().firstName());
                sheet.getRow(i).getCell(2).setCellValue(faker.name().lastName());
                sheet.getRow(i).getCell(3).setCellValue(faker.address().city());
                sheet.getRow(i).getCell(4).setCellValue(faker.country().name());
                sheet.getRow(i).getCell(5).setCellValue(faker.demographic().sex());
                sheet.getRow(i).getCell(6).setCellValue(faker.demographic().educationalAttainment());
                sheet.getRow(i).getCell(7).setCellValue(faker.number().numberBetween(1,99));
                sheet.getRow(i).getCell(8).setCellValue(faker.job().title());
                sheet.getRow(i).getCell(9).setCellValue(faker.bool().bool());
            }

            //Output the generated data to the Excel file specified as an argument
            workbook.write(outputStream);
            System.out.println("Your file has been successfully created.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void createFakerTableWithTwoDimensionalObjectArray(String filePath)  {

            //The try-with-resources statement automatically closes all the resources at the end of the statement
            try
                (XSSFWorkbook workbook = new XSSFWorkbook();
                 FileOutputStream outputStream = new FileOutputStream(filePath)) {

                XSSFSheet sheet = workbook.createSheet("FakerData");
                Faker faker = new Faker();

            //I am using a for loop to create an Excel table with 100 rows and 9 columns.
            //In it, I am initializing and Object-type array to hold all the Faker-generated data.

            //By using an Object array, I can populate each column with different data type
            //(because Object is a super clas other data-type classes inherit from).
            //I am creating here a two-dimensional array which has 9 columns (of various Faker-generated
            //data types) and one row (which I am going to multiply 100 times with the for loop).

                for (int i = 0; i < 100; i++) {
                    Object[][] table =
                        {
                            {faker.idNumber().valid(),
                             faker.name().firstName(),
                             faker.name().lastName(),
                             faker.address().city(),
                             faker.country().name(),
                             faker.job().title(),
                             faker.bool().bool(),
                             faker.demographic().sex(),
                             faker.number().numberBetween(1, 99)}
                        };

                    XSSFRow row = sheet.createRow(i);

                        for (int j = 0; j < 9; j++) {
                            row.createCell(j);
                            Object cellData = table[0][j];

            //Even though Object arrays can hold different data types, it is necessary to cast
            //each type before using the values.
                        if (cellData instanceof String) {
                            sheet.getRow(i).getCell(j).setCellValue((String) cellData);
                        }
                        if (cellData instanceof Integer) {
                            sheet.getRow(i).getCell(j).setCellValue((Integer) cellData);
                        }
                        if (cellData instanceof Double) {
                            sheet.getRow(i).getCell(j).setCellValue((Double) cellData);
                        }
                        if (cellData instanceof Boolean) {
                            sheet.getRow(i).getCell(j).setCellValue((Boolean) cellData);
                        }
                    }
                }

            //Output the generated data to the Excel file specified as an argument
                workbook.write(outputStream);
                System.out.println("Your file has been successfully created.");
            } catch (IOException e) {
            e.printStackTrace();
            }
    }

}
