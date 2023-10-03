import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelEmployeeAnalyzer {
    public static void main(String[] args) {
        try {
            String filePath = "Assignment_Timecard.xlsx"; // Replace with your Excel file path

            FileInputStream fis = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0); // Assuming the data is on the first sheet

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    // Skip the header row
                    continue;
                }

                String name = ""; // Initialize with a default value
                String position = ""; // Initialize with a default value
                double hoursWorked = 0.0; // Initialize with a default value

                // Customize the cell index (0, 1, 3) based on your Excel file structure
                if (row.getCell(0) != null) {
                    name = row.getCell(0).getStringCellValue();
                }

                if (row.getCell(1) != null) {
                    position = row.getCell(1).getStringCellValue();
                }

                if (row.getCell(3) != null) {
                    hoursWorked = row.getCell(3).getNumericCellValue();
                }

                if (hasWorked7ConsecutiveDays(sheet, row.getRowNum())) {
                    System.out.println("Employee Name: " + name);
                    System.out.println("Position: " + position);
                    System.out.println("Worked 7 consecutive days.");
                }

                if (hasLessThan10HoursBetweenShifts(sheet, row.getRowNum())) {
                    System.out.println("Employee Name: " + name);
                    System.out.println("Position: " + position);
                    System.out.println("Has less than 10 hours between shifts but greater than 1 hour.");
                }

                if (hoursWorked > 14) {
                    System.out.println("Employee Name: " + name);
                    System.out.println("Position: " + position);
                    System.out.println("Worked for more than 14 hours in a single shift.");
                }
            }

            fis.close();
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static boolean hasWorked7ConsecutiveDays(Sheet sheet, int currentRow) {
        // Implement your logic to check if an employee has worked 7 consecutive days.
        // You can access the necessary data from the 'sheet' object.
        return false; // Replace with your logic
    }

    private static boolean hasLessThan10HoursBetweenShifts(Sheet sheet, int currentRow) {
        // Implement your logic to check if an employee has less than 10 hours between shifts but greater than 1 hour.
        // You can access the necessary data from the 'sheet' object.
        return false; // Replace with your logic
    }
}
