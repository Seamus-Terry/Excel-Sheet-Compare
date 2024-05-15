package ExcelComp;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.File;

import java.io.IOException;
import java.io.PrintWriter;
import java.util.HashMap;
import java.util.Map;
import java.util.ArrayList;

public class GetExcel {

    // Creates Variables
    public String partNum;
    public static int x = 3, y = 0;

    // Creating Hash Map To Store Items With Keys
    public Map<String, ArrayList<String>> fileOne_Map = new HashMap<String, ArrayList<String>>();
    public Map<String, ArrayList<String>> fileTwo_Map = new HashMap<String, ArrayList<String>>();

    public void main(PrintWriter pw) throws IOException, InvalidFormatException {

        // Checks if 2 files are selected
        if (Window.fileOne != null && Window.fileTwo != null) {
            read(Window.fileOne, fileOne_Map);
            read(Window.fileTwo, fileTwo_Map);

            System.out.println("==================================================");

            System.out.println(fileOne_Map);
            System.out.println(fileTwo_Map);

            System.out.println("==================================================");

            partsCheck(pw);
            quantityCheck(pw);

        } else {
            System.out.println("Invalid File Path\n");
        }

    }

    public void read(File path, Map<String, ArrayList<String>> map) throws IOException, InvalidFormatException {

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(path);

        // Counting number of sheets
        System.out.print("Retrieving Sheets: ");
        for (Sheet sheet : workbook) {
            System.out.println(sheet.getSheetName());
        }

        // Getting the first sheet
        Sheet sheet = workbook.getSheetAt(0);

        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();

        // Collects row and column data
        for (Row row : sheet) {
            if (row != sheet.getRow(y)) {
                Cell ptNum = row.getCell(x);
                partNum = dataFormatter.formatCellValue(ptNum);
                map.put(partNum, new ArrayList<String>());
                for (Cell cell : row) {
                    if (cell != row.getCell(0) && cell != row.getCell(1) && cell != row.getCell(2) && cell != row.getCell(3)) {
                        String partData = dataFormatter.formatCellValue(cell);
                        map.get(partNum).add(partData);
                    }
                }
            }
            map.remove("");
        }

        // Closing the workbook
        workbook.close();
    }

    public void partsCheck(PrintWriter pw) {

        for (String keysOne : fileOne_Map.keySet()) {
            if (!fileTwo_Map.containsKey(keysOne)) {
                pw.println(keysOne + ", -" + fileOne_Map.get(keysOne).get(1) + ", 0, " + fileOne_Map.get(keysOne).get(0) + ", Removed from Inventory");
                System.out.println();
                System.out.print("Removed Parts: " + keysOne + " -" + fileOne_Map.get(keysOne).get(1) + " " + fileOne_Map.get(keysOne).get(0));
            }
        }

        System.out.println();

        for (String keysTwo : fileTwo_Map.keySet()) {
            if (!fileOne_Map.containsKey(keysTwo)) {
                pw.println(keysTwo + ", +" + fileTwo_Map.get(keysTwo).get(1) + ", " + fileTwo_Map.get(keysTwo).get(1) + ", " + fileTwo_Map.get(keysTwo).get(0) + ", Added to Inventory");
                System.out.println();
                System.out.print("Added Parts: " + keysTwo + " +" + fileTwo_Map.get(keysTwo).get(1) + " " + fileTwo_Map.get(keysTwo).get(0));
            }
        }

        System.out.println();
    }

    public void quantityCheck(PrintWriter pw) {

        System.out.println();

        for (String keysTwo : fileTwo_Map.keySet()) {
            if (fileOne_Map.containsKey(keysTwo)) {
                if (!fileTwo_Map.get(keysTwo).get(1).equals(fileOne_Map.get(keysTwo).get(1))) {

                    int change = Integer.parseInt(fileTwo_Map.get(keysTwo).get(1)) - Integer.parseInt(fileOne_Map.get(keysTwo).get(1));
                    int currentInv = Integer.parseInt(fileOne_Map.get(keysTwo).get(1)) + change;

                    if (currentInv > 0) {
                        pw.println(keysTwo + ", " + change + ", " + currentInv + ", " + fileTwo_Map.get(keysTwo).get(0) + ", Maintained Inventory");
                        System.out.println("Invetory Change: " + keysTwo + " +" + change + " " + fileTwo_Map.get(keysTwo).get(0));
                    } else {
                        pw.println(keysTwo + ", " + change + ", 0, " + fileTwo_Map.get(keysTwo).get(0) + ", Amount Change");
                        System.out.println("Invetory Change: " + keysTwo + " " + change + " " + fileTwo_Map.get(keysTwo).get(0));
                    }

                }
            }
        }
    }
}
