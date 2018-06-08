// Queries that I can do:
// 1. how much did I spend on chevron petrol pump in total
// 2. how much did I spend on Mayuri in total
// 3. Add any more categories that you want
import java.io.*;

import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// A category defines what was money spent on. E.g. food, travel, rent, etc.
// Each category may have more than one keywords. E.g. Food may have Pizza, the shops name, etc on the description part of the expense.
class Category {
    public List<String> categoryItems;
    public Category(List<String> list) {
        this.categoryItems = list;
    }
}

// Line represents each line of the csv/excel file from the bank statement. Each line is the list of columnValuesInLine, one from each column
class Line {
    public List<String> columnValuesInLine = new ArrayList<String>();
}

// Class with main function
public class ReadExcel {
    // Map of all the lines that fall in a category. A line would fall in a category of expense when description of that expense(line) contains one of the words in the category items.
    static Map<Category, List<Line>> categoryLinesMap = new HashMap<Category, List<Line>>();

    // column indexes from csv
    private static int amountCol = 4;
    private static int descriptionCol = 6;
    private static int postingDateCol = 1;
    private static int dateColumn = 1;

    private static int totalNoOfColumns = 0;

    // Define which column values are integers or doubles(basically numbers) and which are columnValuesInLine. Integers cells in excel should be set as an int
    // Else excel does not display it very good
    private static List<Integer> columnsWhoseValuesAreNumbers = new ArrayList<Integer>(Arrays.asList(4));

    // path to the bank statement
    private static String BANK_STATEMENT = "C:\\Users\\amaga\\Downloads\\bankstmt.xlsx";
    private static String OUTPUT_PATH = "C:\\Users\\amaga\\Downloads\\bankstmt_out.xlsx";

    // List of all the lines in the bank statement
    private static List<Line> allLinesInBankStmt;

    // main function
    public static void main(String args[]) throws IOException {
        // Create a list of all the lines in the bank statement
        allLinesInBankStmt = getAllLinesInBankStmt();

        createAllExpensesCategories();

        // Create a map of categories and lines, so that each category knows what all lines of expenses fall under it
        categorizeAllExpenses();

        createOutputExcelFile();
    }

    private static void createOutputExcelFile() throws IOException {
        // Create a Workbook
        Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

        /* CreationHelper helps us create instances of various things like DataFormat,
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet("Expenses");

        // Create a Row
        Row headerRow = sheet.createRow(0);
        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        // Create cells
        Cell cell = headerRow.createCell(0);
        cell.setCellValue("Description");
        cell.setCellStyle(headerCellStyle);

        Cell cell2 = headerRow.createCell(1);
        cell2.setCellValue("Date");
        cell2.setCellStyle(headerCellStyle);

        Cell cell3 = headerRow.createCell(2);
        cell3.setCellValue("Amount");
        cell3.setCellStyle(headerCellStyle);

        int rowNum = 1;
        for (Category c: categoryLinesMap.keySet()) {
            List<Line> linesInCategory = categoryLinesMap.get(c);

            int sumCategory = 0;

            // create a row for each line
            for(Line l : linesInCategory) {
                Row row = sheet.createRow(rowNum++);

                Cell one = row.createCell(0);
                one.setCellValue(l.columnValuesInLine.get(descriptionCol));

                Cell two = row.createCell(1);
                two.setCellValue(l.columnValuesInLine.get(dateColumn));

                Cell three = row.createCell(2);
                double expense = Double.parseDouble(l.columnValuesInLine.get(amountCol));
                three.setCellValue(expense);

                sumCategory += expense;
            }

            // now we have sum of all expenses of all lines in a category. Create a row for sum
            Row sumRow = sheet.createRow(rowNum++);
            Cell sumCell = sumRow.createCell(2);
            sumCell.setCellValue(sumCategory);
            sumCell.setCellStyle(headerCellStyle);

            // create a blank row
            sheet.createRow(rowNum++);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(OUTPUT_PATH);
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
    }

    // Create exhaustive list of all categories
    private static void createAllExpensesCategories() {
        // A category is a list of strings. Here we have an array of array of strings. Each array is the list of words in a category.
        // If you want to add a new category, just add a new array of words at the end. The code below will create a new category for each array of strings in the array.
        String members[][] =
                {
                        {"MAYURI FOODS AND VIDEO" },
                        {"KALIA INDIAN CUISINE"},
                        {"CHEVRON"},
                        {"COSTCO"},
                        {"RUCHI"},
                        {"CAFE 16"},
                        {"ACH Deposit MICROSOFT  - EDIPAYMENT"},
                        {"UBER"},
                        {"SKYPE"},
                        {"SAFEWAY"},
                        {"FAMILY PANCAKE HOUSE"},
                        {"CHAAT N ROLL"},
                        {"TARGET"},
                        {"STARBUCKS"},
                        {"KANISHKA"},
                        {"RIAMONEYTRANSFER"}
                };

        for (String[] arr: members){
            List<String> list = Arrays.asList(arr);
            Category c1 = new Category(list);

            categoryLinesMap.put(c1, null);
        }
    }

    // for each line find out which category the expense belongs to.
    private static void categorizeAllExpenses() {
        for(int i = 1; i< allLinesInBankStmt.size(); i++) {
            Line line = allLinesInBankStmt.get(i);
            String description = line.columnValuesInLine.get(descriptionCol);

            for(Category c : categoryLinesMap.keySet()) {
                if(descriptionIsInCategory(c, description)) {
                    List<Line> listOfLinesInCategory = categoryLinesMap.get(c);
                    if(listOfLinesInCategory == null) {
                        listOfLinesInCategory = new ArrayList<Line>();
                        categoryLinesMap.put(c, listOfLinesInCategory);
                    }
                    listOfLinesInCategory.add(line);

                    // ToDo: Remove this break if an expense can be in more than one category.
                    break;
                }
            }
        }
    }

    private static boolean descriptionIsInCategory(Category c, String description) {
        // if description contains any of the words in the category, then we say that description is in that category
        for(String s:c.categoryItems) {
            if(description.toUpperCase().contains(s.toUpperCase())) {
                return true;
            }
        }

        return false;
    }

    private static List<Line> getAllLinesInBankStmt() throws IOException {
        List<Line> allLines = new ArrayList<Line>();
        FileInputStream inputStream = new FileInputStream(new File(BANK_STATEMENT));

        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        totalNoOfColumns = firstSheet.getRow(0).getPhysicalNumberOfCells();

        Iterator<Row> iterator = firstSheet.iterator();

        // get first row to get column names
        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();

            Line l = new Line();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();

                String s  = "";
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        s = cell.getStringCellValue();
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        s = "" + cell.getBooleanCellValue();
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        s = "" + cell.getNumericCellValue();
                        break;
                    default:
                        s = "";
                        break;
                }
                l.columnValuesInLine.add(s);
            }
            allLines.add(l);
            break;
        }

        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            Line l = new Line();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();

                String s  = "";
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        s = cell.getStringCellValue();
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        s = "" + cell.getBooleanCellValue();
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        s = "" + cell.toString();
                        break;
                }
                l.columnValuesInLine.add(s);
            }
            allLines.add(l);
        }

        workbook.close();
        inputStream.close();

        return allLines;
    }

    private static void printLine(Line l) {
        System.out.println();
        for (String s : l.columnValuesInLine) {
            System.out.print(s + ", ");
        }
    }

    private static void printAllLines(List<Line> lines) {
        for (Line l : lines) {
            for(String s: l.columnValuesInLine) {
                System.out.print(s + ", ");
            }
            System.out.println();
        }
    }
}