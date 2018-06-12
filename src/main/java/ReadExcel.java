// Queries that I can do:
// 1. how much did I spend on chevron petrol pump in total
// 2. how much did I spend on Mayuri in total
// 3. Add any more categories that you want
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.*;

import java.util.*;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;

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
    private static int AMOUNT_COLUMN = 4;
    private static int DESCRIPTION_COLUMN = 6;
    private static int DATE_COLUMN = 1;

    // Define which column values are integers or doubles(basically numbers) and which are columnValuesInLine. Integers cells in excel should be set as an int
    // Else excel does not display it very good
    private static List<Integer> columnsWhoseValuesAreNumbers = new ArrayList<Integer>(Arrays.asList(4));

    private static String INPUT_CONFIG_FILE;
    private static String HELP_TEXT = "src/input.txt";

    // path to the bank statement
    private static String BANK_STATEMENT_FILE;

    // List of all the lines in the bank statement
    private static List<Line> allLinesInBankStmt;

    // main function
    public static void main(String args[]) {
        showUI();
    }

    private static void showUI() {
        int panelWidth = 600;
        int panelHeight = 500;
        int width = panelWidth/3 - 5, height = 20;
        int left = 10;
        int down = 20;

        JFrame frame= new JFrame("Categorize your expense");
        JPanel panel=new JPanel();
        panel.setBounds(40,80,panelWidth,panelHeight);

        JButton inputFile = new JButton("Select expense file");
        inputFile.setBounds(left,down,width,height);
        inputFile.addActionListener(new AbstractAction() {
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                int returnValue = fileChooser.showOpenDialog(null);
                if (returnValue == JFileChooser.APPROVE_OPTION) {
                    File selectedFile = fileChooser.getSelectedFile();
                    System.out.println(selectedFile.getName());
                    BANK_STATEMENT_FILE = selectedFile.getAbsolutePath();
                }
            }
        });

        JButton configFile=new JButton("Select config file");
        configFile.addActionListener(new AbstractAction() {
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                int returnValue = fileChooser.showOpenDialog(null);
                if (returnValue == JFileChooser.APPROVE_OPTION) {
                    File selectedFile = fileChooser.getSelectedFile();
                    System.out.println(selectedFile.getName());
                    INPUT_CONFIG_FILE = selectedFile.getAbsolutePath();
                }
            }
        });
        configFile.setBounds(left + width + 10 ,down,width,height);

        JButton learnMore=new JButton("Learn more");
        learnMore.addActionListener(new AbstractAction() {
            public void actionPerformed(ActionEvent e) {
                JTextArea txtArea = new JTextArea();

                BufferedReader in = null;
                StringBuilder sb = new StringBuilder();
                try {
                    in = new BufferedReader(new FileReader(HELP_TEXT));
                    String line = in.readLine();
                    while(line != null){
                        sb.append(line + "\n");
                        line = in.readLine();
                    }
                } catch (FileNotFoundException e1) {
                    e1.printStackTrace();
                } catch (IOException e1) {
                    e1.printStackTrace();
                }

                txtArea.append(sb.toString());

                Dimension d = new Dimension(900, 500);
                Dimension d2 = new Dimension(900, 500);
                txtArea.setPreferredSize(d);
                txtArea.setVisible(true);
                txtArea.setEditable(false);

                JDialog dialog = new JDialog();
                dialog.setPreferredSize(d2);
                dialog.add(txtArea);
                dialog.pack();
                dialog.setVisible(true);
                dialog.setModal(true);
            }
        });
        learnMore.setBounds(left + 2*width + 10, down, width, height);

        JButton execute =new JButton("Execute");
        execute.setBounds(left,2*down + 10,width,height);
        execute.addActionListener(new AbstractAction() {
            public void actionPerformed(ActionEvent e) {
                try {
                    execute();
                    JOptionPane.showConfirmDialog(null,
                            "Close", "Categorization Done", JOptionPane.DEFAULT_OPTION);
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            }
        });

        panel.add(inputFile);
        panel.add(configFile);
        panel.add(learnMore);
        panel.add(execute);

        frame.add(panel);
        frame.setSize(panelWidth,panelHeight);
        frame.setLayout(null);
        frame.setVisible(true);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    }

    // this is the main method that does all calculation
    private static void execute() throws IOException {
        processConfigFile(INPUT_CONFIG_FILE);

        // Create a list of all the lines in the bank statement
        allLinesInBankStmt = getAllLinesInBankStmt();

        // Create a map of categories and lines, so that each category knows what all lines of expenses fall under it
        categorizeAllExpenses();

        createOutputExcelFile();
    }

    // It does these things
    // 1. Reads the path of input file,
    // 2. Reads path of output formatted file
    // 3. Creates List of expense categories
    private static void processConfigFile(String filePath) {
        try {
            BufferedReader br = new BufferedReader(new FileReader(filePath));
            String line = "";
            while ((line = br.readLine()) != null) {
                if(line.startsWith("//")) {
                    continue;
                }

                List<String> itemsInCategory = new ArrayList<String>();
                // All strings in the list are separated by comma
                StringTokenizer stringTokenizer = new StringTokenizer(line, ",");
                while(stringTokenizer.hasMoreTokens()) {
                    String token = stringTokenizer.nextToken();
                    itemsInCategory.add(token);
                }

                // put the key as the category.
                // Value is an empty array list which will be populated when we start iterating over all the lines in the expense file
                categoryLinesMap.put(new Category(itemsInCategory), new ArrayList<Line>());
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getOutputFilePath() {
        return BANK_STATEMENT_FILE.substring(0, BANK_STATEMENT_FILE.lastIndexOf(".")) + "_out.xlsx";
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
                one.setCellValue(l.columnValuesInLine.get(DESCRIPTION_COLUMN));

                Cell two = row.createCell(1);
                two.setCellValue(l.columnValuesInLine.get(DATE_COLUMN));

                Cell three = row.createCell(2);
                double expense = Double.parseDouble(l.columnValuesInLine.get(AMOUNT_COLUMN));
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
        FileOutputStream fileOut = new FileOutputStream(getOutputFilePath());
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
    }

    // For each line find the expense description. Then find out which category the description matches to.
    private static void categorizeAllExpenses() {
        for(int i = 1; i< allLinesInBankStmt.size(); i++) {
            Line line = allLinesInBankStmt.get(i);
            String description = line.columnValuesInLine.get(DESCRIPTION_COLUMN);

            for(Category c : categoryLinesMap.keySet()) {
                if(descriptionIsInCategory(c, description)) {
                    List<Line> listOfLinesInCategory = categoryLinesMap.get(c);
                    if(listOfLinesInCategory == null) {
                        listOfLinesInCategory = new ArrayList<Line>();
                        categoryLinesMap.put(c, listOfLinesInCategory);
                    }
                    listOfLinesInCategory.add(line);

                    if(!oneLineCanBeInMultipleCategories()) {
                        break;
                    }
                }
            }
        }
    }

    // If one line/expense may fall under multiple categories, then true, else false
    private static boolean oneLineCanBeInMultipleCategories() {
        return true;
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
        FileInputStream inputStream = new FileInputStream(new File(BANK_STATEMENT_FILE));

        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        int totalNoOfColumns = firstSheet.getRow(0).getPhysicalNumberOfCells();

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