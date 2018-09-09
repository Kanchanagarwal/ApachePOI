package vserdiuk.apache.demo.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFTableStyleInfo;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;
import vserdiuk.apache.demo.model.Person;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.ZoneId;
import java.util.Date;
import java.util.List;

public class PersonExcelWriter {

    public void write(String fileName, List<Person> personList) {
        Workbook workbook = prepareWorkbook(personList);
        try (FileOutputStream outputStream = new FileOutputStream(fileName)) {
            workbook.write(outputStream);
            System.out.println("An Excel file " + fileName + " has been created");
            workbook.close();
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
    }

    /**
     * The method prepares a Workbook by preparing the header row and the content area
     *
     * @param personList
     * @return Workbook
     */
    private Workbook prepareWorkbook(List<Person> personList) {
        Workbook workbook = new XSSFWorkbook();
        XSSFSheet personSheet = (XSSFSheet) workbook.createSheet("Persons");

        XSSFTable table = createTable(personSheet);
        createTableStyle(table);

        personSheet.setZoom(200); //set autozoom to 200 points
        prepareHeader(personSheet);
        prepareContentArea(personSheet, personList);

        int columnAmount = personSheet.getRow(0).getLastCellNum(); //get count of columns
        addColumnToTable(table, columnAmount);

        int lastRowIndex = personSheet.getLastRowNum(); //get index of the last row

        //creating autofilter
        createAutofilter(workbook, table, columnAmount, lastRowIndex);

        //set table area
        setTableArea(workbook, table, columnAmount, lastRowIndex);

        //set autosizing columns according to a column content
        setAutoSizeColumn(personSheet);

        return workbook;
    }

    /**
     * The method prepares person sheet
     *
     * @param personSheet
     */
    private void prepareHeader(Sheet personSheet) {
        //create row
        Row headerRow = personSheet.createRow(0);

        //create cell style
        CellStyle cellStyle = getHeaderCellStyle(personSheet);

        //create cells
        Cell firstNameHeaderCell = headerRow.createCell(0);
        Cell lastNameHeaderCell = headerRow.createCell(1);
        Cell birthdayHeaderCell = headerRow.createCell(2);
        Cell emailHeaderCell = headerRow.createCell(3);
        Cell phoneNumberHeaderCell = headerRow.createCell(4);
        Cell marriedHeaderCell = headerRow.createCell(5);

        //set cells values
        firstNameHeaderCell.setCellValue("First name");
        lastNameHeaderCell.setCellValue("Last name");
        birthdayHeaderCell.setCellValue("Birthday");
        emailHeaderCell.setCellValue("Email");
        phoneNumberHeaderCell.setCellValue("Phone number");
        marriedHeaderCell.setCellValue("Married");

        //add style to cells
        firstNameHeaderCell.setCellStyle(cellStyle);
        lastNameHeaderCell.setCellStyle(cellStyle);
        birthdayHeaderCell.setCellStyle(cellStyle);
        emailHeaderCell.setCellStyle(cellStyle);
        phoneNumberHeaderCell.setCellStyle(cellStyle);
        marriedHeaderCell.setCellStyle(cellStyle);
    }

    /**
     * The method prepares person content area - lines with Person's data
     * (First name, Last name, Birthday, Email, Phone number, Married)
     *
     * @param personSheet
     * @param personList
     */
    private void prepareContentArea(Sheet personSheet, List<Person> personList) {
        for (int i=0; i<personList.size(); i++) {
            //create row
            Row row = personSheet.createRow(i + 1);

            //create common cell style
            CellStyle commonStyle = getCommontCellStyle(personSheet);

            CellStyle textStyle = getTextStyle(personSheet);
            CellStyle dateStyle = getDateStyle(personSheet);
            CellStyle phoneStyle = getPhoneStyle(personSheet);

            //create cells
            Cell firstNameCell = row.createCell(0);
            Cell lastNameCell = row.createCell(1);
            Cell birthdayCell = row.createCell(2);
            Cell emailCell = row.createCell(3);
            Cell phoneNumberCell = row.createCell(4);
            Cell marriedCell = row.createCell(5);

            //set cells values
            firstNameCell.setCellValue(personList.get(i).getFirstName());
            lastNameCell.setCellValue(personList.get(i).getLastName());
            birthdayCell.setCellValue(Date.from(personList.get(i).getBirthday()
                    .atStartOfDay(ZoneId.systemDefault()).toInstant()));
            emailCell.setCellValue(personList.get(i).getEmail());
            phoneNumberCell.setCellValue(personList.get(i).getPhoneNumber());
            marriedCell.setCellValue(personList.get(i).isMarried());

            //add style to cells
            firstNameCell.setCellStyle(textStyle);
            lastNameCell.setCellStyle(textStyle);
            birthdayCell.setCellStyle(dateStyle);
            emailCell.setCellStyle(textStyle);
            phoneNumberCell.setCellStyle(phoneStyle);
            marriedCell.setCellStyle(commonStyle);
        }
    }

    /**
     * The  method creates the table style
     *
     * @param table
     */
    private void createTableStyle(XSSFTable table) {
        XSSFTableStyleInfo style = (XSSFTableStyleInfo) table.getStyle();
        style.setName("TableStyleMedium2");
        style.setShowRowStripes(true);
    }

    /**
     * Creates the XSSFTable with the name Person
     *
     * @param personSheet
     * @return
     */
    private XSSFTable createTable(XSSFSheet personSheet) {
        XSSFTable table = personSheet.createTable();
        table.setName("Person");
        table.setDisplayName("Person_Table");

        // Create the initial style in a low-level way
        table.getCTTable().addNewTableStyleInfo();
        table.getCTTable().getTableStyleInfo().setName("TableStyleMedium2");
        return table;
    }

    /**
     * The method responsible for adding column to the XSSFTable
     *
     * @param table
     * @param columnAmount
     */
    private void addColumnToTable(XSSFTable table, int columnAmount) {
        for (int i=0; i<columnAmount; i++) {
            table.addColumn();
        }
    }

    private void setTableArea(Workbook workbook, XSSFTable table, int columnAmount, int lastRowIndex) {
        AreaReference reference = workbook.getCreationHelper().createAreaReference(
                new CellReference(0, 0), new CellReference(lastRowIndex, columnAmount-1));
        table.setCellReferences(reference);
    }

    /**
     * The method creates an autrofilter for table
     *
     * @param workbook
     * @param table
     * @param columnAmount
     * @param lastRowIndex
     */
    private void createAutofilter(Workbook workbook, XSSFTable table, int columnAmount, int lastRowIndex) {
        AreaReference referenceFilter = workbook.getCreationHelper().createAreaReference(
                new CellReference(0, 0), new CellReference(lastRowIndex, columnAmount-1));
        CTAutoFilter autoFilter = CTAutoFilter.Factory.newInstance();
        autoFilter.setRef(referenceFilter.formatAsString());
        table.getCTTable().setAutoFilter(autoFilter);
    }

    /**
     * The method resizes columns according to a column content
     *
     * @param personSheet
     */
    private void setAutoSizeColumn(Sheet personSheet) {
        int columnCount = personSheet.getRow(0).getLastCellNum()-1;
        for (int i=0; i<columnCount; i++) {
            personSheet.autoSizeColumn(i); // set the auto size for a column according to a column content
        }
    }

    /**
     * The method prepares the CellStyle for the header row
     *
     * @param personSheet
     * @return
     */
    private CellStyle getHeaderCellStyle(Sheet personSheet) {
        Font font = personSheet.getWorkbook().createFont();
        font.setBold(true); //setting font style as bold
        font.setFontName("Arial"); //setting font name as Arial
        font.setFontHeightInPoints((short) 14); //setting font size
        font.setColor(IndexedColors.WHITE.getIndex()); //setting font color

        CellStyle cellStyle = personSheet.getWorkbook().createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER); //setting horizontal alignment to center
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); //setting vertical alignment to center
        cellStyle.setBorderTop(BorderStyle.MEDIUM); //setting border style to medium (border color is black by default)
        cellStyle.setBorderRight(BorderStyle.MEDIUM);
        cellStyle.setBorderBottom(BorderStyle.MEDIUM);
        cellStyle.setBorderLeft(BorderStyle.MEDIUM);
        cellStyle.setFont(font);
        return cellStyle;
    }

    /**
     * The method prepares the common CellStyle for Person table content area
     *
     * @param personSheet
     * @return
     */
    private CellStyle getCommontCellStyle(Sheet personSheet) {
        CellStyle cellStyle = personSheet.getWorkbook().createCellStyle();

        //create italic Calibri font with 11 points height
        Font font = personSheet.getWorkbook().createFont();
        font.setItalic(true);
        font.setFontName("Calibri");
        font.setFontHeightInPoints((short) 11);

        cellStyle.setFont(font);

        //set thin border around cell
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);

        return cellStyle;
    }

    /**
     * The method prepares the CellStyle for text cells
     *
     * @param personSheet
     * @return
     */
    private CellStyle getTextStyle(Sheet personSheet) {
        CellStyle cellStyle = getCommontCellStyle(personSheet);
        DataFormat dataFormat = personSheet.getWorkbook().createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat("@"));
        return cellStyle;
    }

    /**
     * The method prepares the CellStyle for date cells according to MMMM dd, yyyy mask
     *
     * @param personSheet
     * @return
     */
    private CellStyle getDateStyle(Sheet personSheet) {
        CellStyle cellStyle = getCommontCellStyle(personSheet);
        CreationHelper createHelper = personSheet.getWorkbook().
                getCreationHelper();
        cellStyle.setDataFormat(createHelper.createDataFormat().
                getFormat("MMMM dd, yyyy"));
        return cellStyle;
    }

    /**
     * The method prepares the CellStyle for phone number cells according to (###) ###-#### mask
     *
     * @param personSheet
     * @return
     */
    private CellStyle getPhoneStyle(Sheet personSheet) {
        CellStyle cellStyle = getCommontCellStyle(personSheet);
        CreationHelper createHelper = personSheet.getWorkbook().getCreationHelper();
        cellStyle.setDataFormat(createHelper.createDataFormat()
                .getFormat("(###) ###-####"));
        return cellStyle;
    }

}










