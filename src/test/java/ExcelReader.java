import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReader {
    private final String excelFilePath;
    private XSSFSheet sheet;
    private XSSFWorkbook book;
    private String sheetName;

    /**
     * Конструктор для инициализации листа в эксель файле
     * @param excelFilePath путь к файлу
     * @throws IOException
     */
    public ExcelReader(String excelFilePath) throws IOException {
        this.excelFilePath = excelFilePath;
        File file = new File(excelFilePath);
        try{
            FileInputStream fileInputStream = new FileInputStream(file);
            book = new XSSFWorkbook(fileInputStream);
            sheet = book.getSheet("Лист1");
        } catch (IOException e){
            throw new IOException("Не поддерживаемый формат");
        }
    }
    /**
     * Конструктор для инициализации определенного листа в эксель файле
     * @param excelFilePath путь к файлу
     * @param sheetName название листа
     * @throws IOException
     */
    public ExcelReader(String excelFilePath, String sheetName) throws IOException {
        this.excelFilePath = excelFilePath;
        this.sheetName = sheetName;
        File file = new File(excelFilePath);
        try{
            FileInputStream fileInputStream = new FileInputStream(file);
            book = new XSSFWorkbook(fileInputStream);
            sheet = book.getSheet(sheetName);
        } catch (IOException e){
            throw new IOException("Не поддерживаемый формат");
        }
    }

    /**
     * Преобразование ячейки в строку
     * @param cell
     * @return значение ячейки
     * @throws Exception
     */
    private String cellToString(XSSFCell cell) throws Exception {
        Object result = null;
        CellType type = cell.getCellType();
        switch (type){
            case NUMERIC:
                result = cell.getNumericCellValue();
                break;
            case STRING:
                result = cell.getStringCellValue();
                break;
            case FORMULA:
                result = cell.getCellFormula();
                break;
            case BLANK:
                result = "";
                break;
            default: throw new Exception("Ошибка чтения ячейки");
        }
        return result.toString();
    }

    /**
     * Считает количество столбцов на листе
     * @return
     */
    private int xlsxCountColumn(){
        return sheet.getRow(0).getLastCellNum();
    }

    /**
     * Считает количество строк на листе
     * @return
     */
    private int xlsxCountRow(){
        return sheet.getLastRowNum() + 1;
    }

    /**
     * Инициализирует начальный эксель лист и читает все найденные столбцы и строки
     * @return возвращает массив с объектами для DataProvider
     * @throws Exception
     */
    public String[][] getSheetDataForTDD() throws Exception {
        File file = new File(excelFilePath);
        FileInputStream fileInputStream = new FileInputStream(file);
        book = new XSSFWorkbook(fileInputStream);
        sheet = book.getSheet("Лист1");
        int numberOfColumn = xlsxCountColumn();
        int numberOfRows = xlsxCountRow();
        String[][] data = new String[numberOfRows-1][numberOfColumn];
        for(int i = 1; i<numberOfRows; i++){
            for(int j = 0; j<numberOfColumn; j++){
                XSSFRow row = sheet.getRow(i);
                XSSFCell cell = row.getCell(j);
                String value = cellToString(cell);
                data[i-1][j] = value;
                if(value == null){
                    System.out.println("Пустые ячейки");
                }
            }
        }
        return data;
    }

    /**
     * Инициализирует кастомный эксель лист и читает все найденные столбцы и строки
     * @return возвращает массив с объектами для DataProvider
     * @throws Exception
     */
    public String[][] getCustomSheetForTDD() throws Exception {
        File file = new File(excelFilePath);
        FileInputStream fileInputStream = new FileInputStream(file);
        book = new XSSFWorkbook(fileInputStream);
        sheet = book.getSheet(sheetName); //открывает лист, который указан был через конструктор
        int numberOfColumn = xlsxCountColumn();
        int numberOfRows = xlsxCountRow();
        String[][] data = new String[numberOfRows-1][numberOfColumn];
        for(int i = 1; i<numberOfRows; i++){
            for(int j = 0; j<numberOfColumn; j++){
                XSSFRow row = sheet.getRow(i);
                XSSFCell cell = row.getCell(j);
                String value = cellToString(cell);
                data[i-1][j] = value;
                if(value == null){
                    System.out.println("Пустые ячейки");
                }
            }
        }
        return data;
    }
}
