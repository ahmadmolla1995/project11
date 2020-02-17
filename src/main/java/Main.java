import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;
import java.io.FileOutputStream;


public class Main {
    public static void main(String[] args) {
        List<Employee> employees = createObjects();
        writeExcelFile(employees);

        List<Employee> excelFileContent = new ArrayList<>();
        readExcelData(excelFileContent);
        excelFileContent.forEach(e -> System.out.println(e.toString()));
    }

    private static List<Employee> createObjects() {
        List<Employee> employees = new ArrayList<>();

        employees.add(new Employee("ahmad", "molla", "american"));
        employees.add(new Employee("amir", "ranjbar", "germany"));
        employees.add(new Employee("mohammad", "moridi", "canadian"));
        employees.add(new Employee("parisa", "saat", "france"));
        employees.add(new Employee("kosar", "bakhshi", "switzerland"));
        employees.add(new Employee("romina", "sahebzamani", "russian"));
        employees.add(new Employee("sara", "saffari", "sweden"));
        employees.add(new Employee("sima", "jahedi", "turkish"));
        employees.add(new Employee("ahmad", "roshanfar", "pakistan"));

        return employees;
    }

    private static void writeExcelFile (List<Employee> employees) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Employee Data");

        Map<Integer, Object[]> data = new TreeMap<>();
        data.put(0, new Object[] {"First name", "Last name", "National ID"});
        for(int i = 0; i < employees.size(); i++)
            data.put(i + 1, new Object[] {employees.get(i).getFirstName(), employees.get(i).getLastName(), employees.get(i).getNationalID()});


        Set<Integer> keySet = data.keySet();
        int rowNum = 0;
        for (Integer key : keySet) {
            Row row = sheet.createRow(rowNum++);
            Object [] objArr = data.get(key);

            int cellnum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellnum++);
                cell.setCellValue((String)obj);
            }
        }

        try {
            FileOutputStream fileOutputStream = new FileOutputStream(new File("employee.xlsx"));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void readExcelData(List<Employee> excelFileContent) {
        try {
            FileInputStream file = new FileInputStream(new File("employee.xlsx"));
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0);

            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Iterator<Cell> cellIterator = rowIterator.next().cellIterator();

                String firstName = cellIterator.next().getStringCellValue();
                String lastName = cellIterator.next().getStringCellValue();
                String nationality = cellIterator.next().getStringCellValue();

                excelFileContent.add(new Employee(firstName, lastName, nationality));
            }

            file.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
