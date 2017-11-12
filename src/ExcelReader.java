
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class ExcelReader {
    static HSSFSheet spreadsheet;
    static HSSFWorkbook workbook;
    static Row row;
    static ArrayList<Employee> employeeList = new ArrayList<>();
    static ArrayList<String> nameList = new ArrayList<>();

    public static void main(String[] args){
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(new File("C:\\Users\\willi\\Documents\\Projects\\Commission Calculator\\src\\October.xls"));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            workbook = new HSSFWorkbook(fis);
        } catch (IOException e ) {
            e.printStackTrace();
        }
        //Spreadsheet needs to be in first slot.
        spreadsheet = workbook.getSheetAt(0);;
        Iterator<Row> rowIterator = spreadsheet.iterator();
        rowIterator.next();
        parseSpreadsheet(rowIterator);

        try {
            fis.close();
        }catch(IOException e){

        }

    }

    private static void parseSpreadsheet(Iterator<Row> rowIterator) {
        String employee;
        String customer;
        String productName;
        double revenue;
        while(rowIterator.hasNext()) {
            row = rowIterator.next();
            employee = row.getCell(0).toString();
            customer = row.getCell(2).toString();
            productName = row.getCell(4).toString();
            revenue = row.getCell(9).getNumericCellValue();
            if(productName.contains("硬化劑")||productName.contains("發泡劑")||productName.contains("設備")||productName.contains("薄膜")) {
                //do things
                if(!nameList.contains(employee)) {
                    Employee emp = new Employee(employee);

                    employeeList.add(emp);
                    nameList.add(employee);

                    emp.addProductToCustomer(productName,customer,revenue);

                }else{
                    for(int i = 0; i<employeeList.size();i++){
                        Employee emp = employeeList.get(i);
                        if(employeeList.get(i).getName().equals(employee)){

                            emp.addProductToCustomer(productName,customer,revenue);
                            break;
                        }
                    }


                }

                System.out.println(employee + " " + customer + " " + productName + " " + revenue);
            }
        }

        int i = 0;
        System.out.println("hehe");
    }
}
