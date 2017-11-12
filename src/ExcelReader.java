
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;


import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

public class ExcelReader {
    static HSSFSheet spreadsheet;
    static HSSFWorkbook workbook;
    static Row row;
    static ArrayList<Employee> employeeList = new ArrayList<>();
    static ArrayList<String> nameList = new ArrayList<>();

    public static void main(String[] args){
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(new File("C:\\Users\\willi\\Documents\\Projects\\Commission Calculator\\src\\test.xls"));
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
            outputSpreadsheet();
        }catch(IOException e){
            e.printStackTrace();
        }
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
            if (productName.contains("硬化劑") || productName.contains("發泡劑") || productName.contains("設備") || productName.contains("薄膜")) {
                //do things
                if (!nameList.contains(employee)) {
                    Employee emp = new Employee(employee);

                    employeeList.add(emp);
                    nameList.add(employee);

                    emp.addProductToCustomer(productName, customer, revenue);

                } else {
                    for (int i = 0; i < employeeList.size(); i++) {
                        Employee emp = employeeList.get(i);
                        if (employeeList.get(i).getName().equals(employee)) {

                            emp.addProductToCustomer(productName, customer, revenue);
                            break;
                        }
                    }


                }

                System.out.println(employee + " " + customer + " " + productName + " " + revenue);
            }
        }
    }
    private static void outputSpreadsheet() throws IOException {
        FileOutputStream fos = new FileOutputStream(new File("C:\\Users\\willi\\Documents\\Projects\\Commission Calculator\\src\\output.xls"));
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet ss = wb.createSheet("Output");
        int rowSize = 0;
        Row row = ss.createRow(rowSize++);
        String[] header = {"業務","客戶","產品","金額","2%佣金","3%佣金","5%佣金","Subtotal"};
        for(int i = 0; i<header.length;i++){
            row.createCell(i).setCellValue(header[i]);
        }
        for(int i= 0;i<employeeList.size();i++){
            Employee emp = employeeList.get(i);
            ArrayList<String> customerList = emp.getCustomers();
            for(int j = 0; j<customerList.size();j++){
                String customer = customerList.get(j);
                row = ss.createRow(rowSize++);
                row.createCell(0).setCellValue(emp.toString());
                row.createCell(1).setCellValue(customer);
                Employee.Customer cust = emp.customer(customer);
                for(Map.Entry e:cust.productAndRevenue.entrySet()){
                    String product = (String) e.getKey();
                    double revenue = (double) e.getValue();
                    row.createCell(2).setCellValue(product);
                    row.createCell(3).setCellValue(revenue);
                    row.createCell(4).setCellValue(revenue*0.02);
                    row.createCell(5).setCellValue(revenue*0.03);
                    row.createCell(6).setCellValue(revenue*0.05);
                    row = ss.createRow(rowSize++);
                }
            }
        }
        wb.write(fos);
        fos.close();
    }
}
