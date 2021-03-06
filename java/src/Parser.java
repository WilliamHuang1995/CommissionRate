
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;


import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;

public class Parser {
    static ArrayList<Employee> employeeList;
    static ArrayList<String> nameList;
    private static File input;
    private static File output;
    private static boolean outputExist;
    public Parser(){
        this.input = null;
        this.output = null;
        nameList = null;
        employeeList = null;
    }
    public Parser(File input, File output){
        this.input = input;
        this.output = output;
        this.outputExist = output.exists();
        nameList = new ArrayList<>();
        employeeList = new ArrayList<>();
    }
    public static boolean run(){
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(input);
        } catch (FileNotFoundException e) {
            return false;
        }
        HSSFWorkbook workbook = null;
        try {
            workbook = new HSSFWorkbook(fis);
        } catch (IOException e ) {
            return false;
        }
        //Spreadsheet needs to be in first slot.
        HSSFSheet spreadsheet = workbook.getSheetAt(0);
        ;
        Iterator<Row> rowIterator = spreadsheet.iterator();
        rowIterator.next();
        parseSpreadsheet(rowIterator);
        try {
            outputSpreadsheet();
        }catch(IOException e){
            return false;
        }
        try {
            fis.close();
        }catch(IOException e){
            return false;
        }
        return true;

    }
    private static void parseSpreadsheet(Iterator<Row> rowIterator) {
        String employee;
        String customer;
        String productName;
        double revenue;
        while(rowIterator.hasNext()) {
            Row row = rowIterator.next();
            employee = row.getCell(0).toString();
            customer = row.getCell(2).toString();
            productName = row.getCell(4).toString();
            revenue = row.getCell(9).getNumericCellValue();
            if (productName.contains("硬化劑") || productName.contains("發泡劑") || productName.contains("設備") || productName.contains("薄膜")) {
                if (!nameList.contains(employee.toUpperCase())) {
                    Employee emp = new Employee(employee.toUpperCase());

                    employeeList.add(emp);
                    nameList.add(employee.toUpperCase());

                    emp.addProductToCustomer(productName, customer, revenue);

                } else {
                    for (int i = 0; i < employeeList.size(); i++) {
                        Employee emp = employeeList.get(i);
                        if (employeeList.get(i).getName().equals(employee.toUpperCase())) {

                            emp.addProductToCustomer(productName, customer, revenue);
                            break;
                        }
                    }


                }

                //System.out.println(employee + " " + customer + " " + productName + " " + revenue);
            }
        }
    }
    private static void outputSpreadsheet() throws IOException {
        FileOutputStream fos = new FileOutputStream(output);
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet ss;
        for(int i= 0;i<employeeList.size();i++){
            Employee emp = employeeList.get(i);
            ss = wb.createSheet(emp.getName());
            int rowSize = 0;
            Row row = ss.createRow(rowSize++);
            String[] header = {"業務","客戶","產品","金額","2%佣金","3%佣金","5%佣金","Subtotal"};
            for(int k = 0; k<header.length;k++){
                row.createCell(k).setCellValue(header[k]);
            }
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
