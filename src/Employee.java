import java.util.ArrayList;
import java.util.HashMap;

public class Employee {
    private String name;
    private ArrayList<String> customersList = new ArrayList<>();
    private ArrayList<Customer> customers = new ArrayList<>();
    private double totalRevenue = 0;

    public Employee(){
        this.name = "";
    }
    public Employee(String name){
        this.name = name;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
    public void addCustomers(String customer){
        if(!customersList.contains(customer))
            customersList.add(customer);
    }
    public void addProductToCustomer(String product, String customer, double revenue){
        if(!customersList.contains(customer)) {
            customersList.add(customer);
            Customer cust = new Customer(customer);
            cust.addProduct(product,revenue);
            customers.add(cust);
        }else{
            for(int i = 0 ;i < customers.size();i++){
                Customer cust = customers.get(i);
                if(cust.toString().equals(customer)){
                    cust.addProduct(product,revenue);
                    break;
                }
            }
        }
        addRevenue(revenue);
    }
    public ArrayList getCustomers(){
        return customersList;
    }

    public void addRevenue(double revenue){
        this.totalRevenue += revenue;
    }
    @Override
    public String toString() {
        return getName();
    }

    private class Customer {
        private String name;
        private HashMap<String,Double> productAndRevenue = new HashMap();

        public Customer(){
            this.name = "";
        }
        public Customer(String name){
            this.name = name;
        }
        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }
        public void addProduct(String product, double revenue){
            if(productAndRevenue.containsKey(product)){
                double totalRev = productAndRevenue.get(product)+revenue;
                productAndRevenue.put(product,totalRev);
            } else{
                productAndRevenue.put(product,revenue);
            }
        }

        @Override
        public String toString() {
            return getName();
        }
    }

}
