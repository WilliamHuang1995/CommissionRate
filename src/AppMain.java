import java.io.File;

public class AppMain {
    public static void main(String[] args) {
        String input = "C:\\Users\\willi\\Documents\\Projects\\Commission Calculator\\src\\October.xls";
        String output = "C:\\Users\\willi\\Documents\\Projects\\Commission Calculator\\src\\output.xls";
        File fin = new File(input);
        File fout = new File(output);
        Parser main = new Parser(fin,fout);
        main.run();

    }
}
