import java.awt.BorderLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

public class AppMain {
    static File fin;
    static File fout;
    public static void main(String[] args) throws ClassNotFoundException, InstantiationException, IllegalAccessException, UnsupportedLookAndFeelException {
        //Parser main = new Parser(fin,fout);
        //main.run();

        UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        JFrame window = new JFrame("Commission");
        JPanel topPanel = new JPanel();
        JPanel middlePanel = new JPanel();
        JPanel bottomPanel = new JPanel();
        final JButton openFileChooser = new JButton("選擇檔案");
        final JButton openFileChooser2 = new JButton("選擇位置");
        final JButton ok = new JButton("OK");
        openFileChooser.setSize(100,100);
        openFileChooser2.setSize(100,100);
        JTextField t1=new JTextField("Input",18);
        JTextField t2=new JTextField("Output",18);
        t1.setEditable(false);t2.setEditable(false);
        t1.setBounds(50,100, 200,30);
        t2.setBounds(50,100, 200,30);

        final JFileChooser fc = new JFileChooser();
        final FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Files","xls","xlsx");
        fc.setFileFilter(filter);

        openFileChooser.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                fc.setCurrentDirectory(new java.io.File("user.home"));
                fc.setDialogTitle("選擇Excel檔案");
                fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
                if (fc.showOpenDialog(openFileChooser) == JFileChooser.APPROVE_OPTION) {
                    String sf = fc.getSelectedFile().getAbsolutePath();
                    t1.setText(sf);
                }
            }
        });

        openFileChooser2.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                fc.setCurrentDirectory(new java.io.File("user.home"));
                fc.setDialogTitle("選擇Output位置與名稱");
                fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                if (fc.showOpenDialog(openFileChooser) == JFileChooser.APPROVE_OPTION) {
                    String sf = fc.getSelectedFile().getAbsolutePath();
                    t2.setText(sf+"\\Output.xls");
                }
            }
        });

        ok.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                fin = new File(t1.getText());
                fout = new File(t2.getText());
                Parser main = new Parser(fin, fout);
                if(main.run()){
                    JOptionPane.showMessageDialog(null, "程式執行成功！");
                }else{
                    JOptionPane.showMessageDialog(null, "程式出錯！請確保檔案目前沒有正被使用！");
                }
                ok.setVisible(false);
                System.exit(0);

            }
        });
        window.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        window.add(BorderLayout.NORTH, topPanel);
        topPanel.add(t1);
        topPanel.add(openFileChooser);
        window.add(BorderLayout.CENTER,middlePanel);
        middlePanel.add(t2);
        middlePanel.add(openFileChooser2);
        window.add(BorderLayout.SOUTH,bottomPanel);
        bottomPanel.add(ok);
        window.setSize(400, 250);
        window.setResizable(false);
        window.setVisible(true);
        window.setLocationRelativeTo(null);


    }
}

