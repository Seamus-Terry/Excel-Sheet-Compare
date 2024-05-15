package ExcelComp;


import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextArea;
import org.apache.logging.log4j.simple.SimpleLogger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class Window implements ActionListener {

    //Variables
    public JFrame window;
    public JButton start, first, second;
    public FileOutputStream fos;
    public PrintWriter pw;
    public JTextArea disOne, disTwo;
    public JLabel error;
    public static JFileChooser saveLoc, selectOne, selectTwo;
    public static File fileOne, fileTwo, fileSave;
    

    public Window() {

        //Window Initializer
        window = new JFrame();
        window.setTitle("Excelora - Copyright © 2024 by Seamus P. Terry");
        window.setLayout(null);
        window.setSize(545, 225);
        window.setResizable(false);
        window.setLocationRelativeTo(null);
        window.getContentPane().setBackground(new Color(0x003f5b));
        window.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        first = new JButton("File One");
        first.setFocusable(false);
        first.addActionListener(this);
        first.setBounds(30, 30, 100, 25);

        second = new JButton("File Two");
        second.setFocusable(false);
        second.addActionListener(this);
        second.setBounds(30, 80, 100, 25);

        disOne = new JTextArea();
        disOne.setFocusable(false);
        disOne.setBounds(150, 30, 350, 25);

        disTwo = new JTextArea();
        disTwo.setFocusable(false);
        disTwo.setBounds(150, 80, 350, 25);

        start = new JButton("Start");
        start.setFocusable(false);
        start.addActionListener(this);
        start.setBounds(430, 130, 70, 25);

        error = new JLabel();
        error.setForeground(Color.ORANGE);
        error.setBounds(30, 140, 500, 20);

        //File Chooser
        saveLoc = new JFileChooser();
        saveLoc.addActionListener(this);

        selectOne = new JFileChooser();
        selectOne.addActionListener(this);

        selectTwo = new JFileChooser();
        selectTwo.addActionListener(this);

        //Adding Components to the Window Frame
        window.add(first);
        window.add(second);
        window.add(start);
        window.add(disOne);
        window.add(disTwo);
        window.add(error);
        window.setVisible(true);

    }

    @Override
// Checks for Button Presses
    public void actionPerformed(ActionEvent e) {

        if (e.getSource() == start) {

            if (fileOne != null && fileTwo != null && fileOne.toString().matches("(.*).xlsx") && fileTwo.toString().matches("(.*).xlsx")) {

                error.setText("");

                GetExcel readExcel = new GetExcel();

                // Creates Integer to get whether "Save" or "Cancel" is pressed
                int result = saveLoc.showSaveDialog(window);
                if (result == JFileChooser.APPROVE_OPTION) {
                    fileSave = saveLoc.getSelectedFile();
                    System.out.println("Save Location: " + fileSave + "\n");

                    Thread t1 = new Thread(new Runnable() {
                        @Override
                        public void run() {

                            try {
                                fos = new FileOutputStream(Window.fileSave + ".csv", true);
                                pw = new PrintWriter(fos);
                                pw.println("Effected Part #, Inventory Change, Current Inventory, Part Name, Action");
                            } catch (FileNotFoundException ex) {
                                Logger.getLogger(Window.class.getName()).log(Level.SEVERE, null, ex);
                            }

                            try {
                                readExcel.main(pw);
                            } catch (IOException ex) {
                                SimpleLogger.CATCHING_MARKER.getName();
                                Logger.getLogger(Window.class.getName()).log(Level.SEVERE, null, ex);
                            } catch (InvalidFormatException ex) {
                                SimpleLogger.CATCHING_MARKER.getName();
                                Logger.getLogger(Window.class.getName()).log(Level.SEVERE, null, ex);
                            }

                            pw.close();
                        }
                    });
                    t1.start();
                }
            } else {
                try {
                    Thread.sleep(1000);
                } catch (InterruptedException ex) {
                    Logger.getLogger(Window.class.getName()).log(Level.SEVERE, null, ex);
                }
                error.setText("ERROR: Double Check Selected Files");
            }
        } else if (e.getSource() == first) {
            // Creates Integer to get whether "Open" or "Cancel" is pressed
            int result = selectOne.showOpenDialog(window);
            if (result == JFileChooser.OPEN_DIALOG) {
                fileOne = selectOne.getSelectedFile();
                disOne.setText(fileOne.toString());
                System.out.println("File Two: " + fileOne);
            }
        } else if (e.getSource() == second) {
            // Creates Integer to get whether "Open" or "Cancel" is pressed
            int result = selectTwo.showOpenDialog(window);
            if (result == JFileChooser.OPEN_DIALOG) {
                fileTwo = selectTwo.getSelectedFile();
                disTwo.setText(fileTwo.toString());
                System.out.println("File Two: " + fileTwo);
            }
        } 
    }
}
