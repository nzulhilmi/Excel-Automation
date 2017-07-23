
import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextArea;

public class GUI extends JFrame{
	static JDialog dialog = new JDialog();
	static String[] columnsNeeded;
	static String[] TAMs;
	static String file1;
	static String file2;
	static String outputFile;
	
	public static void main(String[] args) throws IOException {
		new GUI();
	}
	
	public GUI() {
		JFrame frame = new JFrame();
		
		/*
		Dimension screensize = Toolkit.getDefaultToolkit().getScreenSize();
		int width = (int) screensize.getWidth();
		int height = (int) screensize.getHeight();
		System.out.println("Screen widht: " + width);
		System.out.println("Screen height: " + height);
		*/
		
		JPanel masterPanel = new JPanel(new BorderLayout());
		JPanel panel = new JPanel();
		JPanel buttonPanel = new JPanel();
		
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setTitle("Excel Automation");
		frame.setSize(1300,800);
		frame.setMinimumSize(new Dimension(1000, 650));
		
		//Put the frame in the middle of the screen
		frame.setLocationRelativeTo(null);
		
		GridBagLayout layout = new GridBagLayout();
		panel.setLayout(layout);
		
		GridBagConstraints gbc = new GridBagConstraints();
		
		gbc.fill = GridBagConstraints.HORIZONTAL;
		gbc.gridx = 1;
		gbc.gridy = 0;
		JTextArea TAMTextArea = new JTextArea(2, 30);
		TAMTextArea.setBorder(BorderFactory.createLineBorder(getBackground(), 3));
		//String TAMText = convert(TAMs);
		//sTAMTextArea.setText(TAMText);
		panel.add(TAMTextArea, gbc);
		
		gbc.fill = GridBagConstraints.HORIZONTAL;
		gbc.gridx = 1;
		gbc.gridy = 1;
		JTextArea ColumnsTextArea = new JTextArea(2,30);
		ColumnsTextArea.setBorder(BorderFactory.createLineBorder(getBackground(), 3));
		//String columns = convert(columnsNeeded);
		//ColumnsTextArea.setText(columns);
		panel.add(ColumnsTextArea, gbc);
		
		gbc.fill = GridBagConstraints.HORIZONTAL;
		gbc.gridx = 1;
		gbc.gridy = 2;
		JTextArea FileTextArea1 = new JTextArea(2, 30);
		FileTextArea1.setBorder(BorderFactory.createLineBorder(getBackground(), 3));
		FileTextArea1.setText(file1);
		panel.add(FileTextArea1, gbc);
		
		gbc.fill = GridBagConstraints.HORIZONTAL;
		gbc.gridx = 1;
		gbc.gridy = 3;
		JTextArea FileTextArea2 = new JTextArea(2, 30);
		FileTextArea2.setBorder(BorderFactory.createLineBorder(getBackground(), 3));
		FileTextArea2.setText(file2);
		panel.add(FileTextArea2, gbc);
		
		gbc.fill = GridBagConstraints.HORIZONTAL;
		gbc.gridx = 1;
		gbc.gridy = 4;
		JTextArea OutputTextArea = new JTextArea(2,30);
		OutputTextArea.setBorder(BorderFactory.createLineBorder(getBackground(), 3));
		panel.add(OutputTextArea, gbc);
		
		gbc.fill = GridBagConstraints.HORIZONTAL;
		gbc.gridx = 0;
		gbc.gridy = 0;
		JLabel TAMs = new JLabel("TAMs: ");
		TAMs.setFont(new Font("Serif", Font.PLAIN, 20));
		panel.add(TAMs, gbc);
		
		gbc.fill = GridBagConstraints.HORIZONTAL;
		gbc.gridx = 0;
		gbc.gridy = 1;
		JLabel Columns = new JLabel("Columns: ");
		Columns.setFont(new Font("Serif", Font.PLAIN, 20));
		panel.add(Columns, gbc);
		
		gbc.fill = GridBagConstraints.HORIZONTAL;
		gbc.gridx = 0;
		gbc.gridy = 2;
		JLabel File1 = new JLabel("File 1: ");
		File1.setFont(new Font("Serif", Font.PLAIN, 20));
		panel.add(File1, gbc);
		
		gbc.fill = GridBagConstraints.HORIZONTAL;
		gbc.gridx = 0;
		gbc.gridy = 3;
		JLabel File2 = new JLabel("File 2: ");
		File2.setFont(new Font("Serif", Font.PLAIN, 20));
		panel.add(File2, gbc);
		
		gbc.fill = GridBagConstraints.HORIZONTAL;
		gbc.gridx = 0;
		gbc.gridy = 4;
		JLabel outputFile = new JLabel("Output File Name: ");
		outputFile.setFont(new Font("Serif", Font.PLAIN, 20));
		panel.add(outputFile, gbc);
		
		Font font = TAMTextArea.getFont();
		float fontsize =  font.getSize() + 15.0f;
		TAMTextArea.setFont(font.deriveFont(fontsize));
		FileTextArea1.setFont(font.deriveFont(fontsize));
		FileTextArea2.setFont(font.deriveFont(fontsize));
		ColumnsTextArea.setFont(font.deriveFont(fontsize));
		OutputTextArea.setFont(font.deriveFont(fontsize));
		
		JButton infoButton = new JButton("Info");
		infoButton.setPreferredSize(new Dimension(150, 50));
		infoButton.setBorder(BorderFactory.createMatteBorder(0, 15, 0, 15, getBackground()));
		
		infoButton.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				if(!dialog.isShowing()) {
					createDialog();
				}
			}
		});
		buttonPanel.add(infoButton);
		
		JButton extractButton = new JButton("Extract");
		extractButton.setPreferredSize(new Dimension(150, 50));
		extractButton.setBorder(BorderFactory.createMatteBorder(0, 15, 0, 15, getBackground()));
		/*
		 * action listener
		 * 	-get all the data from the text areas
		 * 	-set methods
		 * 	-run (disable the button for 5 seconds)
		 * 
		 */
		extractButton.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				excelWrite ex = new excelWrite();
				try {
					ex.execute_();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
			}
		});
		buttonPanel.add(extractButton);
		
		
		
		JButton closeButton = new JButton("Close");
		closeButton.setPreferredSize(new Dimension(150, 50));
		closeButton.setBorder(BorderFactory.createMatteBorder(0, 15, 0, 15, getBackground()));
		closeButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				frame.dispose();
				//System.exit(0);
			}
		});
		buttonPanel.add(closeButton);
		
		masterPanel.add(panel, BorderLayout.CENTER);
		masterPanel.add(buttonPanel, BorderLayout.SOUTH);
		
		frame.add(masterPanel);
		
		frame.setVisible(true);
	}
	
	private static void createDialog() {
		EventQueue.invokeLater(new Runnable() {
			@Override
			public void run() {
				// TODO Auto-generated method stub
				dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
				dialog.setTitle("Info");
				dialog.setSize(new Dimension(1000, 800));
				dialog.setVisible(true);
				dialog.setResizable(false);
			}
		});
	}
	
	public static String convert(String[] s) {
		String output = s[0];
		
		for(int i = 1; i < s.length; i++) {
			output = output + ", " + s[i];
		}
		
		return output;
	}
	
	//Get and set methods
	public static String[] getColumnsNeeded() {
		return columnsNeeded;
	}
	
	public static void setColumnsNeeded(String[] array) {
		columnsNeeded = array;
	}
	
	public static String[] getTAMs() {
		return TAMs;
	}
	
	public static void setTAMs(String[] array) {
		TAMs = array;
	}
	
	public static String getFile1() {
		return file1;
	}
	
	public static void setFile1(String s) {
		file1 = s;
	}
	
	public static String getFile2() {
		return file2;
	}
	
	public static void setFile2(String s) {
		file2 = s;
	}
	
	public static String getOutputFile() {
		return outputFile;
	}
	
	public static void setOutputFile(String s) {
		outputFile = s;
	}
}
