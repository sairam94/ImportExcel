package main;

import java.awt.BorderLayout;
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.table.DefaultTableModel;

import entities.Product;
import helper.ExcelHelper;

import javax.swing.JTable;
import javax.swing.JButton;
import javax.swing.JFileChooser;

import java.awt.event.ActionListener;
import java.io.IOException;
import java.util.List;
import java.awt.event.ActionEvent;
import javax.swing.JScrollPane;

public class JFrameImport extends JFrame {

	private JPanel contentPane;
	private JTable tableProducts;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					JFrameImport frame = new JFrameImport();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public JFrameImport() {
		setTitle("Import Excel File");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 561, 317);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		
		JButton btnImport = new JButton("Import");
		btnImport.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser jf = new JFileChooser();
				jf.setDialogTitle("Please Select a Excel File to Import");
				int result = jf.showOpenDialog(null);
				if(result==JFileChooser.APPROVE_OPTION){
					String excelPath = jf.getSelectedFile().getAbsolutePath();
					JOptionPane.showMessageDialog(null, excelPath);
					ExcelHelper eh = new ExcelHelper(excelPath);
					try {
						List<Product> listProduct = eh.readXLSXFile();
						DefaultTableModel dtm = new DefaultTableModel();
						dtm.addColumn("Id");
						dtm.addColumn("Name");
						dtm.addColumn("Price");
						dtm.addColumn("Quantity");
						dtm.addColumn("Status");
						dtm.addColumn("Creation Date");
						for(Product p : listProduct){
							dtm.addRow(new Object[] {p.getId(), p.getName(), p.getPrice(), p.getQuantity(), p.isStatus(), p.getCreationDate()});
						}
						tableProducts.setModel(dtm);
					} catch (IOException e) {
						JOptionPane.showMessageDialog(null, e.getMessage());
					}
				}
			}
		});
		btnImport.setBounds(230, 238, 89, 23);
		contentPane.add(btnImport);
		
		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(26, 12, 492, 210);
		contentPane.add(scrollPane);
		
		tableProducts = new JTable();
		scrollPane.setViewportView(tableProducts);
	}
}
