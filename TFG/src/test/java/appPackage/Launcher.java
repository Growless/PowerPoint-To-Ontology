package appPackage;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.GridLayout;
import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.GroupLayout;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.LayoutStyle;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileSystemView;

public class Launcher extends JFrame implements ActionListener {

	private JPanel contentPane;
	static JLabel label;
	static JLabel label2;
	static JLabel label3;
	static String pptRoute = "";
	static String owlRoute = "";
	static String ontURL = "http://purl.org/spar/doco";
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		
		JFrame frame = new JFrame("PptToOntology");
		frame.setSize(500,200);
		frame.setVisible(true);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
		JButton button1 = new JButton(".owl Destination");	
		JButton button2 = new JButton("PPT location");
		JButton button3 = new JButton("Generate .owl");
		
		Launcher l1 = new Launcher();
		
		button1.addActionListener(l1);
		button1.setPreferredSize(new Dimension(200,25));

		
		button2.addActionListener(l1);
		button2.setPreferredSize(new Dimension(200,25));
		
		button3.addActionListener(l1);
		
		label = new JLabel("Nothing selected");
		label.setPreferredSize(new Dimension(100,25));
		
		label2 = new JLabel("Nothing selected");
		label2.setPreferredSize(new Dimension(100,25));
		
		label3 = new JLabel("");
		label3.setPreferredSize(new Dimension(100,25));
		
		JPanel p = new JPanel();
		
		GroupLayout layout = new GroupLayout(p);
		p.setLayout(layout);
		layout.setAutoCreateGaps(true);
		layout.setAutoCreateContainerGaps(true);
		
		layout.setHorizontalGroup(
				   layout.createSequentialGroup()
				   .addGroup(layout.createParallelGroup(GroupLayout.Alignment.CENTER)
				           .addComponent(button2)
				           .addComponent(button1))
				   
				   .addGroup(layout.createParallelGroup(GroupLayout.Alignment.LEADING)
				           .addComponent(label)
				           .addComponent(label2)
				           .addComponent(button3))
				   .addComponent(label3)
				   
				   
				);
				layout.setVerticalGroup(
				   layout.createSequentialGroup()
				      .addGroup(layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
				           .addComponent(button2)
				           .addComponent(label))

				      .addGroup(layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
				           .addComponent(button1)
				           .addComponent(label2))
				      .addPreferredGap(LayoutStyle.ComponentPlacement.UNRELATED, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
				      .addGroup(layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
					           .addComponent(button3)
					           .addComponent(label3))
				);
		//Adiciones finales
		p.add(button2);
		p.add(label);
		p.add(button1);
		p.add(label2);
		p.add(button3);
		frame.add(p);
		
		frame.show();
	}

	/**
	 * Default constructor
	 */
	public Launcher() {
		
		
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		

        String com = e.getActionCommand();
		// if the user presses the PPT button 
        if (com.equals("PPT location")) {
            // create an object of JFileChooser class
            JFileChooser j = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
 
            // invoke the showsSaveDialog function to show the save dialog
            int r = j.showSaveDialog(null);
 
            if (r == JFileChooser.APPROVE_OPTION) {
                // set the label to the path of the selected directory
                label.setText(j.getSelectedFile().getAbsolutePath());
                pptRoute = j.getSelectedFile().getAbsolutePath();
            }
            // if the user cancelled the operation
            else
                label.setText("Cancelled...");
        }
        // if the user presses the open dialog show the open dialog
        else if(com.equals(".owl Destination")) {
            // create an object of JFileChooser class
            JFileChooser j = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
 
            // set the selection mode to directories only
            j.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
 
            // invoke the showsOpenDialog function to show the save dialog
            int r = j.showOpenDialog(null);
 
            if (r == JFileChooser.APPROVE_OPTION) {
                // set the label to the path of the selected directory
                label2.setText(j.getSelectedFile().getAbsolutePath());
                owlRoute = j.getSelectedFile().getAbsolutePath();
            }
            // if the user cancelled the operation
            else
                label2.setText("Cancelled...");
        }
        //if the user presses the Generate .owl button
        else {
        	if(pptRoute.equals("") || owlRoute.equals("")) {
        		label3.setText("Plase complete both selections");
        		
        	}else {
        		
	        	String[] args = {pptRoute,ontURL,owlRoute};
	        	PptToOntology.main(args);
        	}
        }
		
	}

}
