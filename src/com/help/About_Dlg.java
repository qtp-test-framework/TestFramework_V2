/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package com.help;

import com.util.Constants;
import java.awt.*;
import java.io.IOException;
import javax.imageio.ImageIO;
import javax.swing.*;

/**
 *
 * @author Mohit-Ankit
 */
public class About_Dlg extends JDialog {

    public About_Dlg(JFrame parentFrame, String title) {
        super(parentFrame, title, true);

        //setting padding for the border layout
        int hGap = 10, vGap = 10;
        this.setLayout(new BorderLayout(hGap, vGap));

        //set window location relative to its parent window
        if (parentFrame != null) {
            Dimension parentSize = parentFrame.getSize();
            Point p = parentFrame.getLocation();
            setLocation(p.x + parentSize.width / 6, p.y + parentSize.height / 7);
        }

        JLabel lbl_image = new JLabel(new ImageIcon(Constants.LOGO_ABOUT));
        lbl_image.setMaximumSize(new Dimension(480, 200));
        lbl_image.setPreferredSize(new Dimension(480, 200));
        this.add(lbl_image, BorderLayout.NORTH);
        //this.add(new JLabel(new ImageIcon(Constants.LOGO_ABOUT)), BorderLayout.NORTH);

        String help_text = "AM Test Automation Framework. Version 1.0 2014-15 ";

        JTextArea jt_AboutText = new JTextArea(4, 20);//10, 15
        jt_AboutText.setText(help_text);
        jt_AboutText.setLineWrap(true);    // If text doesn't fit on a line, jump to the next
        jt_AboutText.setWrapStyleWord(true);   // Makes sure that words stay intact if a line wrap occurs
        jt_AboutText.setEditable(false);
        jt_AboutText.setOpaque(false);
        //jt_AboutText.setFont(new Font(), vGap, vGap));
        jt_AboutText.setFont(new Font("Times New Roman", Font.BOLD, 16));
        this.add(jt_AboutText, BorderLayout.SOUTH);

        this.setSize(500, 200);
        this.setDefaultCloseOperation(DISPOSE_ON_CLOSE);
        this.pack();
        this.setResizable(false);
        this.setVisible(true);
    }
}
