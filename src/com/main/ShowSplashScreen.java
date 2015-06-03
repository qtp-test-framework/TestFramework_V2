package com.main;

import com.mail.MailSettings_Dlg;
import com.util.Constants;
import java.awt.*;
import java.awt.geom.Rectangle2D;
import javax.swing.UIManager;

/**
 * Example for Splash Screen tutorial
 *
 * @author Joseph Areeda
 */
public class ShowSplashScreen {

    static SplashScreen mySplash;                   // instantiated by JVM we use it to get graphics
    static Graphics2D splashGraphics;               // graphics context for overlay of the splash image
    //static Rectangle2D.Double splashTextArea;       // area where we draw the text
    static Rectangle2D.Double splashProgressArea;   // area where we draw the progress bar
    static Font font;                               // used to draw our text

    public static void main(String[] args) {
        try {
            //Get System (Microsoft Windows) Look and Feel
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(MainWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(MainWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(MainWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(MainWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }

        EventQueue.invokeLater(new Runnable() {

            public void run() {
                try {
                    splashInit();           // initialize splash overlay drawing parameters
                    appInit();              // simulate what an application would do before starting
                    if (mySplash != null){ // check if we really had a spash screen
                        mySplash.close();   // we're done with it
                    }

                } catch (Exception ex) {
                    System.out.println("Exception in Splash Screen code : " + ex);
                }
                
                //Display wth main window
                new MainWindow(Constants.APPLICATION_NAME);
            }
        });
    }

    /**
     * just a stub to simulate a long initialization task that updates the text
     * and progress parts of the status in the Splash
     */
    private static void appInit() throws Exception {
        for (int i = 1; i <= 20; i++) {   // pretend we have 10 things to do
            int pctDone = i * 5;       // this is about the only time I could calculate rather than guess progress
//            splashText("Doing task #" + i);     // tell the user what initialization task is being done
            
            //As MailSettings dialog takes more the 2 secs to load on first execution
            //so loading the class in splash screen
            if(i==18){
                new MailSettings_Dlg(null, "Mail Settings", false);
            }
            
            splashProgress(pctDone);            // give them an idea how much we have completed
            try {
                Thread.sleep(20);             // wait a second
            } catch (InterruptedException ex) {
                break;
            }
        }
    }

    /**
     * Prepare the global variables for the other splash functions
     */
    private static void splashInit() throws Exception {
        // the splash screen object is created by the JVM, if it is displaying a splash image
        //path to image is specidied in the manifest file and also in Properties>>Run  -splash:src/Resources/Images/splash.png
        mySplash = SplashScreen.getSplashScreen();

        // if there are any problems displaying the splash image
        // the call to getSplashScreen will returned null

        if (mySplash != null) {

            // get the size of the image now being displayed
            Dimension ssDim = mySplash.getSize();
            int height = ssDim.height;
            int width = ssDim.width;

            // stake out some area for our status information
//            splashTextArea = new Rectangle2D.Double(15., height * 0.88, width * .45, 32.);
            splashProgressArea = new Rectangle2D.Double(width * .55, height * .92, width * .4, 12);

            // create the Graphics environment for drawing status info
            splashGraphics = mySplash.createGraphics();
            font = new Font("Dialog", Font.PLAIN, 14);
            splashGraphics.setFont(font);

            // initialize the status info
//            splashText("Starting");
            splashProgress(0);
        }
    }

    /**
     * Display text in status area of Splash. Note: no validation it will fit.
     *
     * @param str - text to be displayed
     */
//    public static void splashText(String str) throws Exception {
//        if (mySplash != null && mySplash.isVisible()) {   // important to check here so no other methods need to know if there
//            // really is a Splash being displayed
//
//            // erase the last status text
//            splashGraphics.setPaint(Color.LIGHT_GRAY);
//            splashGraphics.fill(splashTextArea);
//
//            // draw the text
//            splashGraphics.setPaint(Color.BLACK);
//            splashGraphics.drawString(str, (int) (splashTextArea.getX() + 10), (int) (splashTextArea.getY() + 15));
//
//            // make sure it's displayed
//            mySplash.update();
//        }
//    }
    /**
     * Display a (very) basic progress bar
     *
     * @param pct how much of the progress bar to display 0-100
     */
    public static void splashProgress(int pct) throws Exception {
        if (mySplash != null && mySplash.isVisible()) {

            // Note: 3 colors are used here to demonstrate steps
            // erase the old one
            splashGraphics.setPaint(Color.LIGHT_GRAY);
            splashGraphics.fill(splashProgressArea);

            // draw an outline
            splashGraphics.setPaint(Color.BLUE);
            splashGraphics.draw(splashProgressArea);

            // Calculate the width corresponding to the correct percentage
            int x = (int) splashProgressArea.getMinX();
            int y = (int) splashProgressArea.getMinY();
            int wid = (int) splashProgressArea.getWidth();
            int hgt = (int) splashProgressArea.getHeight();

            int doneWidth = Math.round(pct * wid / 100.f);
            doneWidth = Math.max(0, Math.min(doneWidth, wid - 1));  // limit 0-width

            // fill the done part one pixel smaller than the outline
            splashGraphics.setPaint(Color.GREEN);
            splashGraphics.fillRect(x, y + 1, doneWidth, hgt - 1);

            // make sure it's displayed
            mySplash.update();
        }
    }
}
