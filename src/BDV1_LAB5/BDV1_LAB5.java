package BDV1_LAB5;

import java.awt.Cursor;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class BDV1_LAB5 extends javax.swing.JFrame {
    private static final long serialVersionUID = 1L;

    class TThread1 extends Thread {

        public void run() {
            String dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
                    + System.getProperty("file.separator");
            
            // Чтение из файла-шаблона в переменную doc
            HWPFDocument doc = null;
            try (FileInputStream fis = new FileInputStream(dir + "post_template.doc")) {
                doc = new HWPFDocument(fis);
                fis.close();
            } catch (Exception ex) {
                System.err.println("Error template!");
            }

            // Замена в переменной doc данных
            try {
                doc.getRange().replaceText("$ИИНотправителя", jTextField_IINsender.getText());
                doc.getRange().replaceText("$ФИОотправителя", jTextField_FIOsender.getText());
                doc.getRange().replaceText("$адрес", jTextField_Adres.getText());
                doc.getRange().replaceText("$Индекс", jTextField_Index.getText());
                doc.getRange().replaceText("$Номертелефона", jTextField_Phone.getText());
                doc.getRange().replaceText("$ФИОполучателя", jTextField_FIOrecepient.getText());
                doc.getRange().replaceText("$ИИНполучателя", jTextField_IINrecepient.getText());
                doc.getRange().replaceText("$Год", jTextField_Year.getText());
            } catch (Exception ex) {
                System.err.println("Error replaceText!");
            }

            // Сохранение переменной doc в новый файл
            try (FileOutputStream fos = new FileOutputStream(dir + "post.doc")) {
                doc.write(fos);
                fos.close();
                // Открытие файла внешней программой
                if (System.getProperty("os.name").equals("Linux")
                        && System.getProperty("java.vendor").startsWith("Red Hat")) {
                    new ProcessBuilder("xdg-open", dir + "post.doc").start();
                } else {
                    Desktop.getDesktop().open(new File(dir + "post.doc"));
                }
            } catch (Exception ex) {
                System.err.println("Error getDesktop!");
            }
            setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
           
        }
        
    }

    class TThread2 extends Thread {

        
        public void run() {
            String dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
                    + System.getProperty("file.separator");
            
            // Чтение из файла-шаблона в переменную docx
            XWPFDocument doc = null;
            
           try (FileInputStream fis = new FileInputStream(dir + "post_template.docx")) {
                doc = new XWPFDocument(fis);
                fis.close();
            } catch (Exception ex) {
                System.err.println("Error template!");
            }

            // Замена в переменной doc данных
            
            for (XWPFTable tbl : doc.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            if (text != null && text.contains("$ИИНотправителя")) {
                                text = text.replace(text, jTextField_IINsender.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$ФИОотправителя")) {
                                text = text.replace(text, jTextField_FIOsender.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$адрес")) {
                                text = text.replace(text, jTextField_Adres.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$Индекс")) {
                                text = text.replace(text, jTextField_Index.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$Номертелефона")) {
                                text = text.replace(text, jTextField_Phone.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$ФИОполучателя")) {
                                text = text.replace(text, jTextField_FIOrecepient.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$ИИНполучателя")) {
                                text = text.replace(text, jTextField_IINrecepient.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$Год")) {
                                text = text.replace(text, jTextField_Year.getText());
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }

            // Сохранение переменной docx в новый файл
            
                        try (FileOutputStream fos = new FileOutputStream(dir + "post.docx")) {
                doc.write(fos);
                fos.close();
                // Открытие файла внешней программой
                if (System.getProperty("os.name").equals("Linux")
                        && System.getProperty("java.vendor").startsWith("Red Hat")) {
                    new ProcessBuilder("xdg-open", dir + "post.docx").start();
                } else {
                    Desktop.getDesktop().open(new File(dir + "post.docx"));
                }
            } catch (Exception ex) {
                System.err.println("Error getDesktop!");
            }
            setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
        }
        
    }
    
    
    public BDV1_LAB5() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jButton_Save = new javax.swing.JButton();
        jButton_Save1 = new javax.swing.JButton();
        jTextField_IINsender = new javax.swing.JTextField();
        jTextField_FIOsender = new javax.swing.JTextField();
        jTextField_Adres = new javax.swing.JTextField();
        jTextField_Index = new javax.swing.JTextField();
        jTextField_Phone = new javax.swing.JTextField();
        jTextField_FIOrecepient = new javax.swing.JTextField();
        jTextField_IINrecepient = new javax.swing.JTextField();
        jTextField_Year = new javax.swing.JTextField();
        jLabel1 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Квитанция в MS Word");
        setResizable(false);
        getContentPane().setLayout(null);

        jButton_Save.setText("в DOC");
        jButton_Save.setToolTipText("");
        jButton_Save.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton_SaveActionPerformed(evt);
            }
        });
        getContentPane().add(jButton_Save);
        jButton_Save.setBounds(850, 380, 80, 21);

        jButton_Save1.setText("в DOCX");
        jButton_Save1.setToolTipText("");
        jButton_Save1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton_Save1ActionPerformed(evt);
            }
        });
        getContentPane().add(jButton_Save1);
        jButton_Save1.setBounds(950, 380, 80, 21);

        jTextField_IINsender.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        jTextField_IINsender.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_IINsenderActionPerformed(evt);
            }
        });
        getContentPane().add(jTextField_IINsender);
        jTextField_IINsender.setBounds(310, 80, 260, 24);

        jTextField_FIOsender.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        jTextField_FIOsender.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_FIOsenderActionPerformed(evt);
            }
        });
        getContentPane().add(jTextField_FIOsender);
        jTextField_FIOsender.setBounds(700, 80, 260, 30);

        jTextField_Adres.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Adres);
        jTextField_Adres.setBounds(310, 124, 260, 23);

        jTextField_Index.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        jTextField_Index.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_IndexActionPerformed(evt);
            }
        });
        getContentPane().add(jTextField_Index);
        jTextField_Index.setBounds(810, 120, 210, 30);

        jTextField_Phone.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        jTextField_Phone.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_PhoneActionPerformed(evt);
            }
        });
        getContentPane().add(jTextField_Phone);
        jTextField_Phone.setBounds(700, 170, 270, 30);

        jTextField_FIOrecepient.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        jTextField_FIOrecepient.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_FIOrecepientActionPerformed(evt);
            }
        });
        getContentPane().add(jTextField_FIOrecepient);
        jTextField_FIOrecepient.setBounds(630, 233, 270, 30);

        jTextField_IINrecepient.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        jTextField_IINrecepient.setActionCommand("<Not Set>");
        jTextField_IINrecepient.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_IINrecepientActionPerformed(evt);
            }
        });
        getContentPane().add(jTextField_IINrecepient);
        jTextField_IINrecepient.setBounds(630, 267, 270, 30);

        jTextField_Year.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        jTextField_Year.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_YearActionPerformed(evt);
            }
        });
        getContentPane().add(jTextField_Year);
        jTextField_Year.setBounds(970, 330, 30, 23);

        jLabel1.setIcon(new javax.swing.ImageIcon(getClass().getResource("/BDV1_LAB5/post.png"))); // NOI18N
        getContentPane().add(jLabel1);
        jLabel1.setBounds(10, 10, 1040, 390);

        setSize(new java.awt.Dimension(1062, 446));
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void jButton_SaveActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton_SaveActionPerformed
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        new TThread1().start();
    }//GEN-LAST:event_jButton_SaveActionPerformed

    private void jTextField_IINsenderActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_IINsenderActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField_IINsenderActionPerformed

    private void jTextField_IndexActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_IndexActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField_IndexActionPerformed

    private void jTextField_FIOsenderActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_FIOsenderActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField_FIOsenderActionPerformed

    private void jButton_Save1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton_Save1ActionPerformed
          setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        new TThread2().start();
    }//GEN-LAST:event_jButton_Save1ActionPerformed

    private void jTextField_IINrecepientActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_IINrecepientActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField_IINrecepientActionPerformed

    private void jTextField_FIOrecepientActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_FIOrecepientActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField_FIOrecepientActionPerformed

    private void jTextField_PhoneActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_PhoneActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField_PhoneActionPerformed

    private void jTextField_YearActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_YearActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField_YearActionPerformed

    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(BDV1_LAB5.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(BDV1_LAB5.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(BDV1_LAB5.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(BDV1_LAB5.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new BDV1_LAB5().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton_Save;
    private javax.swing.JButton jButton_Save1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JTextField jTextField_Adres;
    private javax.swing.JTextField jTextField_FIOrecepient;
    private javax.swing.JTextField jTextField_FIOsender;
    private javax.swing.JTextField jTextField_IINrecepient;
    private javax.swing.JTextField jTextField_IINsender;
    private javax.swing.JTextField jTextField_Index;
    private javax.swing.JTextField jTextField_Phone;
    private javax.swing.JTextField jTextField_Year;
    // End of variables declaration//GEN-END:variables
}
