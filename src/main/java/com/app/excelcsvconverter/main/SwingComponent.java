/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.main;

import com.app.excelcsvconverter.resultmodel.ResultData;
import com.app.excelcsvconverter.util.MessageUtil;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Toolkit;
import java.util.concurrent.ExecutionException;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTabbedPane;
import javax.swing.JTextField;
import javax.swing.SwingWorker;
import javax.swing.UIManager;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileNameExtensionFilter;

/**
 * 功能描述：
 *
 * @since 2019-08-07
 */
public class SwingComponent {

    public static int CSV_TO_EXECL = 1;

    public static int EXECL_TO_CSV = 2;

    public static int INJECT_MECRO_TO_EXCEL = 3;

    public static int CUSTOM_PACKAGE_UPDATE = 4;

    public static int EXCEL_TO_STYLE = 5;

    private String customPath = "";

    private String language;

    private JFrame frame;

    public SwingComponent(String language) {
        this.language = language;
        this.frame = new JFrame(MessageUtil.getMessage("EXCEL_CSV_CONVERTER_EN", "EXCEL_CSV_CONVERTER_CN", language));
    }

    public void init(boolean isEnhance) {

        try {
            UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
        } catch (Exception e) {
            e.printStackTrace();
        }

        JPanel csvToExcelTab = createTab(
            MessageUtil.getMessage("EXCEL_TRANSFER_ZIP_EN", "EXCEL_TRANSFER_ZIP_CN", language), "zip", CSV_TO_EXECL,
            false, "zip");
        JPanel excelToCsvTab = null;
        if (isEnhance) {
            excelToCsvTab = createTab(
                MessageUtil.getMessage("CSV_TRANSFER_EXCEL_EN", "CSV_TRANSFER_EXCEL_CN", language), "excel",
                EXECL_TO_CSV, true, "xlsm", "xlsx", "csv");
        } else {
            excelToCsvTab = createTab(
                MessageUtil.getMessage("CSV_TRANSFER_EXCEL_EN", "CSV_TRANSFER_EXCEL_CN", language), "excel",
                EXECL_TO_CSV, false, "xlsm", "xlsx");
        }

        JPanel macroInjectTab = createTab(
            MessageUtil.getMessage("INJECT_MACRO_EXCEL_EN", "INJECT_MACRO_EXCEL_CN", language), "excel",
            INJECT_MECRO_TO_EXCEL, false, "xlsx");
        JPanel customPackageTab = createTab(
            MessageUtil.getMessage("CUSTOM_PACKAGE_FILE_EN", "CUSTOM_PACKAGE_FILE_CN", language), "ZipOrTarFile",
            CUSTOM_PACKAGE_UPDATE, false, "zip", "tar");

        // 存放选项卡的组件
        JTabbedPane tabs = new JTabbedPane(JTabbedPane.TOP);
        tabs.addTab(MessageUtil.getMessage("INJECT_MACRO_EN", "INJECT_MACRO_CN", language), null, macroInjectTab,
            "inject macro");
        tabs.addTab(MessageUtil.getMessage("CSV_TO_EXCEL_EN", "CSV_TO_EXCEL_CN", language), null, csvToExcelTab,
            "csv2excel");
        tabs.addTab(MessageUtil.getMessage("EXCEL_TO_CSV_EN", "EXCEL_TO_CSV_CN", language), null, excelToCsvTab,
            "excel2csv");
        tabs.addTab(MessageUtil.getMessage("CUSTOM_PACKAGE_EN", "CUSTOM_PACKAGE_CN", language), null, customPackageTab,
            "update custom package");

        frame.add(tabs, BorderLayout.CENTER);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setSize(500, 200);
        this.frame.setResizable(false);
        int windowWidth = frame.getWidth();
        int windowHeight = frame.getHeight();
        Toolkit kit = Toolkit.getDefaultToolkit();
        Dimension screenSize = kit.getScreenSize();
        int screenWidth = screenSize.width;
        int screenHeight = screenSize.height;
        frame.setLocation(screenWidth / 2 - windowWidth / 2, screenHeight / 2 - windowHeight / 2);
        frame.setVisible(true);
    }

    private JPanel createTab(String title, String FilterName, int opreationType, boolean isSupportDir,
        String... fileType) {
        JPanel tab = new JPanel();
        tab.setLayout(null);
        tab.setSize(500, 200);
        JPanel filePanel = new JPanel();

        filePanel.add(new JLabel(title));
        JTextField titleText = new JTextField(30);
        filePanel.add(titleText);
        JButton chooseFileBotton = new JButton("...", null);
        chooseFileBotton.addActionListener(e -> {
            JFileChooser chooser = new JFileChooser(".");
            if (isSupportDir) {
                chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
            } else {
                chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
            }
            FileNameExtensionFilter filter = new FileNameExtensionFilter(FilterName, fileType);
            chooser.setFileFilter(filter);
            int returnVal = chooser.showDialog(chooseFileBotton,
                MessageUtil.getMessage("CHOOSE_EN", "CHOOSE_CN", language));
            if (returnVal == JFileChooser.APPROVE_OPTION) {
                String filepath = chooser.getSelectedFile().getAbsolutePath();
                titleText.setText(filepath);
            }
        });
        filePanel.add(chooseFileBotton);
        filePanel.setBounds(-20, 35, 500, 50);
        JPanel submitPanel = new JPanel();
        JButton submitBotton = new JButton(MessageUtil.getMessage("SUBMIT_EN", "SUBMIT_CN", language), null);
        submitBotton.addActionListener(e -> {
            this.customPath = titleText.getText();
            //旋转等待显示
            final WaitUtil waitUtil = new WaitUtil();
            SwingWorker<ResultData, Void> sw = new SwingWorker<ResultData, Void>() {

                StringBuffer sb = new StringBuffer();

                @Override
                protected ResultData doInBackground() {
                    ResultData resultData = new Main().operation(customPath, opreationType);
                    return resultData;
                }

                @Override
                protected void done() {
                    if (waitUtil != null) {
                        waitUtil.dispose();
                    }
                    ResultData result = null;
                    try {
                        result = get();
                    } catch (InterruptedException ex) {
                        ex.printStackTrace();
                    } catch (ExecutionException ex) {
                        ex.printStackTrace();
                    }
                    if (result != null) {
                        JOptionPane.showMessageDialog(frame, result.getMessage(),
                            MessageUtil.getMessage("PROMPT_MESSAGE_EN", "PROMPT_MESSAGE_CN", language),
                            JOptionPane.INFORMATION_MESSAGE);
                    }

                }
            };
            sw.execute();
            waitUtil.setVisible(true);  //将旋转等待框WaitUnit设置为可见

        });
        submitPanel.add(submitBotton);
        tab.add(filePanel);
        submitPanel.setBounds(190, 90, 500, 50);
        tab.add(submitPanel);
        return tab;
    }

    public boolean confirmDialog(String title, String message) {
        int result = JOptionPane.showConfirmDialog(frame, message, title, JOptionPane.YES_NO_OPTION);
        if (result == 0) {
            return true;
        } else {
            return false;
        }
    }

    protected class WaitUtil extends JDialog {
        private static final long serialVersionUID = 6987303361741568128L;

        private final JPanel contentPanel = new JPanel();

        public WaitUtil() {
            if ("EN".equals(language)) {
                setBounds(0, 0, 230, 94);
            } else {
                setBounds(0, 0, 130, 94);
            }

            getContentPane().setLayout(new BorderLayout());
            contentPanel.setBorder(new EmptyBorder(5, 5, 5, 5));
            getContentPane().add(contentPanel, BorderLayout.CENTER);
            contentPanel.setLayout(null);
            {
                JLabel lblLoading = new JLabel(
                    MessageUtil.getMessage("WAITING_MESSAGE_EN", "WAITING_MESSAGE_CN", language));
                lblLoading.setForeground(Color.DARK_GRAY);
                lblLoading.setOpaque(false);
                lblLoading.setIcon(new ImageIcon(
                    Toolkit.getDefaultToolkit().getImage(SwingComponent.class.getResource("/icon/waiting.gif"))));
                lblLoading.setFont(new Font("宋体", Font.PLAIN, 20));
                if ("EN".equals(language)) {
                    lblLoading.setBounds(0, 0, 230, 94);
                } else {
                    lblLoading.setBounds(0, 0, 130, 94);
                }

                contentPanel.add(lblLoading);
            }

            setModalityType(ModalityType.APPLICATION_MODAL);
            setUndecorated(true);
            setLocationRelativeTo(frame);
        }
    }

}