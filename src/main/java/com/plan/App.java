package com.plan;

import javax.swing.*;
import java.awt.*;
import java.awt.Font;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.awt.event.WindowEvent;
import java.io.*;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.Desktop;
import java.awt.event.WindowAdapter;
import java.sql.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Date;
import java.util.List;

import com.plan.domain.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class App {

    private static final String DB_DRIVER = "org.apache.derby.jdbc.EmbeddedDriver";
    private static final String DB_CONNECTION = "jdbc:derby:BASE";
    private static Connection conn = null;
    private static PreparedStatement preparedStatement = null;

	private JFrame frmConvertPlanFoamtec;

	private JFrame frmConvertCelestica;

    private JFrame frmConvertDelta;

	public static void main( String[] args ) throws IOException, InvalidFormatException, SQLException {

		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					App window = new App();
					window.frmConvertPlanFoamtec.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
    }

	public App() {
		initialize();
	}

	private void initialize() {
		frmConvertPlanFoamtec = new JFrame();
		frmConvertPlanFoamtec.setTitle("Convert plan foamtec");
		frmConvertPlanFoamtec.setBounds(100, 100, 600, 400);
		frmConvertPlanFoamtec.setResizable(false);
		frmConvertPlanFoamtec.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmConvertPlanFoamtec.getContentPane().setLayout(null);
		frmConvertPlanFoamtec.setLocationRelativeTo(null);

		JPanel panel = new JPanel();
		panel.setBounds(6, 6, 588, 366);
		frmConvertPlanFoamtec.getContentPane().add(panel);
		panel.setLayout(null);

		JButton btnNewButton = new JButton("CELESTICA");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				celesticaWindows();
				frmConvertPlanFoamtec.setVisible(false);
			}
		});
        btnNewButton.setBounds(31, 35, 164, 56);
		panel.add(btnNewButton);

        JButton button = new JButton("DELTA");
        button.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                deltaWindows();
                frmConvertPlanFoamtec.setVisible(false);
            }
        });
        button.setBounds(207, 35, 164, 56);
        panel.add(button);

        JButton button_1 = new JButton("N/A");
        button_1.setBounds(384, 35, 164, 56);
        panel.add(button_1);
	}

	public void celesticaWindows() {
		frmConvertCelestica = new JFrame();
		frmConvertCelestica.setTitle("Convert celestica");
		frmConvertCelestica.setBounds(100, 100, 600, 400);
		frmConvertCelestica.setResizable(false);
		frmConvertCelestica.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmConvertCelestica.getContentPane().setLayout(null);
		frmConvertCelestica.setLocationRelativeTo(null);

		final JPanel panel = new JPanel();
		panel.setBounds(6, 6, 588, 366);
		frmConvertCelestica.getContentPane().add(panel);
		panel.setLayout(null);

		JLabel lblInputFileExcel = new JLabel("Input File Excel :");
		lblInputFileExcel.setFont(new Font("Lucida Grande", Font.PLAIN, 15));
		lblInputFileExcel.setBounds(53, 72, 146, 36);
		panel.add(lblInputFileExcel);

		JButton btnOpenFile = new JButton("Choose file");
		btnOpenFile.setBounds(222, 78, 117, 29);

		final JLabel lblNewLabel = new JLabel("Result");
		lblNewLabel.setBounds(175, 139, 342, 16);
		panel.add(lblNewLabel);

		final JButton btnConvert = new JButton("Convert");
		btnConvert.setBounds(430, 301, 117, 29);

		final JFileChooser choose = new JFileChooser();
		FileFilter filter = new FileNameExtensionFilter("xlsx/xls file", "xlsx", "xls");
		choose.addChoosableFileFilter(filter);

		btnConvert.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
                createConnection();
				File file = choose.getSelectedFile();
		        try {
                    InputStream inputStream = new FileInputStream(file);
                    Workbook workbook = WorkbookFactory.create(inputStream);
                    Sheet datatypeSheet = workbook.getSheetAt(0);
                    int lastRow = datatypeSheet.getLastRowNum();
                    int totalData = 0;
                    SimpleDateFormat formatter = new SimpleDateFormat("MM/dd/yyyy");

                    List<MaterialConvert> materialConvertList = new ArrayList<MaterialConvert>();
                    String celesticaPart = null;
                    String startDateStr = null;
                    double qty = 0;
                    Set<String> partCustomerSet = new HashSet<String>();
                    for(int i = 1; i <= lastRow; i++) {
                        Row currentRow = datatypeSheet.getRow(i);
                        Cell cell1 = currentRow.getCell(1);
                        Cell cell2 = currentRow.getCell(3);
                        Cell cell3 = currentRow.getCell(5);
                        if(cell1.getCellType() != Cell.CELL_TYPE_BLANK && cell2.getCellType() != Cell.CELL_TYPE_BLANK && cell3.getCellType() != Cell.CELL_TYPE_BLANK) {
                            System.out.println("cell1 = " + cell1 + ", cell2 = " + cell2 + ", cell3 = " + cell3);
                            celesticaPart = cell1.getStringCellValue();
                            if(cell2.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                Date stDate = cell2.getDateCellValue();
                                startDateStr = formatter.format(stDate);
                            } else if(cell2.getCellType() == Cell.CELL_TYPE_STRING) {
                                startDateStr = cell2.getStringCellValue();
                            }

                            qty = cell3.getNumericCellValue();
                            partCustomerSet.add(celesticaPart);

                            Date startDate = null;
                            try {
                                startDate = formatter.parse(startDateStr);
                            } catch (ParseException ep) {
                                ep.printStackTrace();
                            }
                            Calendar cal = Calendar.getInstance();
                            cal.setTime(startDate);
                            int month = cal.get(Calendar.MONTH);

                            MaterialConvert materialConvert = new MaterialConvert();
                            materialConvert.setCustomerPart(celesticaPart);
                            materialConvert.setMonth(theMonth(month));
                            materialConvert.setQty((int) qty);
                            materialConvertList.add(materialConvert);
                            totalData = totalData + 1;
                        }
                    }
                    Map<String,MaterialPlanning> partFoamtecMap = new HashMap<String,MaterialPlanning>();
                    for(String s : partCustomerSet) {
                        partFoamtecMap.put(s,findByCustomerPart(s));
                    }

                    List<MaterialSuccess> materialSuccessList = new ArrayList<MaterialSuccess>();
                    for(String s : partCustomerSet) {
                        int total1 = 0;
                        int total2 = 0;
                        int total3 = 0;
                        int total4 = 0;
                        int total5 = 0;
                        int total6 = 0;
                        int total7 = 0;
                        int total8 = 0;
                        int total9 = 0;
                        int total10 = 0;
                        int total11 = 0;
                        int total12 = 0;

                        MaterialSuccess materialSuccess = new MaterialSuccess();
                        materialSuccess.setCustomerPart(s);
                        for(MaterialConvert mc : materialConvertList) {
                            if(mc.getCustomerPart().equals(s)) {

                                if(partFoamtecMap.get(s) == null) {
                                    materialSuccess.setFoamtecPart("N/A");
                                } else {
                                    materialSuccess.setFoamtecPart(partFoamtecMap.get(s).getMaterialFoamtec());
                                }

                                if(mc.getMonth().equals("January")) {
                                    total1 = total1 + mc.getQty();
                                }
                                if(mc.getMonth().equals("February")) {
                                    total2 = total2 + mc.getQty();
                                }
                                if(mc.getMonth().equals("March")) {
                                    total3 = total3 + mc.getQty();
                                }
                                if(mc.getMonth().equals("April")) {
                                    total4 = total4 + mc.getQty();
                                }
                                if(mc.getMonth().equals("May")) {
                                    total5 = total5 + mc.getQty();
                                }
                                if(mc.getMonth().equals("June")) {
                                    total6 = total6 + mc.getQty();
                                }
                                if(mc.getMonth().equals("July")) {
                                    total7 = total7 + mc.getQty();
                                }
                                if(mc.getMonth().equals("August")) {
                                    total8 = total8 + mc.getQty();
                                }
                                if(mc.getMonth().equals("September")) {
                                    total9 = total9 + mc.getQty();
                                }
                                if(mc.getMonth().equals("October")) {
                                    total10 = total10 + mc.getQty();
                                }
                                if(mc.getMonth().equals("November")) {
                                    total11 = total11 + mc.getQty();
                                }
                                if(mc.getMonth().equals("December")) {
                                    total12 = total12 + mc.getQty();
                                }
                            }
                        }
                        materialSuccess.setJanuary(total1);
                        materialSuccess.setFebruary(total2);
                        materialSuccess.setMarch(total3);
                        materialSuccess.setApril(total4);
                        materialSuccess.setMay(total5);
                        materialSuccess.setJune(total6);
                        materialSuccess.setJuly(total7);
                        materialSuccess.setAugust(total8);
                        materialSuccess.setSeptember(total9);
                        materialSuccess.setOctober(total10);
                        materialSuccess.setNovember(total11);
                        materialSuccess.setDecember(total12);

                        materialSuccessList.add(materialSuccess);
                    }

                    for(MaterialSuccess materialSuccess : materialSuccessList) {
                        System.out.println("Part : " + materialSuccess.getCustomerPart() + ", Code SAP : " + materialSuccess.getFoamtecPart() +
                                "|" + materialSuccess.getJanuary() + "|" +
                                "|" + materialSuccess.getFebruary() + "|" +
                                "|" + materialSuccess.getMarch() + "|" +
                                "|" + materialSuccess.getApril() + "|" +
                                "|" + materialSuccess.getMay() + "|" +
                                "|" + materialSuccess.getJune() + "|" +
                                "|" + materialSuccess.getJuly() + "|" +
                                "|" + materialSuccess.getAugust() + "|" +
                                "|" + materialSuccess.getSeptember() + "|" +
                                "|" + materialSuccess.getOctober() + "|" +
                                "|" + materialSuccess.getNovember() + "|" +
                                "|" + materialSuccess.getDecember() + "|"
                        );
                    }

                    XSSFWorkbook wb = new XSSFWorkbook();

                    XSSFCellStyle style = wb.createCellStyle();
                    style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                    Sheet sheet = wb.createSheet();
                    Row row1 = sheet.createRow(0);

                    Cell cell1 = row1.createCell(0);
                    cell1.setCellStyle(style);
                    cell1.setCellValue("PART");

                    Cell cell2 = row1.createCell(1);
                    cell2.setCellStyle(style);
                    cell2.setCellValue("CODE SAP");

                    Cell cell3 = row1.createCell(2);
                    cell3.setCellStyle(style);
                    cell3.setCellValue("January");

                    Cell cell4 = row1.createCell(3);
                    cell4.setCellStyle(style);
                    cell4.setCellValue("February");

                    Cell cell5 = row1.createCell(4);
                    cell5.setCellStyle(style);
                    cell5.setCellValue("March");

                    Cell cell6 = row1.createCell(5);
                    cell6.setCellStyle(style);
                    cell6.setCellValue("April");

                    Cell cell7 = row1.createCell(6);
                    cell7.setCellStyle(style);
                    cell7.setCellValue("May");

                    Cell cell8 = row1.createCell(7);
                    cell8.setCellStyle(style);
                    cell8.setCellValue("June");

                    Cell cell9 = row1.createCell(8);
                    cell9.setCellStyle(style);
                    cell9.setCellValue("July");

                    Cell cell10 = row1.createCell(9);
                    cell10.setCellStyle(style);
                    cell10.setCellValue("August");

                    Cell cell11 = row1.createCell(10);
                    cell11.setCellStyle(style);
                    cell11.setCellValue("September");

                    Cell cell12 = row1.createCell(11);
                    cell12.setCellStyle(style);
                    cell12.setCellValue("October");

                    Cell cell13 = row1.createCell(12);
                    cell13.setCellStyle(style);
                    cell13.setCellValue("November");

                    Cell cell14 = row1.createCell(13);
                    cell14.setCellStyle(style);
                    cell14.setCellValue("December");

                    int rowIndex = 1;
                    for(MaterialSuccess materialSuccess : materialSuccessList) {
                        Row row = sheet.createRow(rowIndex);
                        row.createCell(0).setCellValue(materialSuccess.getCustomerPart());
                        row.createCell(1).setCellValue(materialSuccess.getFoamtecPart());
                        row.createCell(2).setCellValue(materialSuccess.getJanuary());
                        row.createCell(3).setCellValue(materialSuccess.getFebruary());
                        row.createCell(4).setCellValue(materialSuccess.getMarch());
                        row.createCell(5).setCellValue(materialSuccess.getApril());
                        row.createCell(6).setCellValue(materialSuccess.getMay());
                        row.createCell(7).setCellValue(materialSuccess.getJune());
                        row.createCell(8).setCellValue(materialSuccess.getJuly());
                        row.createCell(9).setCellValue(materialSuccess.getAugust());
                        row.createCell(10).setCellValue(materialSuccess.getSeptember());
                        row.createCell(11).setCellValue(materialSuccess.getOctober());
                        row.createCell(12).setCellValue(materialSuccess.getNovember());
                        row.createCell(13).setCellValue(materialSuccess.getDecember());
                        rowIndex = rowIndex + 2;
                    }

                    for (int i=0; i < 14; i++){
                        sheet.autoSizeColumn(i);
                    }

                    String workingDir = System.getProperty("user.home") + "/convertCelestica/";
                    File convFile = new File(workingDir + "celesticaConvert.xlsx");
                    convFile.getParentFile().mkdirs();
                    FileOutputStream fos = new FileOutputStream(convFile);
                    wb.write(fos);
                    fos.close();
					Desktop.getDesktop().open(convFile);
                    conn.close();
				} catch (Exception e1) {
                    JOptionPane.showMessageDialog(frmConvertCelestica,"Format is not convert.", "Error", JOptionPane.ERROR_MESSAGE);
                    try {
                        conn.close();
                    } catch (SQLException e2) {
                        e2.printStackTrace();
                    }
                    e1.printStackTrace();
				}
            }
		});

		btnOpenFile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				int rVal = choose.showOpenDialog(frmConvertCelestica);
				if (rVal == JFileChooser.APPROVE_OPTION) {
					lblNewLabel.setText(choose.getSelectedFile().toString());
				}
			}
		});

		panel.add(btnOpenFile);
        panel.add(btnConvert);

		frmConvertCelestica.setVisible(true);
	}

    public void deltaWindows() {
        frmConvertDelta = new JFrame();
        frmConvertDelta.setTitle("Convert delta");
        frmConvertDelta.setBounds(100, 100, 600, 400);
        frmConvertDelta.setResizable(false);
        frmConvertDelta.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frmConvertDelta.getContentPane().setLayout(null);
        frmConvertDelta.setLocationRelativeTo(null);

        final JPanel panel = new JPanel();
        panel.setBounds(6, 6, 588, 366);
        frmConvertDelta.getContentPane().add(panel);
        panel.setLayout(null);

        JLabel lblInputFileExcel = new JLabel("Input File Excel :");
        lblInputFileExcel.setFont(new Font("Lucida Grande", Font.PLAIN, 15));
        lblInputFileExcel.setBounds(53, 72, 146, 36);
        panel.add(lblInputFileExcel);

        JButton btnOpenFile = new JButton("Choose file");
        btnOpenFile.setBounds(222, 78, 117, 29);

        final JLabel lblNewLabel = new JLabel("Result");
        lblNewLabel.setBounds(175, 139, 342, 16);
        panel.add(lblNewLabel);

        final JButton btnConvert = new JButton("Convert");
        btnConvert.setBounds(430, 301, 117, 29);

        final JFileChooser choose = new JFileChooser();
        FileFilter filter = new FileNameExtensionFilter("xlsx/xls file", "xlsx", "xls");
        choose.addChoosableFileFilter(filter);

        btnConvert.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                createConnection();
                File file = choose.getSelectedFile();
                try {
                    InputStream inputStream = new FileInputStream(file);
                    Workbook workbook = WorkbookFactory.create(inputStream);
                    Sheet datatypeSheet = workbook.getSheetAt(0);
                    int lastRow = datatypeSheet.getLastRowNum();
                    int totalData = 0;

                    List<MaterialDeltaPlaning> materialDeltaPlaningList = new ArrayList<MaterialDeltaPlaning>();

                    for(int i = 1; i <= lastRow; i++) {
                        Row firstRow = datatypeSheet.getRow(0);
                        Row currentRow = datatypeSheet.getRow(i);

                        String partNo = "";
                        int future = 0;
                        int total = 0;
                        int M1 = 0;
                        int M2 = 0;
                        int M3 = 0;
                        int M4 = 0;
                        int M5 = 0;
                        int M6 = 0;
                        int M7 = 0;
                        int M8 = 0;
                        int M9 = 0;
                        int M10 = 0;
                        int M11 = 0;
                        int M12 = 0;

                        for(int j = 1; j < currentRow.getLastCellNum(); j++) {

                            Cell cellData = currentRow.getCell(j);

                            if(firstRow.getCell(j).getCellType() == Cell.CELL_TYPE_STRING) {
                                if(firstRow.getCell(j).getStringCellValue().indexOf("PART NO") >= 0) {
                                    if(cellData.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                        cellData.setCellType(Cell.CELL_TYPE_STRING);
                                    }
                                    if(cellData.getCellType() == Cell.CELL_TYPE_STRING) {
                                        partNo = cellData.getStringCellValue();
                                    }
                                }
                                if(firstRow.getCell(j).getStringCellValue().indexOf("FUTURE") >= 0) {
                                    future = future + (int)cellData.getNumericCellValue();
                                }
                            } else {
                                if(firstRow.getCell(j).getNumericCellValue() == 1) {
                                    M1 = M1 + (int)cellData.getNumericCellValue();
                                }
                                if(firstRow.getCell(j).getNumericCellValue() == 2) {
                                    M2 = M2 + (int)cellData.getNumericCellValue();
                                }
                                if(firstRow.getCell(j).getNumericCellValue() == 3) {
                                    M3 = M3 + (int)cellData.getNumericCellValue();
                                }
                                if(firstRow.getCell(j).getNumericCellValue() == 4) {
                                    M4 = M4 + (int)cellData.getNumericCellValue();
                                }
                                if(firstRow.getCell(j).getNumericCellValue() == 5) {
                                    M5 = M5 + (int)cellData.getNumericCellValue();
                                }
                                if(firstRow.getCell(j).getNumericCellValue() == 6) {
                                    M6 = M6 + (int)cellData.getNumericCellValue();
                                }
                                if(firstRow.getCell(j).getNumericCellValue() == 7) {
                                    M7 = M7 + (int)cellData.getNumericCellValue();
                                }
                                if(firstRow.getCell(j).getNumericCellValue() == 8) {
                                    M8 = M8 + (int)cellData.getNumericCellValue();
                                }
                                if(firstRow.getCell(j).getNumericCellValue() == 9) {
                                    M9 = M9 + (int)cellData.getNumericCellValue();
                                }
                                if(firstRow.getCell(j).getNumericCellValue() == 10) {
                                    M10 = M10 + (int)cellData.getNumericCellValue();
                                }
                                if(firstRow.getCell(j).getNumericCellValue() == 11) {
                                    M11 = M11 + (int)cellData.getNumericCellValue();
                                }
                                if(firstRow.getCell(j).getNumericCellValue() == 12) {
                                    M12 = M12 + (int)cellData.getNumericCellValue();
                                }
                            }
                        }

                        MaterialDeltaPlaning materialDeltaPlaning = new MaterialDeltaPlaning();
                        materialDeltaPlaning.setPartNumber(partNo);
                        materialDeltaPlaning.setM1(M1);
                        materialDeltaPlaning.setM2(M2);
                        materialDeltaPlaning.setM3(M3);
                        materialDeltaPlaning.setM4(M4);
                        materialDeltaPlaning.setM5(M5);
                        materialDeltaPlaning.setM6(M6);
                        materialDeltaPlaning.setM7(M7);
                        materialDeltaPlaning.setM8(M8);
                        materialDeltaPlaning.setM9(M9);
                        materialDeltaPlaning.setM10(M10);
                        materialDeltaPlaning.setM11(M11);
                        materialDeltaPlaning.setM12(M12);
                        materialDeltaPlaning.setFuture(future);
                        materialDeltaPlaning.setTotal(M1+M2+M3+M4+M5+M6+M7+M8+M9+M10+M11+M12+future);
                        materialDeltaPlaningList.add(materialDeltaPlaning);
                    }

                    List<MaterialSuccess> materialSuccessList = new ArrayList<MaterialSuccess>();
                    for(MaterialDeltaPlaning ma : materialDeltaPlaningList) {
                        MaterialSuccess materialSuccess = new MaterialSuccess();
                        MaterialPlanning materialPlanning = findByCustomerPart(ma.getPartNumber());

                        if(materialPlanning == null) {
                            materialSuccess.setFoamtecPart("N/A");
                        } else {
                            materialSuccess.setFoamtecPart(materialPlanning.getMaterialFoamtec());
                        }
                        materialSuccess.setCustomerPart(ma.getPartNumber());

                        materialSuccess.setJanuary(ma.getM1());
                        materialSuccess.setFebruary(ma.getM2());
                        materialSuccess.setMarch(ma.getM3());
                        materialSuccess.setApril(ma.getM4());
                        materialSuccess.setMay(ma.getM5());
                        materialSuccess.setJune(ma.getM6());
                        materialSuccess.setJuly(ma.getM7());
                        materialSuccess.setAugust(ma.getM8());
                        materialSuccess.setSeptember(ma.getM9());
                        materialSuccess.setOctober(ma.getM10());
                        materialSuccess.setNovember(ma.getM11());
                        materialSuccess.setDecember(ma.getM12());
                        materialSuccess.setFuture(ma.getFuture());
                        materialSuccess.setTotal(ma.getTotal());

                        materialSuccessList.add(materialSuccess);
                    }

                    for(MaterialSuccess materialSuccess: materialSuccessList) {
                        System.out.println(materialSuccess.toString());
                    }

                    XSSFWorkbook wb = new XSSFWorkbook();

                    XSSFCellStyle style = wb.createCellStyle();
                    style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                    Sheet sheet = wb.createSheet();
                    Row row1 = sheet.createRow(0);

                    Cell cell1 = row1.createCell(0);
                    cell1.setCellStyle(style);
                    cell1.setCellValue("PART");

                    Cell cell2 = row1.createCell(1);
                    cell2.setCellStyle(style);
                    cell2.setCellValue("CODE SAP");

                    Cell cell3 = row1.createCell(2);
                    cell3.setCellStyle(style);
                    cell3.setCellValue("January");

                    Cell cell4 = row1.createCell(3);
                    cell4.setCellStyle(style);
                    cell4.setCellValue("February");

                    Cell cell5 = row1.createCell(4);
                    cell5.setCellStyle(style);
                    cell5.setCellValue("March");

                    Cell cell6 = row1.createCell(5);
                    cell6.setCellStyle(style);
                    cell6.setCellValue("April");

                    Cell cell7 = row1.createCell(6);
                    cell7.setCellStyle(style);
                    cell7.setCellValue("May");

                    Cell cell8 = row1.createCell(7);
                    cell8.setCellStyle(style);
                    cell8.setCellValue("June");

                    Cell cell9 = row1.createCell(8);
                    cell9.setCellStyle(style);
                    cell9.setCellValue("July");

                    Cell cell10 = row1.createCell(9);
                    cell10.setCellStyle(style);
                    cell10.setCellValue("August");

                    Cell cell11 = row1.createCell(10);
                    cell11.setCellStyle(style);
                    cell11.setCellValue("September");

                    Cell cell12 = row1.createCell(11);
                    cell12.setCellStyle(style);
                    cell12.setCellValue("October");

                    Cell cell13 = row1.createCell(12);
                    cell13.setCellStyle(style);
                    cell13.setCellValue("November");

                    Cell cell14 = row1.createCell(13);
                    cell14.setCellStyle(style);
                    cell14.setCellValue("December");

                    Cell cell15 = row1.createCell(14);
                    cell15.setCellStyle(style);
                    cell15.setCellValue("Future");

                    Cell cell16 = row1.createCell(15);
                    cell16.setCellStyle(style);
                    cell16.setCellValue("Total");

                    int rowIndex = 1;
                    for(MaterialSuccess materialSuccess : materialSuccessList) {
                        Row row = sheet.createRow(rowIndex);
                        row.createCell(0).setCellValue(materialSuccess.getCustomerPart());
                        row.createCell(1).setCellValue(materialSuccess.getFoamtecPart());
                        row.createCell(2).setCellValue(materialSuccess.getJanuary());
                        row.createCell(3).setCellValue(materialSuccess.getFebruary());
                        row.createCell(4).setCellValue(materialSuccess.getMarch());
                        row.createCell(5).setCellValue(materialSuccess.getApril());
                        row.createCell(6).setCellValue(materialSuccess.getMay());
                        row.createCell(7).setCellValue(materialSuccess.getJune());
                        row.createCell(8).setCellValue(materialSuccess.getJuly());
                        row.createCell(9).setCellValue(materialSuccess.getAugust());
                        row.createCell(10).setCellValue(materialSuccess.getSeptember());
                        row.createCell(11).setCellValue(materialSuccess.getOctober());
                        row.createCell(12).setCellValue(materialSuccess.getNovember());
                        row.createCell(13).setCellValue(materialSuccess.getDecember());
                        row.createCell(14).setCellValue(materialSuccess.getFuture());
                        row.createCell(15).setCellValue(materialSuccess.getTotal());
                        rowIndex = rowIndex + 2;
                    }

                    for (int i=0; i < 14; i++){
                        sheet.autoSizeColumn(i);
                    }

                    String workingDir = System.getProperty("user.home") + "/convertDelta/";
                    File convFile = new File(workingDir + "deltaConvert.xlsx");
                    convFile.getParentFile().mkdirs();
                    FileOutputStream fos = new FileOutputStream(convFile);
                    wb.write(fos);
                    fos.close();
                    Desktop.getDesktop().open(convFile);
                    conn.close();
                } catch (Exception e1) {
                    JOptionPane.showMessageDialog(frmConvertDelta,"Format is not convert.", "Error", JOptionPane.ERROR_MESSAGE);
                    try {
                        conn.close();
                    } catch (SQLException e2) {
                        e2.printStackTrace();
                    }
                    e1.printStackTrace();
                }
            }
        });

        btnOpenFile.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                int rVal = choose.showOpenDialog(frmConvertDelta);
                if (rVal == JFileChooser.APPROVE_OPTION) {
                    lblNewLabel.setText(choose.getSelectedFile().toString());
                }
            }
        });

        panel.add(btnOpenFile);
        panel.add(btnConvert);

        frmConvertDelta.setVisible(true);
    }

    public String theMonth(int month){
        String[] monthNames = {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"};
        return monthNames[month];
    }

    public MaterialPlanning findByCustomerPart(String codeCustomer) throws SQLException {
        String selectSQL = "SELECT id, materialFoamtec, materialCustomer, materialGroup FROM MaterialPlanning WHERE materialCustomer = ?";
        preparedStatement = conn.prepareStatement(selectSQL);
        preparedStatement.setString(1, codeCustomer);
        ResultSet rs = preparedStatement.executeQuery();

        List<MaterialPlanning> materialPlannings = new ArrayList<MaterialPlanning>();
        while (rs.next()) {
            Long id  = rs.getLong("id");
            String materialFoamtec = rs.getString("materialFoamtec");
            String materialCustomer = rs.getString("materialCustomer");
            String materialGroup = rs.getString("materialGroup");

            MaterialPlanning materialPlanning = new MaterialPlanning();

            System.out.print("id : " + id);
            System.out.print(" | materialFoamtec : " + materialFoamtec);
            System.out.print(" | materialCustomer : " + materialCustomer);
            System.out.println(" | materialGroup : " + materialGroup);
            materialPlanning.setId(id);
            materialPlanning.setMaterialFoamtec(materialFoamtec);
            materialPlanning.setMaterialCustomer(materialCustomer);
            materialPlanning.setMaterialGroup(materialGroup);
            materialPlannings.add(materialPlanning);
        }

        if(materialPlannings.size() > 1) {
            MaterialPlanning materialPlanning = null;
            for(MaterialPlanning m : materialPlannings) {
                String lastCharm = m.getMaterialFoamtec();
                String s = lastCharm.substring(lastCharm.length() - 1);
                if(s.equals("B")) {
                    materialPlanning = m;
                }
            }
            return materialPlanning;
        } else if(materialPlannings.size() == 1) {
            return materialPlannings.get(0);
        } else {
            return null;
        }

    }

    public static void createConnection() {
        try {
            Class.forName(DB_DRIVER).newInstance();
            //Get a connection
            conn = DriverManager.getConnection(DB_CONNECTION);
            System.out.println("ok");
        }
        catch (Exception except) {
            except.printStackTrace();
        }
    }
}
