package com.plan.domain;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.sql.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class LoadData {

    private static final String DB_DRIVER = "org.apache.derby.jdbc.EmbeddedDriver";
    private static final String DB_CONNECTION = "jdbc:derby:BASE";
    private static Connection conn = null;
    private static PreparedStatement preparedStatement = null;

	public void load() throws IOException, InvalidFormatException, SQLException {

        createConnection();

		String workingDir = System.getProperty("user.dir");
		File file = new File(workingDir + "/MaterialCustomer.xls");
        FileInputStream fs = new FileInputStream(file);
		Workbook workbook = WorkbookFactory.create(fs);
        Sheet datatypeSheet = workbook.getSheetAt(0);
        int lastRow = datatypeSheet.getLastRowNum();
        System.out.println("-= Total row " + lastRow +" =-");
        int totalData = 0;
        for(int i = 1; i <= lastRow; i++) {
            Row currentRow = datatypeSheet.getRow(i);

            String materialFoamtec = null;
            Cell cell1 = currentRow.getCell(0);
            if(cell1 != null) {
                if(cell1.getCellType() == Cell.CELL_TYPE_STRING) {
                    materialFoamtec = cell1.getStringCellValue();
                } else {
                    materialFoamtec = "" + cell1.getNumericCellValue();
                }
            }

            String materialCustomer = null;
            Cell cell2 = currentRow.getCell(1);
            if(cell2 != null) {
                if(cell2.getCellType() == Cell.CELL_TYPE_STRING) {
                    materialCustomer = cell2.getStringCellValue();
                } else {
                    cell2.setCellType(Cell.CELL_TYPE_STRING);
                    materialCustomer = cell2.getStringCellValue();
                }
            }

            String materialGroup = null;
            Cell cell3 = currentRow.getCell(2);
            if(cell3 != null) {
                if(cell3.getCellType() == Cell.CELL_TYPE_STRING) {
                    materialGroup = cell3.getStringCellValue();
                } else {
                    materialGroup = "" + cell2.getNumericCellValue();
                }
            }

            MaterialPlanning materialPlanning = new MaterialPlanning();
            materialPlanning.setId((long)i);
            materialPlanning.setMaterialFoamtec(materialFoamtec);
            materialPlanning.setMaterialCustomer(materialCustomer);
            materialPlanning.setMaterialGroup(materialGroup);

            create(materialPlanning);

            System.out.println("-= row " + i + " | " + materialFoamtec + " | " + materialCustomer + " | " + materialGroup + " =-");
            totalData = totalData + 1;
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

	public static void create(MaterialPlanning materialPlanning) throws SQLException {
	    String sql = "insert into MaterialPlanning (id, materialFoamtec, materialCustomer, materialGroup) values (?,?,?,?)";
        preparedStatement = conn.prepareStatement(sql);
        preparedStatement.setLong(1, materialPlanning.getId());
        preparedStatement.setString(2, materialPlanning.getMaterialFoamtec());
        preparedStatement.setString(3, materialPlanning.getMaterialCustomer());
        preparedStatement.setString(4, materialPlanning.getMaterialGroup());
        preparedStatement.executeUpdate();
	}
}