package util;

import bean.ClassBean;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelUtil {

    public static List<ClassBean> readMoreSheetFromXLS(String filePath) {
        List<ClassBean> classroomList = new ArrayList<ClassBean>();
        String cellStr = null;//单元格，最终按字符串处理
        //创建来自excel文件的输入流
        try {
            FileInputStream is = new FileInputStream(filePath);
            //创建WorkBook实例
            Workbook workbook = null;
            if (filePath.toLowerCase().endsWith("xls")) {//2003
                workbook = new HSSFWorkbook(is);
            } else if (filePath.toLowerCase().endsWith("xlsx")) {//2007
                workbook = WorkbookFactory.create(is);
            }
            //获取excel文件的sheet数量
            int numOfSheets = workbook.getNumberOfSheets();
            //挨个遍历sheet
            for (int i = 0; i < numOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                //挨个遍历sheet的每一行
                for (Iterator<Row> iterRow = sheet.iterator(); iterRow.hasNext(); ) {
                    Row row = iterRow.next();
                    ClassBean classroom = new ClassBean();
                    int j = 0;//标识位，用于标识第几列
                    //挨个遍历每一行的每一列
                    for (Iterator<Cell> cellIter = row.cellIterator(); cellIter.hasNext(); ) {
                        Cell cell = cellIter.next();//获取单元格对象
                        if (j == 0) {
                            if (cell == null) {// 单元格为空设置cellStr为空串
                                cellStr = "";
                            } else if (cell.getCellType() == CellType.BOOLEAN) {// 对布尔值的处理
                                cellStr = String.valueOf(cell.getBooleanCellValue());
                            } else if (cell.getCellType() == CellType.NUMERIC) {// 对数字值的处理
                                cellStr = cell.getNumericCellValue() + "";
                            } else {// 其余按照字符串处理
                                cellStr = cell.getStringCellValue();
                            }
                            classroom.setClassName(cellStr);
                            j++;
                        }
                        classroomList.add(classroom);
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return classroomList;
    }

    public static List<ClassBean> readFromXLSX2007(String filePath) {
        File excelFile = null;// Excel文件对象
        InputStream is = null;// 输入流对象
        String cellStr = null;// 单元格，最终按字符串处理
        List<ClassBean> classBeanList = new ArrayList<ClassBean>();// 返回封装数据的List
        ClassBean classBean = null;// 每一个雇员信息对象
        try {
            excelFile = new File(filePath);
            is = new FileInputStream(excelFile);// 获取文件输入流
//            XSSFWorkbook workbook2007 = new XSSFWorkbook(is);// 创建Excel2007文件对象
            org.apache.poi.ss.usermodel.Workbook workbook2007 = WorkbookFactory.create(is);
//            XSSFSheet sheet = workbook2007.getSheetAt(0);// 取出第一个工作表，索引是0
            org.apache.poi.ss.usermodel.Sheet sheet = workbook2007.getSheetAt(0);
            // 开始循环遍历行，表头不处理，从1开始
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                classBean = new ClassBean();// 实例化Student对象
//            	HSSFRow row = sheet.getRow(i);// 获取行对象
                Row row = sheet.getRow(i);// 获取行对象
                if (row == null) {// 如果为空，不处理
                    continue;
                }
                // 循环遍历单元格
                for (int j = 0; j < row.getLastCellNum(); j++) {
//                    XSSFCell cell = row.getCell(j);// 获取单元格对象
                    Cell cell = row.getCell(j);// 获取单元格对象
                    if (cell == null) {// 单元格为空设置cellStr为空串
                        cellStr = "";
                    } else if (cell.getCellType() == CellType.BOOLEAN) {// 对布尔值的处理
                        cellStr = String.valueOf(cell.getBooleanCellValue());
                    } else if (cell.getCellType() == CellType.NUMERIC) {// 对数字值的处理
                        cellStr = cell.getNumericCellValue() + "";
                    } else {// 其余按照字符串处理
                        cellStr = cell.getStringCellValue();
                    }
                    // 下面按照数据出现位置封装到bean中
                    if (j == 0) {
                        classBean.setCode(cellStr);
                    } else if (j == 1) {
                        classBean.setClassName(cellStr);
                    }else if(j == 7){
                        classBean.setAddress(cellStr);
                    }
                }
                classBeanList.add(classBean);// 数据装入List
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {// 关闭文件流
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return classBeanList;
    }
}
