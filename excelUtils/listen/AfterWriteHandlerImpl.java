package com.zdhs.cms.common.excelUtils.listen;

import com.alibaba.excel.event.WriteHandler;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

/**
 *
 * @Auther : guojianmin
 * @Date : 2019/5/27 08:12
 * @Description : 自定义样式
 */
public class AfterWriteHandlerImpl implements WriteHandler {

    CellStyle cellStyle;

    @Override
    public void sheet(int sheetNo, Sheet sheet) {
        Workbook workbook = sheet.getWorkbook();
        //要锁定单元格需先为此表单设置保护密码，设置之后此表单默认为所有单元格锁定，可使用setLocked(false)为指定单元格设置不锁定。
        //设置表单保护密码
        sheet.protectSheet("your password");
        //创建样式
        cellStyle = workbook.createCellStyle();
        //创建字体
        Font font = workbook.createFont();
        //设置是否锁
        cellStyle.setLocked(false);
        // 水平对齐方式
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        // 垂直对齐方式
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置字体
        font.setFontName("楷体");
        //设置字体大小
        font.setFontHeightInPoints((short) 12);
        cellStyle.setFont(font);
    }

    @Override
    public void row(int rowNum, Row row) {
        Workbook workbook = row.getSheet().getWorkbook();
        //设置行高
        row.setHeight((short) 300);
    }

    @Override
    public void cell(int cellNum, Cell cell) {
        Workbook workbook = cell.getSheet().getWorkbook();
        Sheet currentSheet = cell.getSheet();
        if (cell.getRowIndex() > 0) {
            cell.setCellStyle(cellStyle);
        }
    }
}
