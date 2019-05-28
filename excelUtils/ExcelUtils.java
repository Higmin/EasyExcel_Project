package com.zdhs.cms.common.excelUtils;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.zdhs.cms.common.excelUtils.listen.AfterWriteHandlerImpl;
import com.zdhs.cms.common.excelUtils.model.ExcelModel;
import com.zdhs.cms.common.excelUtils.listen.ExcelListener;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @Auther : guojianmin
 * @Date : 2019/5/24 15:12
 * @Description : excel工具类
 */
public class ExcelUtils {
    /**
     * @param is   导入文件输入流
     * @param clazz Excel实体映射类
     * @return
     */
    public static Boolean readExcel(InputStream is, Class clazz){

        BufferedInputStream bis = null;
        try {
            bis = new BufferedInputStream(is);
            // 解析每行结果在listener中处理
            AnalysisEventListener listener = new ExcelListener();
            ExcelReader excelReader = new ExcelReader(bis, ExcelTypeEnum.XLS, null, listener);
            excelReader.read(new Sheet(1, 2, clazz));
            is.close();
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        } finally {
            if (bis != null) {
                try {
                    bis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return true;
    }

    /**
     *
     * @param os 文件输出流
     * @param clazz Excel实体映射类
     * @param data 导出数据
     * @return
     */
    public static Boolean writeExcel(OutputStream os, Class clazz, List<? extends BaseRowModel> data){
        BufferedOutputStream bos= null;
        try {
            InputStream inputStream = FileUtil.getResourcesFileInputStream("excelTemplate/temp.xlsx");
            bos = new BufferedOutputStream(os);
            ExcelWriter writer = EasyExcelFactory.getWriterWithTempAndHandler(inputStream, bos, ExcelTypeEnum.XLSX, false,
                    new AfterWriteHandlerImpl());

            //写第一个sheet, sheet1  数据全是List<String>
            Sheet sheet1 = new Sheet(1, 0,clazz);
            writer.write(data, sheet1);
            writer.finish();
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        } finally {
            if (bos != null) {
                try {
                    bos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return true;
    }

    public static void main(String[] args) {
        //1.读Excel
//        FileInputStream fis = null;
//        try {
//            fis = new FileInputStream("D:\\2007.xlsx");
//            Boolean flag = ExcelUtils.readExcel(fis, ExcelModel.class);
//            System.out.println("导入是否成功："+flag);
//        } catch (FileNotFoundException e) {
//            e.printStackTrace();
//        }finally {
//            if (fis != null){
//                try {
//                    fis.close();
//                } catch (IOException e) {
//                    e.printStackTrace();
//                }
//            }
//        }


        //2.写Excel
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream("D:\\export.xlsx");
            //FileOutputStream fos, Class clazz, List<? extends BaseRowModel> data
            List<ExcelModel> list = new ArrayList<>();
            for (int i = 0; i < 5; i++){
                ExcelModel excelModel = new ExcelModel();
                excelModel.setName("我是名字"+i);
                list.add(excelModel);
            }
            Boolean flag = ExcelUtils.writeExcel(fos, ExcelModel.class,list);
            System.out.println("导出是否成功："+flag);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }finally {
            if (fos != null){
                try {
                    fos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}
