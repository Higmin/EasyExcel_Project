package com.zdhs.cms.controller;

import com.zdhs.cms.common.excelUtils.ExcelUtils;
import com.zdhs.cms.common.excelUtils.model.ExcelModel;
import com.zdhs.cms.service.DepartmentService;
import com.zdhs.cms.service.PositionService;
import com.zdhs.cms.service.StaffService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.Resource;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.util.List;

/**
 * @Auther : guojianmin
 * @Date : 2019/5/24 17:03
 * @Description : TODO用一句话描述此类的作用
 */
@Controller
@RequestMapping("/excel")
public class ExcelDownloadController {

    private static final Logger log = LoggerFactory.getLogger(ExcelDownloadController.class);

    @Resource
    StaffService staffService;

    /**
     * 单个文件上传
     * @param file
     * @return
     */
    @RequestMapping(value = "/upload")
    public String upload(@RequestParam("file") MultipartFile file) {
        try {
            if (file.isEmpty()) {
                return "文件为空";
            }
            // 获取文件名
            String fileName = file.getOriginalFilename();
            log.info("上传的文件名为：" + fileName);
            // 获取文件的后缀名
            String suffixName = fileName.substring(fileName.lastIndexOf("."));
            log.info("文件的后缀名为：" + suffixName);
            InputStream inputStream = file.getInputStream();
            ExcelUtils.readExcel(inputStream, ExcelModel.class);
            return "上传成功";
        } catch (IllegalStateException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return "上传失败";
    }

    /**
     * 下载模板，用于填写导入数据
     * @param request
     * @param response
     */
    @RequestMapping("/downloadExcel")
    public void cooperation(HttpServletRequest request, HttpServletResponse response) {
        ServletOutputStream out = null;
        try {
            out = response.getOutputStream();
            response.setContentType("multipart/form-data");
            response.setCharacterEncoding("utf-8");
            String fileName = "导入模板";
            response.setHeader("Content-disposition", "attachment;filename="+ URLEncoder.encode(fileName,"UTF-8")+".xlsx");
            ExcelUtils.writeExcel(out,ExcelModel.class,null);
            out.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            if (out != null){
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * 下载模板，导入数据文件
     * @param request
     * @param response
     */
    @RequestMapping("/downloadExcelData")
    public void cooperationData(HttpServletRequest request, HttpServletResponse response) {
        ServletOutputStream out = null;
        try {
            out = response.getOutputStream();
            response.setContentType("multipart/form-data");
            response.setCharacterEncoding("utf-8");
            String fileName = "导出明细";
            response.setHeader("Content-disposition", "attachment;filename="+ URLEncoder.encode(fileName,"UTF-8")+".xlsx");

            List<ExcelModel> data = staffService.selectAll();
            //把数据明细放在list data中
            System.out.println("把数据明细放在list data中:请完善查询数据接口调用，并把查询结果写入list data中");
            Boolean flag = ExcelUtils.writeExcel(out, ExcelModel.class, data);
            System.out.println("导出是否成功："+flag);
            out.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            if (out != null){
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}
