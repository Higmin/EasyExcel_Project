package com.zdhs.cms.common.excelUtils.model;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;

/**
 * @Auther : guojianmin
 * @Date : 2019/5/24 14:53
 * @Description : TODO用一句话描述此类的作用
 */
public class ExcelModel extends BaseRowModel {
    @ExcelProperty(index = 0 , value = "工号")
    private String staff_code;
    @ExcelProperty(index = 1 , value = "姓名")
    private String name;
    @ExcelProperty(index = 2 , value = "性别")
    private String sex;
    @ExcelProperty(index = 3 , value = "联系电话")
    private String tel;
    @ExcelProperty(index = 4 , value = "邮箱")
    private String email;
    @ExcelProperty(index = 5 , value = "微信号")
    private String weixin;
    @ExcelProperty(index = 6 , value = "部门")
    private String department;
    @ExcelProperty(index = 7 , value = "职位")
    private String position;

    public String getStaff_code() {
        return staff_code;
    }

    public void setStaff_code(String staff_code) {
        this.staff_code = staff_code;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    public String getTel() {
        return tel;
    }

    public void setTel(String tel) {
        this.tel = tel;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public String getWeixin() {
        return weixin;
    }

    public void setWeixin(String weixin) {
        this.weixin = weixin;
    }

    public String getDepartment() {
        return department;
    }

    public void setDepartment(String department) {
        this.department = department;
    }

    public String getPosition() {
        return position;
    }

    public void setPosition(String position) {
        this.position = position;
    }

    @Override
    public String toString() {
        return "ExcelModel{" +
                "staff_code='" + staff_code + '\'' +
                ", name='" + name + '\'' +
                ", sex='" + sex + '\'' +
                ", tel='" + tel + '\'' +
                ", email='" + email + '\'' +
                ", weixin='" + weixin + '\'' +
                ", department='" + department + '\'' +
                ", position='" + position + '\'' +
                '}';
    }
}
