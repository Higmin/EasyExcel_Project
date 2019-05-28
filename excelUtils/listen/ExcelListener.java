package com.zdhs.cms.common.excelUtils.listen;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.zdhs.cms.common.excelUtils.model.ExcelModel;
import com.zdhs.cms.entity.Staff;
import com.zdhs.cms.quartz.util.SpringUtils;
import com.zdhs.cms.service.impl.DepartmentServiceImpl;
import com.zdhs.cms.service.impl.PositionServiceImpl;
import com.zdhs.cms.service.impl.StaffServiceImpl;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 *
 * @Auther : guojianmin
 * @Date : 2019/5/20 15:06
 * @Description : 解析监听器，
 *   每解析一行会回调invoke()方法。
 *   整个excel解析结束会执行doAfterAllAnalysed()方法
 *
 *   下面只是我写的一个样例而已，可以根据自己的逻辑修改该类。
 *
 */
public class ExcelListener extends AnalysisEventListener {

    //自定义用于暂时存储data。
    //可以通过实例获取该值
    private List<Object> datas = new ArrayList<Object>();

    public void invoke(Object object, AnalysisContext context) {
        System.out.println("当前行："+context.getCurrentRowNum());
        System.out.println(object);
        datas.add(object);//数据存储到list，供批量处理，或后续自己业务逻辑处理。
        doSomething(object);//根据自己业务做处理
    }
    private void doSomething(Object object) {
        ExcelModel excel = (ExcelModel) object;
        //1、入库调用接口
        System.out.println("导入完成 ！ 数据入库==>");
        Staff staff = ModelToEntity(excel);
        StaffServiceImpl staffService = SpringUtils.getBean(StaffServiceImpl.class);
        Map<String, Object> map = staffService.addStaff(staff);
        System.out.println("此处为：入库调用接口------------->结果 1 为成功， 0为失败！结果为："+map.get("code"));

    }
    public void doAfterAllAnalysed(AnalysisContext context) {
         datas.clear();//解析结束销毁不用的资源
    }
    public List<Object> getDatas() {
        return datas;
    }
    public void setDatas(List<Object> datas) {
        this.datas = datas;
    }
    public Staff ModelToEntity(ExcelModel excelModel){
        Staff staff = new Staff();
        staff.setStaff_code(excelModel.getStaff_code());
        staff.setName(excelModel.getName());
        staff.setSex(excelModel.getSex());
        staff.setTel(excelModel.getTel());
        staff.setEmail(excelModel.getEmail());
        staff.setWeixin(excelModel.getWeixin());
        DepartmentServiceImpl departmentService = SpringUtils.getBean(DepartmentServiceImpl.class);
        Integer departID = departmentService.queryIdByName(excelModel.getDepartment());
        System.out.println("departID ==> "+departID);
        staff.setDepartment_id(departID);
        PositionServiceImpl positionService = SpringUtils.getBean(PositionServiceImpl.class);
        Integer positionID = positionService.queryIdByName(excelModel.getPosition());
        System.out.println("positionID ==> "+positionID);
        staff.setPosition_id(positionID);
        return staff;
    }
}