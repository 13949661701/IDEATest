package com.sqp;


import com.sqp.utils.ExcelWriter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("poi")
public class PoiController {

    @RequestMapping(value = "exportToExcel", method = RequestMethod.GET)
    public void exportToExcel(HttpServletRequest request, HttpServletResponse response,
                              String observe_day) throws Exception {
//        List<CsRank> list = csRankService.queryCsRankByDay(observe_day);
        List<CsRank> list = new ArrayList<>();
        String[] titles = {"日期","风险分组","风险等级","人数"};

        String filename = "催收模型统计信息详情";
        int colsNum = 4;
        int rowIndex = 0;

        response.reset();
        response.setHeader("Content-Disposition", "attachment;filename=" + new String((filename + ".xls").getBytes(), "iso8859-1"));
        ServletOutputStream os = response.getOutputStream();

        // 创建sheet
        ExcelWriter ew = new ExcelWriter();
        HSSFWorkbook wb = ew.getWorkBook();
        wb.setSheetName(0, filename);

        // 默认设置宽高
        ew.setDefaultColumnWidth((short) (25));
        ew.setDefaultRowHeight(20);
        ew.addTitle(rowIndex++, filename, 0, colsNum - 1);

        // 标题列
        ew.addRow(rowIndex++, 0, titles, ew.getExcelStyle().getTitleStyle());

        // 内容列
        for (int i = 0; list != null && i < list.size(); i++) {
            CsRank entity = list.get(i);
            Object[] tvs = new Object[colsNum];
            tvs[0] = entity.getObserve_day();
            tvs[1] = entity.getGroup();
            tvs[2] = entity.getCs_rank();
            tvs[3] = entity.getCs_count();
            ew.addRow(rowIndex++, 0, tvs, ew.getExcelStyle().getNormalStyle());
        }
        ew.output(os);
        os.flush();
        os.close();

    }


}
