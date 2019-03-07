package spring.csv.excel.demo.api.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.util.FileCopyUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import spring.csv.excel.demo.api.util.ExportUtil;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

/**
 * HelloController
 *
 * @author Lunhao Hu
 * @date 2019-03-06 13:29
 **/
@RestController
public class HelloController {

    @GetMapping("csv")
    public void csv(
            HttpServletRequest request,
            HttpServletResponse response
    ) throws IOException {
        String fileName = this.getFileName(request, "测试数据.csv");
        response.setContentType(MediaType.APPLICATION_OCTET_STREAM.toString());
        response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\";");

        LinkedHashMap<String, Object> header = new LinkedHashMap<>();
        LinkedHashMap<String, Object> body = new LinkedHashMap<>();
        header.put("1", "姓名");
        header.put("2", "年龄");
        List<LinkedHashMap<String, Object>> bodys = new ArrayList<>();
        body.put("1", "小明");
        body.put("2", "小王");
        bodys.add(body);
        bodys.add(body);
        FileCopyUtils.copy(ExportUtil.exportCSV(header, bodys), response.getOutputStream());
    }

    @GetMapping("xlsx")
    public void xlsx(
            HttpServletRequest request,
            HttpServletResponse response
    ) throws IOException {
        String fileName = this.getFileName(request, "测试数据.xlsx");
        response.setContentType(MediaType.APPLICATION_OCTET_STREAM.toString());
        response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\";");

        List<LinkedHashMap<String, Object>> days = new ArrayList<>();
        LinkedHashMap<String, Object> day = new LinkedHashMap<>();
        day.put("1", "姓名");
        day.put("2", "年龄");
        days.add(day);
        for (int i = 0; i < 5; i++) {
            day = new LinkedHashMap<>();
            day.put("1", "小明");
            day.put("2", "小王");
            days.add(day);
        }

        List<LinkedHashMap<String, Object>> weeks = new ArrayList<>();
        LinkedHashMap<String, Object> week = new LinkedHashMap<>();
        week.put("1", "姓名");
        week.put("2", "年龄");
        weeks.add(week);
        for (int i = 0; i < 5; i++) {
            week = new LinkedHashMap<>();
            week.put("1", "小明");
            week.put("2", "小王");
            weeks.add(week);
        }

        List<LinkedHashMap<String, Object>> months = new ArrayList<>();
        LinkedHashMap<String, Object> month = new LinkedHashMap<>();
        month.put("1", "姓名");
        month.put("2", "年龄");
        months.add(month);
        for (int i = 0; i < 5; i++) {
            month = new LinkedHashMap<>();
            month.put("1", "小明");
            month.put("2", "小王");
            months.add(week);
        }

        LinkedHashMap<String, List<LinkedHashMap<String, Object>>> tableData = new LinkedHashMap<>();
        tableData.put("日报表", days);
        tableData.put("月报表", weeks);
        tableData.put("周报表", months);


        FileCopyUtils.copy(ExportUtil.exportXlsx(tableData), response.getOutputStream());
    }

    /**
     * 根据UA设置标题的编码，防止出现中文乱码
     *
     * @param request
     * @param name
     * @return
     * @throws UnsupportedEncodingException
     */
    private String getFileName(HttpServletRequest request, String name) throws UnsupportedEncodingException {
        String userAgent = request.getHeader("USER-AGENT");
        return userAgent.contains("Mozilla") ? new String(name.getBytes(), "ISO8859-1") : name;
    }

}
