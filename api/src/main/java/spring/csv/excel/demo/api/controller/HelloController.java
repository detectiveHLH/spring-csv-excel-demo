package spring.csv.excel.demo.api.controller;

import org.springframework.http.MediaType;
import org.springframework.util.FileCopyUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import spring.csv.excel.demo.api.util.ExportUtil;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.*;

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
        List<LinkedHashMap<String, Object>> data = new ArrayList<>();
        body.put("1", "小明");
        body.put("2", "小王");
        data.add(header);
        data.add(body);
        data.add(body);
        data.add(body);
        FileCopyUtils.copy(ExportUtil.exportCSV(data), response.getOutputStream());
    }

    @GetMapping("xlsx")
    public void xlsx(
            HttpServletRequest request,
            HttpServletResponse response
    ) throws IOException {
        String fileName = this.getFileName(request, "测试数据.xlsx");
        response.setContentType(MediaType.APPLICATION_OCTET_STREAM.toString());
        response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\";");

        List<LinkedHashMap<String, Object>> datas = new ArrayList<>();
        LinkedHashMap<String, Object> data = new LinkedHashMap<>();
        data.put("1", "姓名");
        data.put("2", "年龄");
        datas.add(data);
        for (int i = 0; i < 5; i++) {
            data = new LinkedHashMap<>();
            data.put("1", "小青");
            data.put("2", "小白");
            datas.add(data);
        }

        Map<String, List<LinkedHashMap<String, Object>>> tableData = new HashMap<>();
        tableData.put("日报表", datas);
        tableData.put("周报表", datas);
        tableData.put("月报表", datas);

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
