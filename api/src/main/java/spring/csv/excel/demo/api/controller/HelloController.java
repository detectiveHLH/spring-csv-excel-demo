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
    public void test(
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
