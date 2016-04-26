package com.file.controller;

import java.io.IOException;
import java.util.List;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletResponse;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import com.file.forms.Computer;
import com.file.service.PoiService;
/**
 * 测试导入导出
 * @author jinhang
 *
 */
@Controller  
@RequestMapping("/report") 
public class PoiController {

	@Resource(name = "poiService")  
    private PoiService service;  
  
    /** 
     * 跳转到主页. 
     */  
    @RequestMapping(value = "", method = RequestMethod.GET)  
    public String getIndex() {  
        
        return "report";  
    }  
  
    /** 
     * 导出excel报表 
     */  
    @RequestMapping(value = "/export", method = RequestMethod.GET)  
    public void getXLS(HttpServletResponse response) {  
        service.exportXLS(response);  
    }  
  
    /** 
     * 读取excel报表 
     */  
    @RequestMapping(value = "/read", method = RequestMethod.POST)  
    public String getReadReport(@RequestParam MultipartFile file)  
            throws IOException {  
        List<Computer> list = service.readReport(file.getInputStream());  
        service.insertComputer(list);  
        return "addedReport";  
  
    }  
}
