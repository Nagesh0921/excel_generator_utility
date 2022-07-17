package com.nm.excel.service;

import com.nm.excel.bo.CVBBo;
import com.nm.excel.utiity.ContentCreator;
import com.nm.excel.utiity.ExcelGenerator;
import com.nm.excel.utiity.ExcelInput;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ConcurrentHashMap;

@Service
public class ExcelService {
    private static final Logger LOG = LoggerFactory.getLogger(ExcelService.class);

    public ResponseEntity<ByteArrayResource> download() throws IOException {
        LOG.info("ExcelService download :: invoked");
        ExcelGenerator generator = new ExcelGenerator();
        XSSFWorkbook wb = generator.generate();
        //Table Headers
        List<String> headers = new ArrayList<>();
        headers.add("ESN"); headers.add("Derate");headers.add("Temperature");headers.add("Amount");
        //Headers List
        //Trying to add data
        List<CVBBo> list = craeteBo(new ArrayList<>());
        ExcelInput input = new ExcelInput(list.size(), headers.size());
        ConcurrentHashMap<String, String> headersMap = new ConcurrentHashMap<>();
        headersMap.put("ESN", "1001");headersMap.put("Contract Name", "Indigo LEAP CFM Deal");headersMap.put("Engine", "LEAP");
        input.setHeadersMap(headersMap);
        input.setHeaders(headers);
        //Creating Data
        ContentCreator content = new ContentCreator();
        int[][] indices = content.createHeaderContent(wb.getSheetAt(0), input, 0, 0);
        content.createContent(wb.getSheetAt(0), input, indices[0][0], 2);

        Object[][] obj = new Object[list.size()][headers.size()];
        for(int r = 0; r < obj.length; r++){
            int c = 0;
            while(c < obj[0].length){
                if(c == 0){
                    obj[r][c] = list.get(r).getEsn();
                } else if (c == 1){
                    obj[r][c] = list.get(r).getDerate();
                } else if (c == 2){
                    obj[r][c] = list.get(r).getTemp();
                } else {
                    obj[r][c] = list.get(r).getAmt();
                }
                c++;
            }
        }
        input.setObj(obj);
        content.createData(wb.getSheetAt(0), input, indices[0][0]+4);
        //Covert To Byte Array Output Stream
        ByteArrayOutputStream stream = new ByteArrayOutputStream();
        wb.write(stream);
        stream.close();
        wb.close();
        //byte[] bytes = stream.toByteArray();
        HttpHeaders header = new HttpHeaders();
        header.setContentType(new MediaType("application", "force-download"));
        header.set(HttpHeaders.CONTENT_DISPOSITION, "attachment;filename="+"Test.xlsx");

        LOG.info("ExcelService download :: executed");
        return new ResponseEntity<ByteArrayResource>(new ByteArrayResource(stream.toByteArray()), header, HttpStatus.CREATED);
    }

    private List<CVBBo> craeteBo(List<CVBBo> list){
        for(int i=0; i<3; i++){
            CVBBo bo = new CVBBo();
            bo.setEsn("1000"+i);
            bo.setDerate(10D+i);
            bo.setTemp(35D+i);
            bo.setAmt(100D+(i*50));
            list.add(bo);
        }
        return list;
    }
}
