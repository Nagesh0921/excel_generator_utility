package com.nm.excel.controller;

import com.nm.excel.service.ExcelService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

@RestController
public class ExcelController {
    private static final Logger LOG = LoggerFactory.getLogger(ExcelController.class);

    @Autowired
    private ExcelService service;

    @GetMapping("/download")
    public ResponseEntity<ByteArrayResource> download() throws IOException {
        LOG.info("Success");
        return service.download();
    }
}
