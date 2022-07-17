package com.nm.excel.utiity;

import java.util.List;
import java.util.concurrent.ConcurrentHashMap;

public class ExcelInput {
    private List<String> headers;
    private ConcurrentHashMap<String, String> headersMap;
    private Object[][] obj;

    public ExcelInput(int rowSize, int colSize){
        obj = new Object[rowSize][colSize];
    }

    public List<String> getHeaders() {
        return headers;
    }

    public void setHeaders(List<String> headers) {
        this.headers = headers;
    }

    public ConcurrentHashMap<String, String> getHeadersMap() {
        return headersMap;
    }

    public void setHeadersMap(ConcurrentHashMap<String, String> headersMap) {
        this.headersMap = headersMap;
    }

    public Object[][] getObj() {
        return obj;
    }

    public void setObj(Object[][] obj) {
        this.obj = obj;
    }
}
