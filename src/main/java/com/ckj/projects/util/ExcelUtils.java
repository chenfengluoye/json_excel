package com.ckj.projects.util;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.ckj.projects.entity.Entry;
import jxl.Cell;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import java.io.File;
import java.util.Collections;
import java.util.LinkedList;
import java.util.List;

/**
 * created by ChenKaiJu on 2018/11/1  18:26
 */
public class ExcelUtils {

    public static void writeObjectToNewExcel(JSONObject src,String filePath,boolean forceDel){
        File file=new File(filePath);
        if(!file.getName().endsWith(".xls")){
            throw new RuntimeException("仅支持xls格式的文件");
        }
        if(forceDel){
            try {
                WritableWorkbook writableWorkbook=Workbook.createWorkbook(file);
                WritableSheet sheet=writableWorkbook.createSheet("转换结果",0);
                List<Entry> linkedList= JSONUtils.jsonToSingle(src,"",new LinkedList());
                Collections.sort(linkedList);
                writeToSheet(linkedList,sheet);
                writableWorkbook.write();
                writableWorkbook.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    public static void writeArrayToNewExcel(JSONArray array, String filePath, boolean forceDel){
        File file=new File(filePath);
        if(!file.getName().endsWith(".xls")){
            throw new RuntimeException("仅支持xls格式的文件");
        }
        if(forceDel){
            try{
                if(array!=null){
                    WritableWorkbook writableWorkbook=Workbook.createWorkbook(file);
                    WritableSheet sheet=writableWorkbook.createSheet("转换结果",0);
                    for(int i=0;i<array.size();i++){
                        Object item=array.get(i);
                        try{
                            JSONObject jsonObject=JSON.parseObject(JSON.toJSONString(item));
                            List<Entry> jsonO=JSONUtils.jsonToSingle(jsonObject);
                            writeToSheet(jsonO,sheet,sheet.getRows());
                        }catch (Exception e){
                            try{
                                JSONArray jsonArray=JSON.parseArray(JSON.toJSONString(item));
                                List<Entry> jsonA=JSONUtils.arrayToSingle(jsonArray);
                                writeToSheet(jsonA,sheet,sheet.getRows());
                            }catch (Exception e1){
                                if(item.getClass().equals(String.class)){
                                    sheet.addCell(new Label(0,i, (String) item));
                                }else {
                                    sheet.addCell(new Label(0,i, JSON.toJSONString(item)));
                                }
                            }
                        }
                    }
                    writableWorkbook.write();
                    writableWorkbook.close();
                }
            }catch (Exception e){
                e.printStackTrace();
            }
        }
    }

    private static void writeToSheet(List<Entry> list,WritableSheet sheet,int beginRow){
        for(Entry entry:list){
            try {
                Cell[] rows=sheet.getRow(0);
                int index=indexOfCells(rows,entry.getKey());
                if(index<0){
                    index=sheet.getColumns();
                    sheet.addCell(new Label(index,0,entry.getKey()));
                }
                Cell[] columns=sheet.getColumn(index);
                int row=indexOfCells(columns,beginRow,"");
                if(row<0){
                    row=Math.max(columns.length,beginRow);
                }
                Object value=entry.getValue();
                if(String.class.equals(value.getClass())){
                    sheet.addCell(new Label(index,row,(String)value));
                }else {
                    sheet.addCell(new Label(index,row,JSON.toJSONString(value)));
                }
            }catch (Exception e){
                e.printStackTrace();
            }
        }
    }


    private static void writeToSheet(List<Entry> list,WritableSheet sheet){
        for(Entry entry:list){
            try {
                Cell[] rows=sheet.getRow(0);
                int index=indexOfCells(rows,entry.getKey());
                if(index<0){
                    index=sheet.getColumns();
                    sheet.addCell(new Label(index,0,entry.getKey()));
                }
                Cell[] columns=sheet.getColumn(index);
                int row=indexOfCells(columns,"");
                if(row<0){
                    row=columns.length;
                }
                Object value=entry.getValue();
                if(String.class.equals(value.getClass())){
                    sheet.addCell(new Label(index,row,(String)value));
                }else {
                    sheet.addCell(new Label(index,row,JSON.toJSONString(value)));
                }
            }catch (Exception e){
                e.printStackTrace();
            }
        }
    }

    public static int indexOfCells(Cell[] cells,int begin,String content){
        int re=-1;
        if(cells==null){
            return re;
        }
        if(StringUtils.isBlank(content)){
            content="";
        }
        for(int i=begin;i<cells.length;i++){
            try{
                Cell cell=cells[i];
                String t=cell.getContents();
                if(StringUtils.isBlank(t)){
                    t="";
                }
                if(content.trim().equals(t.trim())){
                    return i;
                }
            }catch (Exception e){
                e.printStackTrace();
            }
        }
        return re;
    }

    public static int indexOfCells(Cell[] cells,String content){
        int re=-1;
        if(cells==null){
            return re;
        }
        if(StringUtils.isBlank(content)){
            content="";
        }
        for(int i=0;i<cells.length;i++){
            try{
                Cell cell=cells[i];
                String t=cell.getContents();
                if(StringUtils.isBlank(t)){
                    t="";
                }
                if(content.trim().equals(t.trim())){
                    return i;
                }
            }catch (Exception e){
                e.printStackTrace();
            }
        }
        return re;
    }
}
