package com.hp.test;

import com.hp.entity.ExcelEntityField;
import com.hp.entity.ExportCaseInfo;
import com.hp.entity.ImportCaseInfo;
import com.hp.entity.ImportDemoCaseInfo;
import com.hp.utils.POIUtil;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * @Classname TestList2workbook
 * @Description 把List数据转为workbook列表
 * @Date 2021/11/17 13:55
 * @Created by huangwencai
 */
public class TestList2workbook {

    public static void main(String[] args) {
        //内存List对象数据解析为workbook
        List2Workbook();

        //客户上传文件解析为List对象
       // analyseWorkbook2List();

        //客户上传文件解析为List对象
      //  analyseDemoWorkbook2List();


    }

    private static void List2Workbook() {
        List<ExportCaseInfo> exportCaseInfos =new ArrayList<>();
        exportCaseInfos.add(new ExportCaseInfo("A","张三","专业催员"));
        exportCaseInfos.add(new ExportCaseInfo("B","李四","专业催员"));
        exportCaseInfos.add(new ExportCaseInfo("C","王五","专业催员"));
        XSSFWorkbook xSSFWorkbook= POIUtil.analyseList2workbook(ExportCaseInfo.class, exportCaseInfos,"催员报表");
        System.out.println(xSSFWorkbook);
    }

    private static void analyseWorkbook2List() {
        XSSFWorkbook xssfWorkbook = null;
        try {
            xssfWorkbook = new XSSFWorkbook("E://11.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }
        List<Object> list=POIUtil.analyseWorkbook2List(xssfWorkbook, ImportCaseInfo.class);
        if(CollectionUtils.isNotEmpty(list)){
            List<ImportCaseInfo>importCaseInfos=new ArrayList<>();
            Iterator<Object> iterator= list.iterator();
            while (iterator.hasNext()){
                importCaseInfos.add ((ImportCaseInfo)iterator.next());
            }
            System.out.println(importCaseInfos);
        }else {
            System.out.println("列表无数据");
        }
    }

    private static void analyseDemoWorkbook2List() {
        XSSFWorkbook xssfWorkbook = null;
        try {
            xssfWorkbook = new XSSFWorkbook("E://20211116.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }
        List<Object> list=POIUtil.analyseWorkbook2List(xssfWorkbook, ImportDemoCaseInfo.class);
        if(CollectionUtils.isNotEmpty(list)){
            List<ImportDemoCaseInfo>importCaseInfos=new ArrayList<>();
            Iterator<Object> iterator= list.iterator();
            while (iterator.hasNext()){
                importCaseInfos.add ((ImportDemoCaseInfo)iterator.next());
            }
            System.out.println(importCaseInfos);
        }else {
            System.out.println("列表无数据");
        }
    }


}
