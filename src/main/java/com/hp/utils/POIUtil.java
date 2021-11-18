package com.hp.utils;

import com.hp.anotation.MineExcelEntity;
import com.hp.anotation.MineExcelHeaderAnnotation;
import com.hp.anotation.MineExcelProperty;
import com.hp.entity.ExcelEntityField;
import com.hp.entity.ExcelHeader;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;


/**
 * @Classname POIUtil
 * @Description POI工具类
 * @Date 2021/11/17 13:56
 * @Created by huangwencai
 */
public class POIUtil {

    /**
     * 把传入的List数据集合封装为workbook表格
     * @param classType 实体类的class类型
     * @param objectList      实体类集合
     * @param sheetName  sheet表格名字
     * @return
     */
    public static XSSFWorkbook analyseList2workbook(Class<?> classType, List<?> objectList, String sheetName){
        XSSFWorkbook xssfWorkbook=new XSSFWorkbook();
        XSSFSheet xssfSheet=xssfWorkbook.createSheet(sheetName);
        //获取实体类字段，找到有注解的字段，按照排序顺序，生成第一行表头
        Field[] fields= classType.getDeclaredFields();
        List<Field> filedList=Arrays.asList(fields);
        //leftFieldsList  是含有注解的所有字段
        List<Field> leftFieldsList=filedList.stream().filter(field -> {
            MineExcelHeaderAnnotation mineExcelHeaderAnnotation =field.getAnnotation(MineExcelHeaderAnnotation.class);
            if(mineExcelHeaderAnnotation ==null){
                return false;
            }else {
                return true;
            }
        }).collect(Collectors.toList());
        List<ExcelHeader> excelHeaders=new ArrayList<>();
        leftFieldsList.stream().forEach(field -> {
            ExcelHeader excelHeader=new ExcelHeader();
            excelHeader.setColName(field.getAnnotation(MineExcelHeaderAnnotation.class).cloHeader());
            excelHeader.setIndex(field.getAnnotation(MineExcelHeaderAnnotation.class).index());
            excelHeader.setColProperty(field.getName());
            excelHeaders.add(excelHeader);
        });
        //根据列序号排序
        Collections.sort(excelHeaders);
        //创建第一行表头
        XSSFRow xssfrow=xssfSheet.createRow(0);
        excelHeaders.stream().forEach(excelHeader -> {
            xssfrow.createCell(excelHeader.getIndex()).setCellValue(excelHeader.getColName());
        });


        //创建内容
        if(!CollectionUtils.isEmpty(objectList)){
            int rowIndex=1;
            for(Object entityDemo:objectList){
                Row row=xssfSheet.createRow(rowIndex);
                for(ExcelHeader header : excelHeaders){
                    String methodName="get"+header.getColProperty().substring(0,1).toUpperCase()+header.getColProperty().substring(1);
                    try{
                        Method method=entityDemo.getClass().getDeclaredMethod(methodName);
                        Object o=method.invoke(entityDemo);
                        row.createCell(header.getIndex()).setCellValue(String.valueOf(o).equals("null")?"":String.valueOf(o));
                    }catch (Exception e){
                        System.out.println("反射出现问题");
                    }
                }
                rowIndex++;
            }
        }


        return xssfWorkbook;
    }


    /**
     * 将workbook解析为指定类对象的集合
     * @param xssfWorkbook
     * @param classType
     * @return
     */
    public static List<Object> analyseWorkbook2List(XSSFWorkbook xssfWorkbook,Class<?> classType){
        /**
         *  被解析的类对象必须要有自己定义的注解 MineExcelEntity.class
         */
        MineExcelEntity mineExcelEntity=classType.getAnnotation(MineExcelEntity.class);
        if(mineExcelEntity==null){
         throw  new RuntimeException("转换的实体必须存在@ExcelEntity!");
        }
        //把实体类解析列等信息处理  这是系统模板的列信息等
        /**
         * 把模板类先解析出来，后面与上传的文件进行格式比较
         * 解析模板类获取列信息： 【列名】【列序号】【字段类型】【是否可以为空】【最大长度】【最小长度】
         */
        List<ExcelEntityField> excelEntityFieldListEntity= getTemplateEntityFields(classType);
        /**
         *  解析上传上来的的文件的列信息
         */
        List<ExcelEntityField> excelEntityFieldListUpload=getUploadEntityFields(xssfWorkbook.getSheetAt(0));
        /**
         * 比对excelEntityFieldListEntity与excelEntityFieldListUpload   只需要 【列名】 【列序号】 【字段类型】 相同就可以了，
         * 需要重写ExcelEntityField类的equals方法。
         * 注意需要使用【上传解析的模板】来contains【系统模板】，因为上传的文件可以有其他无用的列，不用去校验，只要关键的列和模板一样就通过校验。
         */
        if(!excelEntityFieldListUpload.containsAll(excelEntityFieldListEntity)){
            throw new RuntimeException("请使用标准模板");
        }
        /**
         * 通过上面的模板比对第一行表头完成后，下面再校验表数据，没有出错的话，就返回解析的类集合。第一行表头的模板需要使用【系统模板解析的数据，因为有对应的field字段】excelEntityFieldListEntity
         */
        List<Object> list=getListObject(xssfWorkbook.getSheetAt(0),classType,excelEntityFieldListEntity);
        return list;
    }


    /**
     * 解析模板类获取列信息： 【列名】【列序号】【字段类型】【是否可以为空】【最大长度】【最小长度】
     * @param classType
     * @return
     * @throws Exception
     */
    public static List<ExcelEntityField> getTemplateEntityFields(Class<?> classType)  {
        List<ExcelEntityField> eefs = new ArrayList<ExcelEntityField>();
        /**
         *  遍历所有字段
         */
        Field[] allFields = classType.getDeclaredFields();
        for (Field field : allFields) {
            MineExcelProperty excelProperty = field.getAnnotation(MineExcelProperty.class);
            /**
             * 只对含有@ExcelProperty注解的字段进行赋值
             */
            if (excelProperty == null) {
                continue;
            }
            ExcelEntityField eef = new ExcelEntityField();

            eef.setField(field);
            eef.setColumnIndex( excelProperty.fieldIndex());
            eef.setColumnName( excelProperty.fieldName().trim());
            eef.setColumnType(excelProperty.fieldType());
            eef.setMaxLength(excelProperty.maxLength());
            eef.setMinLength(excelProperty.minLength());
            eef.setNullable(excelProperty.nullable());

            eefs.add(eef);
        }
        return eefs;
    }


    /**
     * 解析用户上传上来的文件，获得第一行表头的信息
     * @param xssfSheet
     * @return
     */
    public static List<ExcelEntityField> getUploadEntityFields(XSSFSheet xssfSheet)  {


        int rowSize=xssfSheet.getLastRowNum();
        if(rowSize==-1){
            throw new RuntimeException("文件无数据");
        }
        XSSFRow firstRow=xssfSheet.getRow(0);;
        if(rowSize==0){
            if(firstRow==null){
                throw new RuntimeException("文件没有第一行数据,请调整格式");
            }
        }
        //getLastCellNum就是实际的列数,比如表格有三列，那么这个值就是3，但是在获取第一列值的时候，还是使用0的下标获取
        int cellSize=firstRow.getLastCellNum();
        if(firstRow.getLastCellNum()==0){
            throw new RuntimeException("文件第一行无列数据,请调整格式");
        }

        List<ExcelEntityField> excelEntityFieldList=new ArrayList<>();
        //不管是行还是列，下标都是从0开始
        for(int i=0;i<cellSize;i++){
            XSSFCell xssfCell=firstRow.getCell(i);
            if(xssfCell!=null){
                ExcelEntityField excelEntityField=new ExcelEntityField();
                //获取上传上来的字段的值
                excelEntityField.setColumnName(getCellValue(xssfCell));
                //index 从0开始
                excelEntityField.setColumnIndex(xssfCell.getColumnIndex());
                //这里获取上传文件检测出来的字段类型，后面要与模板字段类型比对
                excelEntityField.setColumnType(xssfCell.getCellType());
                excelEntityFieldList.add(excelEntityField);
            }
        }
        return excelEntityFieldList;
    }

    /**
     * 返回类集合
     * @param xssfSheet                    表格
     * @param classType                    类
     * @param excelEntityFieldListEntity   表头也就是第一行每一列的元数据
     * @return
     */
    public static List<Object> getListObject(XSSFSheet xssfSheet,Class<?> classType,List<ExcelEntityField> excelEntityFieldListEntity){
        /**
         *   Map<Integer, Map<String,Object>> 记录每一列的值得类型，以及长度校验等
         * Integer               ---->第几列 从0开始
         * Map<String,Object>>   ---->该列元数据信息：如果key是nullFlag，value就是该字段是否可以为null
         *                                         如果key是max，value就是该字段最大长度限制多少
         *                                         如果key是min，value就是该字段最小长度限制多少
         *                                         如果key是field，value就是该字段Field
         */
        //要使用Object作为value,是因为下面要存储 Field类型的值
        Map<Integer, Map<String,Object>> multiMap=new HashMap<>();
        for(ExcelEntityField excelEntityField:excelEntityFieldListEntity){
            int index=excelEntityField.getColumnIndex();
            String flag=excelEntityField.isNullable()==true?"TRUE":"FALSE";
            int max=excelEntityField.getMaxLength();
            int min=excelEntityField.getMinLength();
            HashMap<String,Object> map=new HashMap();
            map.put("nullFlag",flag);
            map.put("max",String.valueOf(max));
            map.put("min",String.valueOf(min));
            map.put("field",excelEntityField.getField());
            multiMap.put(index,map);
        }
        /**
         * 获取最后一行的下标值
         */
        int rowNum=xssfSheet.getLastRowNum();
        List<Object> resultList=new ArrayList<>();
        /**
         * 因为表头占据了第一行数据，也就是下标为0的数据，这里从下标为1的地方开始取数据
         * 因为要获取最后下标的数据，所以这里使用的<=
         */
        for(int i=1;i<=rowNum;i++){
            Object object=null;
            try {
                object= classType.newInstance();
            } catch (Exception e) {
                throw new RuntimeException("newInstance出现异常");
            }
            //开始获取指定行的数据
            XSSFRow xssfRow= xssfSheet.getRow(i);
            if(xssfRow!=null){
                 /*
                 写这段代码的时候，认为只要这行的列数与标准模板列数一样就可以，
                 但是其实不对，如果客户传入的文件刚好这行，他只填写了一列数据，那么这里的列数肯定不匹配。
                 但是如果只有第一行有数据的话，其实也是可以的，因为其他列可以为空，不传入数据
                 int cellNums= xssfRow.getLastCellNum();
                 if(cellNums!=excelEntityFieldListEntity.size()){
                     throw new RuntimeException("文档内容列数不对");
                 }
                 */


                  /**
                   * 所以这里还是使用标准模板的列数，来循环取数据，再结合每列的限制校验数据合法性
                   * 列获取从0下标开始
                  */

              for(int j=0;j<excelEntityFieldListEntity.size();j++){
                  //获取列数据值，这里的值怎么样都有值，最坏情况都是 空字符串 ""
                  String cellValue=getCellValue(xssfRow.getCell(j));
                  XSSFCell x=xssfRow.getCell(j);
                  //如果该列的值是空或者空串，那么就需要该列是否可以为空
                 if(xssfRow.getCell(j)==null||"".equals(cellValue)){
                     String booleanFlag=multiMap.get(j).get("nullFlag").toString();
                     if("FALSE".equals(booleanFlag)){
                         throw new RuntimeException("文档第"+(i+1)+"行存在必填项无数据");
                     }
                 }

                 int max=Integer.parseInt(multiMap.get(j).get("max").toString());
                 int min=Integer.parseInt(multiMap.get(j).get("min").toString());
                 if(max!=0&&cellValue.length()>max){
                     throw new RuntimeException("第"+(i+1)+"行第"+(j+1)+"列数据项超过最大长度限制");
                 }
                 if(min!=0&&cellValue.length()<min){
                     throw new RuntimeException("第"+(i+1)+"行第"+(j+1)+"列数据项超过最小长度限制");

                 }
                 try
                 {
                     Field field=(Field)multiMap.get(j).get("field");
                     invoke(object,field,cellValue);
                 }catch (Exception e){
                     throw new RuntimeException("解析文档出错");
                 }
             }
          }else {
                /**
                 * 如果改行没数据，给出报错提示
                 */
              throw new RuntimeException("文档第"+(i+1)+"无数据，请调整格式");
          }
            //这里才把object加到集合里面，是因为上面有很多列，都要使用invoke设置值
            resultList.add(object);
        }
        return resultList;



    }

    public static void invoke(Object entity, Field field, String value) {
        try {
            Method method = entity.getClass()
                    .getDeclaredMethod("set" + toCapitalizeCamelCase(field.getName()), field.getType());
            method.invoke(entity, value);
        } catch (Exception e) {
            throw new RuntimeException("字段" + field.getName() + "出错!");
        }
    }

    /**
     *
     * @param name
     * @return
     */
    public static String toCapitalizeCamelCase(String name) {
        if (name == null) {
            return null;
        }
        StringBuilder sb = new StringBuilder(name.length());
        boolean upperCase = false;
        for (int i = 0; i < name.length(); i++) {
            char c = name.charAt(i);

            if (c == '_') {
                upperCase = true;
            } else if (upperCase) {
                sb.append(Character.toUpperCase(c));
                upperCase = false;
            } else {
                sb.append(c);
            }
        }
        name = sb.toString();
        return name.substring(0, 1).toUpperCase() + name.substring(1);
    }

    /**
     * 传入Cell,跟进cell类型不同获取不一样的值
     * @param cell
     * @return
     */
    public static String getCellValue(Cell cell) {
        /**
         *如果单元格为空的，则返回空字符串
         */
        if (cell == null) {
            return "";
        }
        /**
         *   根据单元格类型，以不同的方式读取单元格的值
         */
        String value = "";
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                value = cell.getRichStringCellValue().getString();
                break;
            //数值型
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    //如果是date类型则 ，获取该cell的date值
                    value = new SimpleDateFormat("yyyy-MM-dd").format(cell.getDateCellValue());
                } else {
                    // 纯数字
                    Long longVal = Math.round(cell.getNumericCellValue());
                    Double doubleVal = cell.getNumericCellValue();
                    //判断是否含有小数位.0
                    if(Double.parseDouble(longVal + ".0") == doubleVal){
                        value = longVal.toString();
                    } else {
                        value = doubleVal.toString();
                    }
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                value = cell.getBooleanCellValue() ? "TRUE" : "FALSE";
                break;
            // 公式类型
            case Cell.CELL_TYPE_FORMULA:
                //读公式计算值
                value = String.valueOf(cell.getNumericCellValue());
                if (null != value && !"".equals(value.trim())) {
                    String[] item = value.split("[.]");
                    if (1 < item.length && "0".equals(item[1])) {
                        value = item[0];
                    }
                }
                // 如果获取的数据值为非法值,则转换为获取字符串
                if ("NaN".equals(value)) {
                    value = cell.getStringCellValue();
                }
                break;
            default:
        }
        return value;
    }
}
