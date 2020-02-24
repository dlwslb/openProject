
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.beans.BeanInfo;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.io.BufferedOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * dlw
 * excel 生成下拉列表及读取excel转javaBean
 */
public final class ExcelReadUtil {

    /**
     * 生成下拉列表
     * firstRow 開始行號 根据此项目，默认为2(下标0开始)
     * lastRow  根据此项目，默认为最大65535
     * firstCol 区域中第一个单元格的列号 (下标0开始)
     * lastCol 区域中最后一个单元格的列号
     * strings 下拉内容
     * */
    public static void selectList(Workbook workbook, int firstCol, int lastCol, String[] strings,String sheetName){


        Sheet sheet = workbook.getSheetAt(0);
        //  生成下拉列表
        //  只对(x，x)单元格有效
        CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(2, 65535, firstCol, lastCol);
        //  生成下拉框内容
        //DVConstraint dvConstraint = DVConstraint.createExplicitListConstraint(strings);

        DVConstraint dvConstraint=null;
        if(sheetName!=null && !"".equals(sheetName)){
            Sheet hidden = workbook.createSheet(sheetName);
            Cell cell = null;
            for (int i = 0, length = strings.length; i < length; i++)
            {
                String name = strings[i];
                Row row = hidden.createRow(i);
                cell = row.createCell(0);
                cell.setCellValue(name);
            }
            Name namedCell = workbook.createName();
            namedCell.setNameName(sheetName);
            namedCell.setRefersToFormula(sheetName+"!$A$1:$A$" + strings.length);
            //加载数据,将名称为hidden的
            dvConstraint = DVConstraint.createFormulaListConstraint(sheetName);
        }else{
            dvConstraint = DVConstraint.createExplicitListConstraint(strings);
        }

        HSSFDataValidation dataValidation = new HSSFDataValidation(cellRangeAddressList, dvConstraint);

        //获取所有sheet页个数
        int sheetTotal = workbook.getNumberOfSheets();
        //将第二个sheet设置为隐藏
        workbook.setSheetHidden(sheetTotal-1, true);
        //  对sheet页生效
        sheet.addValidationData(dataValidation);

    }


    /**
     * 下载 xls
     * @param fileName
     * @param workbook
     * @param request
     * @param response
     * @throws IOException
     */
    public static void downLoadExcel(String fileName, Workbook workbook, HttpServletRequest request,
                                     HttpServletResponse response) throws IOException {

        OutputStream output = null;
        BufferedOutputStream bufferedOutPut = null;
        try {
            // 重置响应对象
            response.reset();
            // 当前日期，用于导出文件名称
            SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");
            String dateStr = fileName + "-" + sdf.format(new Date());

            String UserAgent = request.getHeader("USER-AGENT").toLowerCase();
            // 指定下载的文件名--设置响应头
            if (UserAgent.indexOf("firefox") >= 0) {
                response.setHeader("content-disposition", "attachment;filename=\"" + new String(dateStr.getBytes("UTF-8"), "iso-8859-1") +".xls\"");
            }else {
                response.setHeader("Content-Disposition","attachment;filename=" + URLEncoder.encode(dateStr, "UTF-8")+".xls");
            }


            response.setContentType("application/vnd.ms-excel;charset=UTF-8");
            // 编码
            response.setCharacterEncoding("UTF-8");
            output = response.getOutputStream();
            bufferedOutPut = new BufferedOutputStream(output);

            workbook.write(bufferedOutPut);
            bufferedOutPut.flush();
        } catch (Exception e) {

        } finally {
            if (bufferedOutPut != null) {
                bufferedOutPut.close();
            }
            if (output != null) {
                output.close();
            }
        }
    }

    /**
     * 读取 Excel文件内容
     *
     * @param excel_name
     * @return
     * @throws Exception
     */
    public static List<String[]> readExcel(String excel_name) throws Exception {
        // 结果集
        List<String[]> list = new ArrayList<String[]>();

        HSSFWorkbook hssfworkbook = new HSSFWorkbook(new FileInputStream(excel_name));

        // 遍历该表格中所有的工作表，i表示工作表的数量 getNumberOfSheets表示工作表的总数
        HSSFSheet hssfsheet = hssfworkbook.getSheetAt(0);

        Row row = hssfsheet.getRow(0);

        // 遍历该行所有的行,j表示行数 getPhysicalNumberOfRows行的总数
        for (int j = 0; j < hssfsheet.getPhysicalNumberOfRows(); j++) {
            HSSFRow hssfrow = hssfsheet.getRow(j);
            if(hssfrow!=null){
                int col = hssfrow.getPhysicalNumberOfCells();
                // 单行数据
                String[] arrayString = new String[col];
                for (int i = 0; i < col; i++) {
                    HSSFCell cell = hssfrow.getCell(i);
                    if (cell == null) {
                        arrayString[i] = "";
                    } else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                        // arrayString[i] = new Double(cell.getNumericCellValue()).toString();
                        if (HSSFCell.CELL_TYPE_NUMERIC == cell.getCellType()) {
                            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                Date d = cell.getDateCellValue();
//						    DateFormat formater = new SimpleDateFormat("yyyy-MM-dd");
                                DateFormat formater = new SimpleDateFormat("yyyy");
                                arrayString[i] = formater.format(d);
                            } else {
                                arrayString[i] = new BigDecimal(cell.getNumericCellValue()).longValue()+"";
                            }
                        }
                    } else {// 如果EXCEL表格中的数据类型为字符串型
                        arrayString[i] = cell.getStringCellValue().trim();
                    }
                }
            }
        }
        return list;
    }

    /**
     * 读取excel-xls格式返回javaBean对象
     * @param sheetName       要读取sheet名称
     * @param excel_file_path 文件路径
     * @param keys
     * List<String> keys = new ArrayList<String>();
     * keys.add("字段名菜例如 userName");
     * @param classBean 定义的要转换的javaBean类
     * @param <T>
     * @return
     * @throws Exception
     */
    public static <T> List<T> readXlsExcelBackBean(String sheetName,String excel_file_path,List<String> keys,Class<T> classBean) throws Exception {
        // 结果集
        List<T> list = new ArrayList<T>();
        HSSFWorkbook hssfworkbook = new HSSFWorkbook(new FileInputStream(excel_file_path));
        // 遍历该表格中所有的工作表，i表示工作表的数量 getNumberOfSheets表示工作表的总数
        //HSSFSheet hssfsheet = hssfworkbook.getSheetAt(0);
        HSSFSheet hssfsheet=null;
        for (int i = 0; i < hssfworkbook.getNumberOfSheets(); i++) {//获取每个Sheet表
            hssfsheet=hssfworkbook.getSheetAt(i);
            if(sheetName.equals(hssfsheet.getSheetName())){
                break;
            }
        }
        if(hssfsheet!=null){
            // 遍历该行所有的行,j表示行数 getPhysicalNumberOfRows行的总数
            for (int j = 0; j < hssfsheet.getPhysicalNumberOfRows(); j++) {
                HSSFRow hssfrow = hssfsheet.getRow(j);
                if(hssfrow!=null){
                    int col = hssfrow.getPhysicalNumberOfCells();
                    // 单行数据
                    Map<String, Object> map = new HashMap<String, Object>();
                    for (int i = 0; i < col; i++) {
                        HSSFCell cell = hssfrow.getCell(i);
                        if (cell == null) {
                            map.put(keys.get(i), "");
                        }else{
                            map.put(keys.get(i), getValueAllString(cell));
                        }

                    }
                    //IndicatorItemVoPoi t = mapToObject(IndicatorItemVoPoi.class,map);
                    T t2 = convertMap(classBean,map);
                    list.add(t2);
                }
            }
        }

        return list;
    }

    /**
     * 读取excel-xls格式返回javaBean对象
     * @param keys
     * List<String> keys = new ArrayList<String>();
     * keys.add("字段名菜例如 userName");
     * @param listXlsx    List<String[]> listXlsx = XLSXCovertCSVReader.readerExcel(filepathStr,sheetName, sheetColosLength);
     * @param classBean   定义的要转换的javaBean类
     * @param <T>
     * @return
     * @throws Exception
     */
    public static <T> List<T> readXlsxExcelBackBean(List<String> keys,List<String[]> listXlsx,Class<T> classBean) throws Exception {
        // 结果集
        List<T> list = new ArrayList<T>();
        if(keys!=null && listXlsx!=null){
            // 遍历该行所有的行,j表示行数 getPhysicalNumberOfRows行的总数
            for (int j = 0; j < listXlsx.size(); j++) {
                String[] hssfrow = listXlsx.get(j);
                if(hssfrow!=null){
                    int col = hssfrow.length;
                    // 单行数据
                    Map<String, Object> map = new HashMap<String, Object>();
                    for (int i = 0; i < col; i++) {
                        String cell = hssfrow[i];
                        if (cell == null) {
                            map.put(keys.get(i), "");
                        }else{
                            map.put(keys.get(i), cell);
                        }
                    }
                    T t2 = convertMap(classBean,map);
                    list.add(t2);
                }
            }
        }
        return list;
    }
    /**
     * map 转换 javabean对象
     * @param type
     * @param map
     * @return
     * @throws Exception
     */
    public static <T> T  convertMap(Class<T> type, Map map) throws Exception {
        BeanInfo beanInfo = Introspector.getBeanInfo(type);
        T obj = type.newInstance();
        PropertyDescriptor[] propertyDescriptors =  beanInfo.getPropertyDescriptors();
        for (PropertyDescriptor descriptor : propertyDescriptors) {
            String propertyName = descriptor.getName();
            if (map.containsKey(propertyName)) {
                Object value = map.get(propertyName);
                descriptor.getWriteMethod().invoke(obj, value);
            }
        }
        return obj;
    }

    private static <T> T mapToObject(Class<T> c,Map<String, Object> map) throws Exception {
        BeanInfo beanInfo = Introspector.getBeanInfo(c);
        T t = c.newInstance();
        PropertyDescriptor[] propertyDescriptors = beanInfo.getPropertyDescriptors();
        for(int i = 0; i < propertyDescriptors.length; i++){
            PropertyDescriptor descriptor = propertyDescriptors[i];
            String propertyName = descriptor.getName();
            if(map.containsKey(propertyName)){
                Object value = map.get(propertyName);
                Object[] args = new Object[1];
                args[0] = value;
                //这里捕获异常为了让不正常的值可以暂时跳过不影响正常字段的赋值
                try {
                    descriptor.getWriteMethod().invoke(t, args);
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                } catch (IllegalArgumentException e) {
                    e.printStackTrace();
                }
            }
        }
        return t;
    }
    /**
     * 按类型转换
     * @param cell
     * @return
     */
    private static Object getValue(Cell cell) {
        if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
            return cell.getBooleanCellValue();
        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            return cell.getNumericCellValue();
        } else {
            return String.valueOf(cell.getStringCellValue());
        }
    }
    /**
     * 全转字符串
     * @param cell
     * @return
     */
    private static Object getValueAllString(Cell cell) {
        if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
            return cell.getBooleanCellValue();
        }else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
            String rsstr ="";
            // arrayString[i] = new Double(cell.getNumericCellValue()).toString();
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                Date d = cell.getDateCellValue();
                DateFormat formater = new SimpleDateFormat("yyyy-MM-dd");
                rsstr = formater.format(d);
            } else {
                rsstr = new BigDecimal(cell.getNumericCellValue()).longValue()+"";
            }
            return rsstr;
        }else {
            return String.valueOf(cell.getStringCellValue());
        }
    }

    public static void main(String args[]) throws Exception{
        ExcelSelectListUtil.readExcel("C:\\Users\\dlw\\Downloads\\20200224.xls");

    }
}