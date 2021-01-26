package com.eps.service.impl;

import com.eps.service.DealDataService;
import com.eps.ui.TestInfoImportUI;
import com.eps.util.APIExceptionUtil;
import com.eps.util.ExceptionUtil;
import com.eps.util.IntegrityUtil;
import com.mks.api.response.APIException;
import com.sun.corba.se.impl.orb.ParserTable;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Map.Entry;

/**
 * Description: 处理数据
 *
 * @author @ModifyDate
 * Yi Gang  2020/12/17
 */
public class DealDataServiceImpl implements DealDataService {

    public static final Logger logger = Logger.getLogger(DealDataServiceImpl.class);

    private static List<String> caseFields = new ArrayList<>();//Test Case数据
    private static List<String> stepFields = new ArrayList<>();//Test Step数据
    private static List<String> resultFields = new ArrayList<>();//Test Result数据
    private static List<String> defectFields = new ArrayList<>();//Defect 数据
    private static List<String> sessionFields = new ArrayList<>();//Session 数据
    private static List<String> allHeaders = new ArrayList<>();//所有标题记录
    private static Map<String, String> resultFieldsMap = new HashMap<>();
    public static String[][] tableFields = null;
    private static Map<String, Map<String, String>> headerConfig = new HashMap<>();
    private static final String FIELD_CONFIG_FILE = "FieldMapping.xml";//记录Field Mapping配置
    private static final String CATEGORY_CONFIG_FILE = "Category.xml";//Category配置
    private static final String TEST_STEP = "Test Step";//
    private static final String TEST_RESULT = "Test Result";
    private static final String TEST_SESSION = "Test Session";
    private static final String DEFECT = "Defect";
    private static final String SESSION_STATE = "In Testing";//Test In Testing状态
    private static final String SESSION_INIT_STATE = "Submit";//Test In Testing状态
    private static final String ID_FIELD = "ID";
    private static final String VERDICT = "Verdict";
    private static final String SESSION_ID = "Session ID";//Test Session ID
    private static final String SESSION_HEADER_FLAG = "测试基本信息";

    private static String DEFAULT_CATEGORY;

    private static final String TEST_INFO_FLAG = "测试用例及结果";

    private static final String TEST_RESULT_FLAG = "测试结果";

    private static final String STEP_RESULT_FLAG = "Cycle Verdict";

    private static final List<String> CURRENT_CATEGORIES = new ArrayList<String>();//记录导入对象的正确Category

    private static final Map<String, List<String>> PICK_FIELD_RECORD = new HashMap<String, List<String>>();

    private static final Map<String, String> FIELD_TYPE_RECORD = new HashMap<String, String>();

    public static final List<String> RICH_FIELDS = new ArrayList<String>();

    private static String IMPORT_DOC_TYPE = "Test Suite";//导入测试文档类型

    private static String CONTENT_TYPE;//文档条目类型

    private static final String SPERATOR = "_";
    ;//定义分隔符

    private static final String INIT_DOC_STATE = "Open";//Test Suite的创建初始状态

    private static final String INIT_CONTENT_STATE = "Active";//Test Case的创建初始状态

    private static final String INIT_STATE = "Active";//Test Case 的初始状态

    private Map<String, CellRangeAddress> cellRangeMap = new HashMap<String, CellRangeAddress>();//记录所有的合并单元格

    private static final List<String> USER_FULLNAME_RECORD = new ArrayList<String>();//记录 loginID-fullname
    private static boolean IS_USER = false;
    private static boolean RELATIONSHIP_MISTAKEN = false;
    private static boolean DOC_STATE = false;

    /**
     * 利用Jsoup解析配置文件，得到相应的参数，为Type选项和创建Document提供信息 (1)
     * Document:Type,Project,State,Shared Category (2) Content:Type
     *
     * @return
     * @throws Exception
     */
    @Override
    public List<String> parsFieldMapping() throws Exception {

        DealDataServiceImpl.logger.info("start to parse xml : " + FIELD_CONFIG_FILE);
        Document doc = DocumentBuilderFactory.newInstance().newDocumentBuilder()
                .parse(DealDataServiceImpl.class.getClassLoader().getResourceAsStream(FIELD_CONFIG_FILE));
        Element root = doc.getDocumentElement();
        List<String> typeList = new ArrayList<String>();
        if (root == null)
            return typeList;
        // 得到xml配置
        NodeList importTypes = root.getElementsByTagName("importType");  // 拿到mapping里面所有的 ImportType
        if (importTypes == null || importTypes.getLength() == 0) {
            throw new Exception("Can't not parse xml because of don't has \"importType\"");
        } else {
            // 循环 刚才拿到的所有ImportType
            Map<String, String> map = null;
            for (int j = 0; j < importTypes.getLength(); j++) {
                Element importType = (Element) importTypes.item(j);
                // 获取XML 文件的name 和  Type
                String documentType = importType.getAttribute("type");
                IMPORT_DOC_TYPE = documentType;
                DEFAULT_CATEGORY = importType.getAttribute("defaultCategory");
                NodeList excelFields = importType.getElementsByTagName("excelField");
                try {
                    if (excelFields == null || excelFields.getLength() == 0) {
                        throw new Exception("Can't not parse xml because of don't has \"excelField\"");
                    } else {
                        tableFields = new String[excelFields.getLength()][2];
                        for (int i = 0; i < excelFields.getLength(); i++) {
                            Element fields = (Element) excelFields.item(i);
                            String name = fields.getAttribute("name");
                            map = new HashMap<>();
                            String type = fields.getAttribute("type");
                            String field = fields.getAttribute("field");
                            allHeaders.add(name);
                            map.put("type", type);
                            if (TEST_STEP.equals(type) && !stepFields.contains(name)) {
                                stepFields.add(name);
                            } else if (TEST_RESULT.equals(type) && !resultFields.contains(name)) {
                                resultFields.add(name);
                                resultFieldsMap.put(name, field);
                            } else if (TEST_SESSION.equals(type) && !sessionFields.contains(name)) {
                                sessionFields.add(name);
                            } else if (DEFECT.equals(type) && !defectFields.contains(name)) {
                                defectFields.add(name);
                            } else if (!TEST_STEP.equals(type) && !TEST_RESULT.equals(type) && !DEFECT.equals(type)
                                    && !caseFields.contains(name)) {
                                caseFields.add(name);
                                CONTENT_TYPE = type;
                            }

                            map.put("field", field);
                            // 获取 excelField 的  onlyCreate 属性 ， 若没有填写则默认为 false
                            String onlyCreate = fields.getAttribute("onlyCreate");
                            if (onlyCreate == null || onlyCreate.equals("")) {
                                map.put("onlyCreate", "false");
                            } else {
                                map.put("onlyCreate", onlyCreate);
                            }
                            String overRide = fields.getAttribute("overRide");
                            if (overRide == null || overRide.equals("")) {
                                map.put("overRide", "true");
                            } else {
                                map.put("overRide", overRide);
                            }
                            tableFields[i][0] = name;
                            tableFields[i][1] = field;
                            headerConfig.put(name, map);
                        }
                    }
                } catch (ParserConfigurationException e) {
                    logger.error("parse config file exception", e);
                } catch (SAXException e) {
                    logger.error("get config file exception", e);
                } catch (IOException e) {
                    logger.error("io exception", e);
                } finally {
                    logger.info("get info : \nheaderConfig : " + headerConfig);
                }
            }
        }
        return typeList;
    }

    /**
     * Description 查询当前要导入类型的 正确Category
     *
     * @param documentType
     * @throws Exception
     */
    @Override
    public void parseCurrentCategories(String documentType) throws Exception {
        Document doc = DocumentBuilderFactory.newInstance().newDocumentBuilder()
                .parse(DealDataServiceImpl.class.getClassLoader().getResourceAsStream(CATEGORY_CONFIG_FILE));
        Element root = doc.getDocumentElement();
        // 得到xml配置
        NodeList importTypes = root.getElementsByTagName("documentType");
        for (int j = 0; j < importTypes.getLength(); j++) {
            Element importType = (Element) importTypes.item(j);
            String typeName = importType.getAttribute("name");
            if (typeName.equals(documentType)) {
                NodeList categoryNodes = importType.getElementsByTagName("category");
                for (int i = 0; i < categoryNodes.getLength(); i++) {
                    Element categoryNode = (Element) categoryNodes.item(i);
                    CURRENT_CATEGORIES.add(categoryNode.getAttribute("name"));
                }
            }
        }
    }

    /**
     * 获得Excel中的数据
     *
     * @param
     * @return
     * @throws
     * @throws IOException
     */
    @Override
    public List<List<Map<String, Object>>> parseExcel(File file) throws Exception {
        Workbook wb = null;
        String fileName = file.getName();
        if (fileName.endsWith(".xlsx")) {
            wb = new XSSFWorkbook(file);
        } else if (fileName.endsWith(".xls")) {
            wb = new HSSFWorkbook(new FileInputStream(file));
        }
        List<List<Map<String, Object>>> allDataList = new ArrayList<>();
        List<Map<String, Object>> sessionList = new ArrayList<>();//存储Session信息
        List<Map<String, Object>> caseList = new ArrayList<>();//存储 Test Case/Test Result /Defect /Test Step信息
        allDataList.add(sessionList);
        allDataList.add(caseList);
        Sheet sheet = wb.getSheetAt(0);
        Integer IDIndex = 0;
        if (allHeaders.indexOf(ID_FIELD) > -1)
            IDIndex = allHeaders.indexOf(ID_FIELD);
        List<CellRangeAddress> mergeList = sheet.getMergedRegions();
        cellRangeMap = new HashMap<String, CellRangeAddress>();
        int rowNum = this.getRealRowNum(sheet, mergeList);
        int colNum = this.getRealColNum(sheet);
        int rowIndex = 1;//定义开始循环行

//		int merge = getMergeRow(mergeList);
//		if(merge > 0) {
//			row = row + merge;
//		}
        int endRow = rowIndex + rowNum;
        //找到Session Field的标题行
        Row sessionFieldRow = sheet.getRow(0);//假设第一行为 sessionField行
        Cell sessionFieldcell = sessionFieldRow.getCell(0);
        String valueVal = getCellVal(sessionFieldcell);
        if (SESSION_HEADER_FLAG.equals(valueVal)) {//如果第一行获取到的内容为 "测试基本信息"，向下一行为 Sesssion标题
            sessionFieldRow = sheet.getRow(1);
            rowIndex++;
        }
        Map<String, Object> map = null;//定义保存数据的Map
        //1. 先获取Test Session信息
        for (; rowIndex < endRow; rowIndex++) {//从第一行开始，找到第一列内容为：Test Case的 “测试用例ID” 或 “测试用例”结束
            Row sessionRow = sheet.getRow(rowIndex);
            boolean allEmptyValue = true;//如果整行数据都为空，不添加
            boolean endSessionInfo = false;//是否结束Test Session数据
            map = new HashMap<>();
            for (int col = 0; col < sessionFields.size(); col++) {//循环读取内容、判断Test Session有几个标题，多余内容不读取
                Cell fieldCell = sessionFieldRow.getCell(col);
                String fieldName = getCellVal(fieldCell);
                Cell cell = sessionRow.getCell(col);
                String fieldValue = getCellVal(cell);
                if (fieldValue != null && !"".equals(fieldValue)) {
                    if (TEST_INFO_FLAG.equals(fieldValue)) {//循环到了Test Info位置,结束循环
                        endSessionInfo = true;
                        break;
                    } else {
                        allEmptyValue = false;
                    }
                    map.put(fieldName, fieldValue);//如果此列信息为空 不保存
                }
            }
            if (!allEmptyValue) {
                sessionList.add(map);
            }
            if (endSessionInfo) {//结束 test session循环
                rowIndex++;
                break;
            }
        }
        Row mergeRow = sheet.getRow(rowIndex++);//合并表头
        Row testHeaderRow = sheet.getRow(rowIndex++);//字段表头
        colNum = getRealTestInfoCol(sheet, rowIndex);
        //2. 在循环获取其他信息
        for (; rowIndex < endRow; rowIndex++) {
            try {
                map = new HashMap<>();
                Map<String, String> stepMap = null;
                Map<String, String> resultMap = null;
                Map<String, String> defectMap = null;
                List<Map<String, String>> stepList = new ArrayList<Map<String, String>>();//保存Test Step信息
                List<Map<String, String>> resultList = new ArrayList<Map<String, String>>();//保存Test Result信息
                List<Map<String, String>> defectList = new ArrayList<Map<String, String>>();//保存defect信息
                //Test Case可关联多个Test Step信息，通过多行关联
                //Test Case可关系多个Test Result信息，通过多列关联
                //Test Case可关系多个Defect信息，通过多列关联
                int caseMerge = 0;
                CellRangeAddress IDCellRange = cellRangeMap.get(rowIndex + SPERATOR + IDIndex);
                if (IDCellRange != null) {
                    int endMergeRow = IDCellRange.getLastRow();
                    caseMerge = endMergeRow - rowIndex;
                }
                int temp = rowIndex;
                for (; temp <= rowIndex + caseMerge; temp++) {
                    Row dataRow = sheet.getRow(temp);
                    stepMap = new HashMap<String, String>();
                    int stepResultIndex = 1;
                    for (int col = 0; col < colNum; col++) {
                        Cell fieldCell = mergeRow.getCell(col);
                        Cell secondCell = testHeaderRow.getCell(col);
                        String field = getCellVal(fieldCell);
                        String secondFieldVal = getCellVal(secondCell);
                        Cell valueCell = dataRow.getCell(col);
                        String fieldValue = getCellVal(valueCell);
                        if (field != null && field.indexOf(TEST_RESULT_FLAG) > 0) {//第一个标题是 “**测试结果”，新建Map
                            resultMap = new HashMap<String, String>();
                            resultList.add(resultMap);
                            defectMap = new HashMap<String, String>();
                            defectList.add(defectMap);
                        }
                        if (secondFieldVal != null && stepFields.contains(secondFieldVal)) {//test Step数据
                            if (fieldValue != null && !"".equals(fieldValue)) {//Test Step可能有多个测试结果录入
                                Map<String, String> headerMap = headerConfig.get(secondFieldVal);
                                if (headerMap != null) {
                                    String fieldName = headerMap.get("field");
                                    if (STEP_RESULT_FLAG.equals(fieldName)) {
                                        String realFieldName = "Cycle" + (stepResultIndex > 1 ? stepResultIndex : "") + " Verdict";//如果多个测试结果，添加数字后缀
                                        stepMap.put(realFieldName, fieldValue);
                                        stepResultIndex++;
                                    } else {
                                        stepMap.put(secondFieldVal, fieldValue);
                                    }
                                }
                            }
                        } else {
                            if (temp == rowIndex) {//只要Test Step存在多行，其他数据只有第一行有效
                                if (secondFieldVal != null && caseFields.contains(secondFieldVal)
                                        && fieldValue != null && !"".equals(fieldValue)) {//这是Test Case数据
                                    map.put(secondFieldVal, fieldValue);
                                } else if (secondFieldVal != null && resultFields.contains(secondFieldVal)) {//循环处理Test Result
                                    if (fieldValue != null && !"".equals(fieldValue)) {
                                        resultMap.put(secondFieldVal, fieldValue);
                                    }
                                } else if (secondFieldVal != null && defectFields.contains(secondFieldVal)) {//Defect信息处理
                                    if (fieldValue != null && !"".equals(fieldValue)) {
                                        defectMap.put(secondFieldVal, fieldValue);
                                    }
                                }
                            }
                        }
                    }
                    if (!stepMap.isEmpty()) {
                        stepList.add(stepMap);
                    }
                }
                rowIndex = rowIndex + caseMerge;
                if (!stepList.isEmpty()) {
                    map.put(TEST_STEP, stepList);
                }
                if (!resultList.isEmpty()) {
                    map.put(TEST_RESULT, resultList);
                }
                if (!defectList.isEmpty()) {
                    map.put(DEFECT, defectList);
                }
                if (!map.isEmpty()) {
                    caseList.add(map);
                }
            } catch (Exception e) {
                e.printStackTrace();
                System.out.println(rowIndex);
            }
        }
        return allDataList;
    }

    /**
     * @param cell
     * @return
     */
    @SuppressWarnings("deprecation")
    public String getCellVal(Cell cell) {
        String value = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                    value = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_BLANK:
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    value = String.valueOf(cell.getCellFormula());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        Date dateCellValue = cell.getDateCellValue();
                        SimpleDateFormat format = new SimpleDateFormat("yyyy/MM/dd");
                        value = format.format(dateCellValue);
                    } else {
                        value = String.valueOf(Math.round(cell.getNumericCellValue()));//当前项目 没有Number类型，只有String。取整
                    }
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    value = String.valueOf(cell.getBooleanCellValue());
                    break;
            }
        }
        return value;
    }

    /**
     * 处理Excel中的数据，将Test Step信息和Test Case信息拆分开
     *
     * @param
     * @return
     */
    @Override
    @SuppressWarnings("unchecked")
    public List<List<Map<String, Object>>> dealExcelData(List<List<Map<String, Object>>> datas) {
        List<List<Map<String, Object>>> dealDataList = new ArrayList<>();
        List<Map<String, Object>> newSessionData = new ArrayList<>();
        List<Map<String, Object>> newData = new ArrayList<>();
        Map<String, Object> newMap = null;
        List<Map<String, Object>> sessionData = datas.get(0);//session 信息
        List<Map<String, Object>> data = datas.get(1);//Case 信息
        for (int i = 0; i < sessionData.size(); i++) {//循环处理Test Session信息
            Map<String, Object> rowMap = sessionData.get(i);
            newMap = new HashMap<String, Object>();
            for (String header : sessionFields) {
                Map<String, String> headerMap = headerConfig.get(header);//保存的 标题 Map映射信息
                String fieldName = headerMap.get("field");//实际字段名
                String value = (String) rowMap.get(header);
                if (value != null && !"".equals(value)) {
                    newMap.put(fieldName, value);
                }
            }
            newSessionData.add(newMap);
        }
        dealDataList.add(newSessionData);//添加保存
        for (int i = 0; i < data.size(); i++) {
            Map<String, Object> rowMap = data.get(i);
            String caseID = (String) rowMap.get(ID_FIELD);
            newMap = new HashMap<String, Object>();
            if (caseID != null && !"".equals(caseID)) {
                newMap.put("ID", caseID);
            }
            for (String header : caseFields) {//test Case本身的信息
                Map<String, String> fieldConfig = headerConfig.get(header);
                if (fieldConfig != null) {
                    String field = fieldConfig.get("field");
                    Object value = rowMap.get(header);
                    if (value != null && !"".equals(value.toString())) {
                        newMap.put(field, value);
                    }
                }
            }
            if (rowMap.containsKey(TEST_STEP)) {//Test Case包含有 Test Step信息
                Object steps = rowMap.get(TEST_STEP);
                if (steps instanceof List) {
                    List<Map<String, String>> currentSteps = (List<Map<String, String>>) steps;
                    if (!currentSteps.isEmpty()) {//Test Case包含有 Test Step信息
                        List<Map<String, String>> stepList = new ArrayList<Map<String, String>>();
                        Map<String, String> stepMap = null;
                        boolean hasVal = false;
                        for (Map<String, String> map : currentSteps) {//循环处理Test Step信息
                            hasVal = false;
                            stepMap = new HashMap<String, String>();
                            for (String header : map.keySet()) {
                                Map<String, String> fieldConfig = headerConfig.get(header);
                                String value = map.get(header);
                                if (fieldConfig != null) {
                                    String field = fieldConfig.get("field");
                                    if (value != null && !"".equals(value)) {
                                        stepMap.put(field, value);
                                        hasVal = true;
                                    }
                                } else {
                                    stepMap.put(header, value);
                                }
                            }
                            if (hasVal) {
                                stepList.add(stepMap);
                            }

                        }
                        newMap.put(TEST_STEP, stepList);
                    }
                }
            }
            if (rowMap.containsKey(TEST_RESULT)) {//Test Case包含有 Test Result信息
                Object results = rowMap.get(TEST_RESULT);
                if (results instanceof List) {
                    List<Map<String, String>> currentResults = (List<Map<String, String>>) results;
                    if (!currentResults.isEmpty()) {//Test Case包含有 Test Result信息
                        List<Map<String, String>> resultList = new ArrayList<Map<String, String>>();
                        Map<String, String> resultMap = null;
                        boolean hasVal = false;
                        for (Map<String, String> map : currentResults) {//循环处理Test Result信息
                            hasVal = false;
                            resultMap = new HashMap<String, String>();
                            for (String header : resultFields) {
                                if (map.containsKey(header)) {
                                    Map<String, String> fieldConfig = headerConfig.get(header);
                                    if (fieldConfig != null) {
                                        String field = fieldConfig.get("field");
                                        String value = map.get(header);
                                        if (value != null && !"".equals(value)) {
                                            resultMap.put(field, value);
                                            hasVal = true;
                                        }
                                    }
                                }
                            }
                            if (hasVal) {
                                resultList.add(resultMap);
                            }
                        }
                        newMap.put(TEST_RESULT, resultList);
                    }
                }
            }
            if (rowMap.containsKey(DEFECT)) {//处理Defect信息
                Object defects = rowMap.get(DEFECT);
                if (defects instanceof List) {
                    List<Map<String, String>> recordDefects = (List<Map<String, String>>) defects;
                    List<Map<String, String>> defectList = new ArrayList<Map<String, String>>();
                    Map<String, String> defectMap = null;
                    for (Map<String, String> map : recordDefects) {
                        defectMap = new HashMap<>();
                        for (String header : defectFields) {
                            Map<String, String> fieldConfig = headerConfig.get(header);
                            if (fieldConfig != null) {
                                String field = fieldConfig.get("field");
                                String value = map.get(header);
                                if (value != null && !"".equals(value)) {
                                    defectMap.put(field, value);
                                }
                            }
                        }
                        if (!defectMap.isEmpty()) {
                            defectList.add(defectMap);
                        }
                    }
                    if (!defectList.isEmpty()) {
                        newMap.put(DEFECT, defectList);
                    }
                }
            }
            newData.add(newMap);
        }
        dealDataList.add(newData);
        return dealDataList;
    }

    /**
     * 获得真正的row数：<br/>
     * <li>根据Test Case ID，整行数据确定真正的行数</li>
     *
     * @param sheet
     * @param
     * @return
     */
    public int getRealRowNum(Sheet sheet, List<CellRangeAddress> mergeList) throws Exception {
        int realRow = 0;
        int i = 1;
        int merge = getMergeRow(mergeList);
        i = i + merge;//如果有合并单元格，加上
        int titleCount = 1 + merge;
        for (; i <= sheet.getLastRowNum(); i++) {
            Row currentRow = sheet.getRow(i);
            if (currentRow == null || "".equals(currentRow.toString())) {
                break;
            }
            realRow = i + 1;
        }
        return (realRow - titleCount);
    }

    /**
     * Description 判断列头是否有合并单元格
     *
     * @param
     */
    public Integer getMergeRow(List<CellRangeAddress> mergeList) {
        int merge = 0;
        if (mergeList != null && !mergeList.isEmpty()) {
            for (CellRangeAddress range : mergeList) {
                int firstRow = range.getFirstRow();
                int lastRow = range.getLastRow();
                int firstCell = range.getFirstColumn();
                cellRangeMap.put(firstRow + SPERATOR + firstCell, range);
                if (firstRow == 0 && lastRow > 0) {
                    if (merge < (lastRow - firstRow)) {
                        merge = lastRow - firstRow;
                    }
                }
            }
        }
        return merge;
    }

    /**
     * 获得真正的column数
     *
     * @param sheet
     * @return
     */
    public int getRealColNum(Sheet sheet) {
        int num = 0;
        Row headRow = sheet.getRow(0);
        Row secondRow = sheet.getRow(1);
        num = headRow.getLastCellNum();
        if (num < secondRow.getLastCellNum()) {
            num = secondRow.getLastCellNum();
        }
        return num;
    }

    /**
     * 返回真正的Column数
     *
     * @param sheet
     * @param rowIndex
     * @return
     */
    public int getRealTestInfoCol(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        return row.getLastCellNum();
    }


    /**
     * Description 校验下拉框输入
     *
     * @return
     * @throws APIException
     */
    public String checkPickVal(String header, String field, String value, IntegrityUtil cmd) throws APIException {
        if (value == null || "".equals(value)) {
            return null;
        }
        List<String> valList = PICK_FIELD_RECORD.get(field);
        if (valList == null) {
            valList = cmd.getAllPickValues(field);
        }
        if (valList == null) {
            return "Column [" + (header != null ? header : field) + "] has no valid option value!";
        } else if (!valList.contains(value)) {
            return "Value [" + value + "] is invalid for Column [" + (header != null ? header : field) + "], valid values is " + Arrays.toString(valList.toArray()) + "!";
        }
        return null;
    }

    /**
     * Description 校验关联字段输入
     *
     * @return
     */
    public String checkRelationshipVal() {

        return "";
    }

    /**
     * Description 校验用户输入
     *
     * @return
     */
    public String checkUserVal(String value, String field) {
        int leftIndex = -1;
        int rightIndex = -1;
        boolean endFormat = false;
        if (value.indexOf("(") > -1) {
            leftIndex = value.indexOf("(");
        } else if (value.indexOf("（") > -1) {
            leftIndex = value.indexOf("（");
        }
        if (value.indexOf(")") > -1) {
            rightIndex = value.indexOf(")");
            endFormat = value.endsWith(")");
        } else if (value.indexOf("）") > -1) {
            rightIndex = value.indexOf("）");
            endFormat = value.endsWith("）");
        }
        String formatValue = null;
        if (leftIndex > 0 && rightIndex > 0 && endFormat) {
            formatValue = value.substring(leftIndex + 1, rightIndex);
        } else {
            formatValue = value;
        }
        if (USER_FULLNAME_RECORD.contains(formatValue.toLowerCase())) {
            IS_USER = true; // 若用户存在修改标识 ， 往下执行好判断
            return "";
        }
        return "Column [" + field + "] input value [" + value + "] is not exist";
    }

    /**
     * Description 校验relationship 输入的ID 是否带[]，是的话去掉
     *
     * @return
     */
    public String checkRelationshipVal(String value) {
        if (value.startsWith("[") && value.endsWith("]")) {
            RELATIONSHIP_MISTAKEN = true;
        }
        return "";
    }

    /**
     * Description 校验组输入
     *
     * @return
     */
    public String checkGroupVal() {

        return "";
    }

    /**
     * Description 校验组输入
     *
     * @return
     */
    public String checkBooleanVal() {

        return "";
    }

    /**
     * Description 校验输入值是否合法
     *
     * @return
     * @throws APIException
     */
    public String checkFieldValue(String header, String field, String value, IntegrityUtil cmd) throws APIException {
        String fieldType = FIELD_TYPE_RECORD.get(field);

        if ("pick".equalsIgnoreCase(fieldType)) {
            return checkPickVal(header, field, value, cmd);
        }
        if ("Category".equalsIgnoreCase(field)) {
            return checkCategory(value);
        }
        if ("Date".equalsIgnoreCase(fieldType)) {
            return checkDate(value);
        }
        if ("User".equalsIgnoreCase(fieldType)) {
            return checkUserVal(value, field);
        }
        if ("relationship".equalsIgnoreCase(fieldType)) {
            return checkRelationshipVal(value); // 检查关联的ID是不是带 []
        }
        return null;
    }

    /**
     * Description 校验Category
     *
     * @return
     */
    public String checkCategory(String value) {
        if (!CURRENT_CATEGORIES.contains(value)) {
            return "[" + value + "] is invalid for Category, valid values is " + Arrays.toString(CURRENT_CATEGORIES.toArray()) + "!";
        }
        return null;
    }

    /**
     * Description 校验时间格式
     *
     * @return
     */
    public String checkDate(String value) {
        value = value.trim();
        SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        SimpleDateFormat sdf3 = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
        Date date = null;
        try {
            date = sdf2.parse(value);
            if (date == null) {
                date = sdf3.parse(value);
            }
        } catch (ParseException e) {
            e.printStackTrace();
        }
        if (date == null) {
            return "[" + value + "] input error, The date and date you entered is incorrectly formatted."
                    + "The Correct Format : [yyyy-MM-dd HH:mm:ss] [yyyy/MM/dd HH:mm:ss] ";
        }
        return null;

    }

    /**
     * Description 处理数据，并校验
     *
     * @param cmd
     * @return
     * @throws Exception
     */
    @SuppressWarnings({"unchecked", "static-access", "unused"})
    @Override
    public List<List<Map<String, Object>>> checkExcelData(List<List<Map<String, Object>>> datas, Map<String, String> errorRecord, IntegrityUtil cmd) throws Exception {
        List<List<Map<String, Object>>> resultDataList = new ArrayList<>();

        StringBuffer allMessage = new StringBuffer();
        List<Map<String, Object>> sessionDatas = datas.get(0);//CASE 信息
        resultDataList.add(sessionDatas);
        List<Map<String, Object>> testInfoDatas = datas.get(1);//CASE 信息
        List<Map<String, Object>> testInfoList = new ArrayList<>(testInfoDatas.size());
        TestInfoImportUI.logger.info("Begin Deal Excel Data ,Total Data is :" + testInfoDatas.size());
        if (FIELD_TYPE_RECORD == null || FIELD_TYPE_RECORD.isEmpty()) {
            /** 查询Field ，为Field校验做准备*/
            List<String> importFields = new ArrayList<String>();
            for (String header : allHeaders) {
                if (resultFields.contains(header)) {//跳过Test Result字段
                    continue;
                }
                Map<String, String> fieldConfig = headerConfig.get(header);
                if (fieldConfig != null) {
                    String field = fieldConfig.get("field");
                    if (!"-".equals(field)) {
                        importFields.add(field);
                    }
                }
            }
            FIELD_TYPE_RECORD.putAll(cmd.getAllFieldType(importFields, PICK_FIELD_RECORD));
        }
        if (CURRENT_CATEGORIES.isEmpty()) {
            parseCurrentCategories(IMPORT_DOC_TYPE);
        }
        this.USER_FULLNAME_RECORD.addAll(cmd.getAllUserIdAndName()); // 查询出所有的user的name 和  Id 然后存放在 USER_FULLNAME_RECORD
        Map<String, Object> newMap = null;
        List<String> sessionIds = new ArrayList<String>();
        Map<String, List<String>> sessionInfoRecord = new HashMap<>();// 只获取系统当前Session关联的Test Case的信息。
        Map<String, String> sessionStateRecord = new HashMap<>();// 只获取系统当前Session的状态信息。
        for (Map<String, Object> sessionMap : sessionDatas) {//找出所有Session ID
            for (String header : sessionFields) {
                Map<String, String> fieldConfig = headerConfig.get(header);
                if (fieldConfig != null) {
                    String field = fieldConfig.get("field");
                    String value = (String) sessionMap.get(field);
                    if (!"-".equals(field) && value != null && !"".equals(value)) {
                        String message = checkFieldValue(header, field, value, cmd);//校验Test Session字段值
                        if (message != null && !"".equals(message)) {
                            allMessage.append(message);
                        }
                    }
                }
            }
            Object sessionObj = sessionMap.get(ID_FIELD);
            if (sessionObj != null && !"".equals(sessionObj.toString())) {
                sessionIds.add(sessionObj.toString());
            }
        }
        if (!sessionIds.isEmpty()) {//根据 Test Session查询信息
            List<Map<String, String>> sessionInfoList = cmd
                    .getItemByIds(sessionIds, Arrays.asList("ID", "Tests", "State"));
            if (sessionInfoList != null && !sessionInfoList.isEmpty()) {
                for (Map<String, String> sessionInfo : sessionInfoList) {
                    String sysTestsId = sessionInfo.get("Tests");// 从系统中获取到的关联 ID。可能是Test Case、Test Suite
                    String sessionId = sessionInfo.get("ID");
                    String state = sessionInfo.get("State");
                    List<String> caseList = new ArrayList<>();
                    if (sysTestsId != null && !"".equals(sysTestsId)) {
                        String[] sysTestsIdArr = sysTestsId.split(",");
                        List<Map<String, String>> testsList = cmd.findItemsByIDs(Arrays.asList(sysTestsIdArr),
                                Arrays.asList("ID,Type"));
                        for (Map<String, String> testMap : testsList) {
                            String type = testMap.get("Type");
                            String ID = testMap.get("ID");
                            if ("Test Case".equals(type)) {
                                caseList.add(ID);
                            } else {
                                List<String> allContains = cmd.allContents(ID);
                                if (allContains.size() > 0) {
                                    caseList.addAll(allContains);
                                }
                            }
                        }
                    }
                    sessionInfoRecord.put(sessionId, caseList);
                    sessionStateRecord.put(sessionId, state);
                }
            }

        }
        for (int i = 0; i < testInfoDatas.size(); i++) {
            boolean hasError = false;//校验出错误
            StringBuffer errorMessage = new StringBuffer();
            Map<String, Object> rowMap = testInfoDatas.get(i);
            String caseID = (String) rowMap.get(ID_FIELD);
            newMap = new HashMap<String, Object>();
            if (caseID != null && !"".equals(caseID)) {
                newMap.put("ID", caseID);
            }
            for (String header : caseFields) {
                Map<String, String> fieldConfig = headerConfig.get(header);
                if (fieldConfig != null) {
                    String field = fieldConfig.get("field");
                    String value = (String) rowMap.get(field);
                    if (!"-".equals(field) && value != null && !"".equals(value)) {
                        String message = checkFieldValue(header, field, value, cmd);//校验Test Case字段值
                        if (message == null || "".equals(message)) {
                            // 在此已经判断用户是否存在  ， 若存在 IS_USER 标识为 ture , 若不存在为 false
                            if (IS_USER) {
                                // list.get(p).toString()
                                // 判断导入的user类型的数据格式是不是 : 用户(ID) 是的话截取 ()内ID 。
                                int leftIndex = -1;
                                int rightIndex = -1;
                                boolean endFormat = false;
                                if (value.indexOf("(") > -1) {
                                    leftIndex = value.indexOf("(");
                                } else if (value.indexOf("（") > -1) {
                                    leftIndex = value.indexOf("（");
                                }
                                if (value.indexOf(")") > -1) {
                                    rightIndex = value.indexOf(")");
                                    endFormat = value.endsWith(")");
                                } else if (value.indexOf("）") > -1) {
                                    rightIndex = value.indexOf("）");
                                    endFormat = value.endsWith("）");
                                }
                                if (leftIndex > 0 && rightIndex > 0 && endFormat) {
                                    String userId = value.substring(leftIndex + 1, rightIndex);
                                    if (userId.matches("d{0,9}") || userId.matches("d{0,9}")
                                            || userId.matches("d{0,9}") || userId.matches("d{0,9}")) {
                                        // 判断里面ID格式是不是 GW + 数字  是的话在之前查询的数据获取值
                                        newMap.put(field, userId);
                                    }

                                } else if (value.matches("d{0,9}") || value.matches("d{0,9}")
                                        || value.matches("d{0,9}") || value.matches("d{0,9}")) { // 判断如果不是用户(ID)的格式 , 在判断是不是直接填写的ID GW+数字 格式。

                                    newMap.put(field, value);

                                } else {
                                    errorMessage.append(" Field [" + field + "]  data format should be \"name(Login ID)\" or \"Login ID\" \n");
                                    hasError = true;
                                }
                                IS_USER = false;
                            } else if (RELATIONSHIP_MISTAKEN) { //如果是Relationship类型的字段，并且数字前面带[] ，就将中括号去掉
                                value = value.substring(1, value.length() - 1);//
                                newMap.put(field, value);
                                RELATIONSHIP_MISTAKEN = false;
                            } else {
                                newMap.put(field, value);
                            }

                        } else {
                            errorMessage.append("第 " + (i + 1) + " 条测试用例: ").append(message).append("\n");
                            hasError = true;
                        }
                    }
                }
            }
            if (hasError) {
                allMessage.append(errorMessage);
                continue;
            }
            if (rowMap.containsKey(TEST_STEP)) {//Test Case包含有 Test Step信息
                Object steps = rowMap.get(TEST_STEP);
                if (steps instanceof List) {
                    List<Map<String, String>> currentSteps = (List<Map<String, String>>) steps;
                    if (!currentSteps.isEmpty()) {//Test Case包含有 Test Step信息
                        for (Map<String, String> stepMap : currentSteps) {
                            for (String header : stepFields) {
                                Map<String, String> fieldConfig = headerConfig.get(header);
                                if (fieldConfig != null) {
                                    String field = fieldConfig.get("field");
                                    String value = (String) rowMap.get(field);
                                    if (!"-".equals(field) && value != null && !"".equals(value)) {
                                        String message = checkFieldValue(header, field, value, cmd);//校验Test Step字段值
                                        if (message != null && !"".equals(message)) {
                                            allMessage.append(message);
                                        }
                                    }
                                }
                            }
                        }
                        newMap.put(TEST_STEP, currentSteps);
                    }
                }
            }
            if (rowMap.containsKey(TEST_RESULT)) {//Test Case包含有 Test Result信息
                Object results = rowMap.get(TEST_RESULT);
                if (results instanceof List) {
                    List<Map<String, String>> currentResults = (List<Map<String, String>>) results;
                    if (!currentResults.isEmpty()) {//Test Case包含有 Test Result信息
                        if (currentResults.size() > sessionIds.size()) {//
                            allMessage.append("第 " + (i + 1) + " 条测试用例的测试结果，必须有足够的Test Session关联，且Test Session处于In Testing状态。").append("\n");
                        } else {
                            for (int j = 0; j < currentResults.size(); j++) {//循环校验Test Result信息
                                Map<String, String> map = currentResults.get(j);
                                String sessionId = sessionIds.get(j);//一个Test Session 对应一轮 测试结果
                                String sessionState = sessionStateRecord.get(sessionId);
                                List<String> caseList = sessionInfoRecord.get(sessionId);
                                if (!SESSION_STATE.equals(sessionState)) {//校验Test Session状态
                                    allMessage.append("第 " + (i + 1) + " 条测试用例关联的Test Session未处于In Testing状态。不能编辑测试结果。").append("\n");
                                }
                                if (caseList != null) {//校验输入测试结果的列，有没有与Test Session关联
                                    if (!caseList.contains(caseID)) {
                                        allMessage.append("第 " + (i + 1) + " 条测试用例，" + "第" + j + "轮"
                                                + "Test Session与当前测试用例未建立关联关系！ \n");
                                    }
                                }
                                Set<Entry<String, String>> entrySet = map.entrySet();
                                for (Entry<String, String> entry : entrySet) {
                                    String displayKey = entry.getKey();
                                    String key = resultFieldsMap.get(displayKey);
                                    String fieldType = FIELD_TYPE_RECORD.get(key);
                                    String value = entry.getValue();
                                    if (PICK_FIELD_RECORD.containsKey(key)) {
                                        List<String> includes = PICK_FIELD_RECORD.get(key);
                                        if (!includes.contains(value)) {
                                            allMessage.append("第 " + (i + 1) + " 条测试用例 ").append(String.format("字段【%s】不正确，合法值范围【%s】\r\n", key, StringUtil.join(",", includes)));
                                        }
                                    } else if ("date".equals(fieldType)) {
                                        String msg = checkDate(value);
                                        if (msg != null && msg.length() > 0) {
                                            allMessage.append(msg);
                                        } else {
                                            SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                                            SimpleDateFormat valueSdf = new SimpleDateFormat("MMM d, yyyy h:mm:ss a", Locale.ENGLISH);
                                            Date date = sdf2.parse(value);
                                            value = valueSdf.format(date);
                                            map.put(displayKey, value);
                                        }
                                    }
                                }
                            }
                        }
                        newMap.put(TEST_RESULT, currentResults);
                    }
                }
            }
            if (rowMap.containsKey(DEFECT)) {
                Object defect = rowMap.get(DEFECT);
                if (defect instanceof List) {
                    List<Map<String, String>> currentDefect = (List<Map<String, String>>) defect;
                    if (!currentDefect.isEmpty()) {
                        for (Map<String, String> stepMap : currentDefect) {
                            for (String header : defectFields) {
                                Map<String, String> fieldConfig = headerConfig.get(header);
                                if (fieldConfig != null) {
                                    String field = fieldConfig.get("field");
                                    String value = (String) rowMap.get(field);
                                    if (!"-".equals(field) && value != null && !"".equals(value)) {
                                        String message = checkFieldValue(header, field, value, cmd);//校验Test Step字段值
                                        if (message != null && !"".equals(message)) {
                                            allMessage.append(message);
                                        }
                                    }
                                }
                            }
                        }
                        newMap.put(DEFECT, currentDefect);
                    }
                }
            }
            testInfoList.add(newMap);
        }
        TestInfoImportUI.logger.info("End Deal Excel Data ,Total Data is :" + testInfoDatas.size());
        resultDataList.add(testInfoList);
        allMessage.append(cmd.checkIssueType(sessionIds, TEST_SESSION));
        errorRecord.put("error", allMessage.toString());
        return resultDataList;
    }

    /**
     * Description 开始导入数据
     *
     * @param data
     * @param cmd
     * @param importType
     * @param shortTitle
     * @param project
     * @param testSuiteID
     * @throws Exception
     */
    @SuppressWarnings("unchecked")
    @Override
    public void startImport(List<List<Map<String, Object>>> dataList, IntegrityUtil cmd, String importType, String shortTitle, String project, String testSuiteID) throws APIException {
        // 删除Token
        // TestCaseImport.TOKEN = null;
        // 下面List用于收集操作信息，用于统计
        List<String> caseUpdate = new ArrayList<String>(), caseCreate = new ArrayList<String>(),
                stepUpdate = new ArrayList<String>(), stepCreate = new ArrayList<String>();
        List<String> caseUpdateF = new ArrayList<String>(), caseCreateF = new ArrayList<String>(),
                stepUpdateF = new ArrayList<String>(), stepCreateF = new ArrayList<String>();

//		int totalSheetNum = data.size();
        boolean hasStep = false;
        // 遍历信息
        List<Map<String, Object>> sessionData = dataList.get(0);//Session 信息数据
        List<Map<String, Object>> datas = dataList.get(1);//Case 信息数据
        List<Map<String, String>> sessionList = new ArrayList<>();
        List<String> sessionIdList = new ArrayList<>();
        boolean createTest = false;
        if (datas.isEmpty()) {
            return;
        }
        try {
            if (testSuiteID == null || "".equals(testSuiteID)) {
                Map<String, String> docInfo = new HashMap<String, String>();
                docInfo.put("Document Short Title", shortTitle);
                docInfo.put("Project", project);
                docInfo.put("State", INIT_DOC_STATE);
                if (IMPORT_DOC_TYPE.endsWith("Document"))
                    docInfo.put("Shared Category", "Document");
                else if ("Test Suite".equals(IMPORT_DOC_TYPE))
                    docInfo.put("Shared Category", "Suite");
                testSuiteID = cmd.createDocument(IMPORT_DOC_TYPE, docInfo);
                createTest = true;
            } else {
                Map<String, String> map = cmd.getItemById(testSuiteID, Arrays.asList("State"));
                if (IMPORT_DOC_TYPE.equals(map.get("State"))) {
                    DOC_STATE = true;
                }
            }
            if (!createTest) {
                project = cmd.getItemByIds(Arrays.asList(testSuiteID), Arrays.asList("Project")).get(0).get("Project");
            }
            String parentId = testSuiteID;//
            int totalCaseNum = datas.size();
            List<String> caseIds = new ArrayList();//记录所有的Test Case ID
            TestInfoImportUI.logger.info("Start Import Data , all Data size is :" + totalCaseNum);
            /** 1. 处理Test Session信息*/

            if (sessionData != null && !sessionData.isEmpty()) {
                for (Map<String, Object> sessionMap : sessionData) {
                    String sessionId = dealSessionInfo(sessionMap, project, cmd);
                    sessionMap.put(ID_FIELD, sessionId);//将
                    sessionIdList.add(sessionId);
                }
            }
            if (!sessionIdList.isEmpty()) {
                sessionList = cmd.getItemByIds(sessionIdList, Arrays.asList("ID", "Tests", "State"));
            }
            /** 处理Test Session信息*/
            /** 2. 处理Test Case 、Test Step、Test Result、 Defect信息*/
            for (int index = 0; index < totalCaseNum; index++) {
                Map<String, Object> testCaseData = datas.get(index);
                //把Test Step信息取出
                List<Map<String, String>> testStepData = null;
                if (testCaseData.get(TEST_STEP) != null) {
                    testStepData = (List<Map<String, String>>) testCaseData.get(TEST_STEP);
                    testCaseData.remove(TEST_STEP);
                }
                // 把Test Result信息获取出来
                List<Map<String, String>> resultList = null;
                if (testCaseData.get(TEST_RESULT) != null) {
                    resultList = (List<Map<String, String>>) testCaseData.get(TEST_RESULT);
                    testCaseData.remove(TEST_RESULT);
                }
                // 把Defect信息获取出来
                List<Map<String, String>> defectList = null;
                if (testCaseData.get(DEFECT) != null) {
                    defectList = (List<Map<String, String>>) testCaseData.get(DEFECT);
                    testCaseData.remove(DEFECT);
                }
                logger.info("Now Deal row " + index + " data");
                int caseNum = index + 1;
                String caseId = null;
                boolean newCase = false;
                if (testCaseData.containsKey("ID")) {
                    caseId = testCaseData.get("ID").toString();
                }
                if (caseId == null || "".equals(caseId)) {
                    newCase = true;
                    TestInfoImportUI.showLogger(" \tStart to Create " + importType);
                } else {
                    TestInfoImportUI.showLogger(" \tStart to deal " + importType + "  : " + caseId);
                }
                Map<String, String> newTestCaseData = new HashMap<>();
                List<String> newRelatedStepIds = new ArrayList<>();


                // 1. 处理Test Case的信息(更新或创建，不包括创建)
                String beforeId = "last";//涉及结构
                parentId = testSuiteID;
                caseId = this.getTestCase(parentId, newTestCaseData, testCaseData, project, cmd, caseId, beforeId,
                        caseCreate, caseCreateF, caseUpdate, caseUpdateF, importType);
                testCaseData.put("ID", caseId);
                /*if (!newCase) {//如果是更新Test Case。则查询系统已关联的Test Step信息 进行更新
                    Map<String, String> relatedSteps = cmd.getItemById(caseId, Arrays.asList("ID", "Test Steps"));
                    List<String> stepIds = Arrays.asList(relatedSteps.get("Test Steps").split(","));
                    for (int k = 0; k < stepIds.size(); k++) {//将查询出来的ID 写入Map，方便更新
                        String stepId = stepIds.get(k);
                        Map<String, String> stepMap = testStepData.get(k);
                        stepMap.put("ID", stepId);
                    }
                }*/

                // 2. 处理Test Step信息(更新创建或删除)，newTestCaseData中
                //可以把结果取出来  Cycle Result，在Test Step Info不再进行更新Cycle Result信息
                // 以 Case ID - [{stepId,Cycle Result,Cycle2 Result},{stepId,Cycle Result,Cycle2 Result}]
                List<Map<String, String>> maps = deepCopyList(testStepData);
                if (testStepData != null && !testStepData.isEmpty()) {
                    this.getTestStep(newRelatedStepIds, testStepData, testCaseData, project, cmd, stepCreate,
                            stepCreateF, stepUpdate, stepUpdateF);
                    hasStep = true;
                }
                // 3. 关联Test Case与Test Step
                if (testStepData != null && !testStepData.isEmpty() && newRelatedStepIds.size() > 0) {
                    this.relatedCaseAndStep(caseId, newRelatedStepIds, cmd);
                }

                // 4. 导入测试结果
                //把 前面记录的 Case - Test Step Result信息传递到方法内
                if (resultList != null && !resultList.isEmpty()) {
                    dealTestResults(sessionList, resultList, cmd, caseId, maps);
                }

                // 5. Defect信息处理
                TestInfoImportUI.logger.info("defectList" + defectList);
                if (defectList != null && !defectList.isEmpty()) {
                    dealDefect(defectList, cmd, caseId, project);
                }
                caseIds.add(caseId);
                TestInfoImportUI.showProgress(1, datas.size(), caseNum, totalCaseNum);
            }
            /**  判断Test Session状态，如果处于初始状态，关联Test Cases*/

            dealSessionCaseRelation(sessionList, caseIds, cmd);
        } catch (APIException e) {
            logger.error(APIExceptionUtil.getMsg(e));
        } catch (Exception e) {
            logger.error(e.getMessage());
        }

        TestInfoImportUI.showLogger("End to deal Test : " + testSuiteID);
        TestInfoImportUI.showLogger("==============================================");
        TestInfoImportUI.showLogger("Deal Test Session : success (" + sessionList.size() + ")");
        TestInfoImportUI.showLogger("Create " + CONTENT_TYPE + ": success (" + caseCreate.size() + "," + caseCreate + "), failed ("
                + caseCreateF.size() + ")");
        TestInfoImportUI.showLogger("Update " + CONTENT_TYPE + ": success (" + caseUpdate.size() + "," + caseUpdate + "), failed ("
                + caseUpdateF.size() + "," + caseUpdateF + ")");
        if (hasStep) {
            TestInfoImportUI.showLogger("Create Test Step: success (" + stepCreate.size() + "," + stepCreate + "), failed ("
                    + stepCreateF.size() + ")");
            TestInfoImportUI.showLogger("Update Test Step: success (" + stepUpdate.size() + "," + stepUpdate + "), failed ("
                    + stepUpdateF.size() + "," + stepUpdateF + ")");
        }
    }

    public static <T> List<T> deepCopyList(List<T> src) throws IOException, ClassNotFoundException {
        ByteArrayOutputStream byteOut = new ByteArrayOutputStream();
        ObjectOutputStream out = new ObjectOutputStream(byteOut);
        out.writeObject(src);
        ByteArrayInputStream byteIn = new ByteArrayInputStream(byteOut.toByteArray());
        ObjectInputStream in = new ObjectInputStream(byteIn);
        @SuppressWarnings("unchecked")
        List<T> dest = (List<T>) in.readObject();
        return dest;
    }

    /**
     * 处理Test Session与Test Case关系
     *
     * @param sessionList
     * @param caseList
     * @param cmd
     * @throws APIException
     */
    private void dealSessionCaseRelation(List<Map<String, String>> sessionList, List<String> caseList, IntegrityUtil cmd) throws APIException {
        if (sessionList == null || sessionList.isEmpty()) {
            return;
        }
        for (Map<String, String> sessionMap : sessionList) {
            String sessionState = sessionMap.get("State");
            String id = sessionMap.get(ID_FIELD);
            if (SESSION_INIT_STATE.equals(sessionState)) {
                String tests = sessionMap.get("Tests");
                String[] testArr = tests.split(",");
                StringBuffer removeTests = new StringBuffer();//需要移除的关联Case
                List<String> tempCaseList = new ArrayList<>(caseList);
                for (int i = 0; i < testArr.length; i++) {
                    String caseId = testArr[i];
                    if (tempCaseList.contains(caseId)) {
                        tempCaseList.remove(caseId);
                    } else if (caseId != null && !"".equals(caseId)) {
                        removeTests.append(caseId).append(",");
                    }
                }
                if (tempCaseList.size() > 0) {//需要添加的关系
                    StringBuffer addTests = new StringBuffer();
                    for (String caseId : tempCaseList) {
                        addTests.append(caseId).append(",");
                    }
                    cmd.addRelationship(id, "Tests", addTests.toString());
                }
                if (removeTests.toString().length() > 0) {
                    cmd.removeRelationship(id, "Tests", removeTests.toString());
                }
            }
        }
    }

    /**
     * 将Test Case与Test Step的关联关系进行更新
     *
     * @param caseId
     * @param newRelatedStepIds
     * @param cmd
     * @throws APIException
     */
    public void relatedCaseAndStep(String caseId, List<String> newRelatedStepIds, IntegrityUtil cmd) throws APIException {
        if (caseId != null && caseId.length() > 0) {
            StringBuffer sb = new StringBuffer();
            for (String step : newRelatedStepIds) {
                sb.append(sb.toString().length() > 0 ? "," + step : step);
            }
            Map<String, String> map = new HashMap<>();
            map.put("Test Steps", sb.toString());
            cmd.editissue(caseId, map);
        }
    }

    /**
     * 创建或更新Test Case
     *
     * @param documentId      Suite ID
     * @param newTestCaseData 新的Case信息集合
     * @param caseMap         原有的Case信息集合
     * @param project         Suite的Project
     * @param cmd
     * @param caseId
     * @param beforeId
     * @param caseCreate
     * @param caseCreateF
     * @param caseUpdate
     * @param caseUpdateF
     * @throws Exception
     */
    public String getTestCase(String parentId, Map<String, String> newTestCaseData, Map<String, Object> caseMap,
                              String project, IntegrityUtil cmd, String caseId, String beforeId, List<String> caseCreate,
                              List<String> caseCreateF, List<String> caseUpdate, List<String> caseUpdateF, String importType) throws APIException {

        logger.info("Data Of " + CONTENT_TYPE + " ID [" + caseId + "]");
        // 需修改
        for (Entry<String, Object> entrty : caseMap.entrySet()) {
            String field = entrty.getKey();
            Object value = entrty.getValue();
            if (value != null && value.toString().length() > 0) {
                newTestCaseData.put(field, value.toString());
            }
        }
        String containedBy = newTestCaseData.get("Contained By");
        newTestCaseData.remove("ID");
        newTestCaseData.remove("Document ID");
        newTestCaseData.remove("Test Step");
        newTestCaseData.remove("Contained By");
        if (caseId == null || caseId.length() == 0) {
            // 创建Test Case
            try {
                if (containedBy != null && !"".equals(containedBy) && containedBy.matches("[0-9]*")) {
                    parentId = containedBy;
                }
                newTestCaseData.put("Project", project);
                newTestCaseData.put("Category", DEFAULT_CATEGORY);
                newTestCaseData.put("State", INIT_CONTENT_STATE);
                caseId = cmd.createContent(parentId, newTestCaseData, CONTENT_TYPE, beforeId);
                caseCreate.add(caseId);
                TestInfoImportUI.showLogger(" \tSuccess to create " + CONTENT_TYPE + " : " + caseId);
            } catch (APIException e) {
                caseCreateF.add(caseId);
                logger.error(APIExceptionUtil.getMsg(e));
                TestInfoImportUI.showLogger(" \tFailed to create " + CONTENT_TYPE + " : " + caseId);
                logger.error("Failed to create test case : " + ExceptionUtil.catchException(e));
            }
        } else {
            if (DOC_STATE) {
                // 更新Test Case
                // 遍历出所有 overRide为 true 的字段，
                Map<String, Map<String, String>> fieldMaps = headerConfig;
                Collection<Map<String, String>> fieldMapValues = fieldMaps.values();
                List<String> fields = new ArrayList<String>();
                for (Map<String, String> values : fieldMapValues) {
                    if (values.get("overRide").equals("false")) {
                        fields.add(values.get("field"));
                    }
                }
                // 然后调用 mks命令查询出导入的 所有 ids 的内容。判断当前为true字段是否有值 , getItemByIds(List<String> ids,List<String> field) 此方法通过Id 获取字段的值
                List<String> ids = new ArrayList<String>();
                ids.add(caseId);
                List<Map<String, String>> data = cmd.getItemByIds(ids, fields);
                Map<String, String> dataMap = data.get(0);
                for (String field : fields) {
                    String fieldValue = dataMap.get(field);
                    // 有  ： 不更新       没有 ： 更新
                    if (!"".equals(fieldValue) && null != fieldValue) {
                        newTestCaseData.remove(field);
                    }
                }
                // 判断当前条目中是否 含有 Text 字段，如果有，检查此字段是否可以编辑更新（含有Text字段的条目，是否可以更新，在XML里有属性OnlyCreate 规定  。false为可编辑，true为不可编辑）
                checkOnlyCreate(newTestCaseData, importType);
                try {
                    cmd.editissue(caseId, newTestCaseData);
                    caseUpdate.add(caseId);
                    // 1.更新顺序
                    if (beforeId != null && !"".equals(beforeId)) {
                        cmd.moveContent(parentId, beforeId, caseId);
                    }
                    TestInfoImportUI.showLogger(" \tSuccess to update Test Case : " + caseId);
                } catch (APIException e) {
                    caseUpdateF.add(caseId);
                    logger.error(APIExceptionUtil.getMsg(e));
                    TestInfoImportUI.showLogger(" \tFailed to update Test Case : " + caseId);
                    logger.error("Failed to edit test case : " + ExceptionUtil.catchException(e));
                }
            }
        }
        return caseId;
    }

    /**
     * 处理test Session信息。 新建或者更新
     *
     * @param sessionInfo
     * @param cmd
     * @return
     * @throws APIException
     */
    @SuppressWarnings("unused")
    private String dealSessionInfo(Map<String, Object> sessionInfo, String project, IntegrityUtil cmd) throws APIException {
        Object sessionId = sessionInfo.get(ID_FIELD);
        Map<String, String> newInfo = new HashMap<>();
        for (Entry<String, Object> entrty : sessionInfo.entrySet()) {
            String field = entrty.getKey();
            if (ID_FIELD.equals(field) || "-".equals(field)) {//不更新ID列
                continue;
            }
            Object value = entrty.getValue();
            if (value != null && value.toString().length() > 0) {
                newInfo.put(field, value.toString());
            }
        }
        if (sessionId == null || "".equals(sessionId.toString())) {
            newInfo.put("Project", project);
            newInfo.put("State", SESSION_INIT_STATE);
            sessionId = cmd.createIssue(TEST_SESSION, newInfo, null);
        } else {
            cmd.editissue(sessionId.toString(), newInfo);
        }
        return sessionId.toString();
    }

    /**
     * 检测当前要更新Case里面有没有 Text 字段 ， 并且判断该字段是否可以编辑        xml 中有 onlyCreate 属性规定是否可以更新
     *
     * @param newTestCaseData
     * @param importType
     */
    private void checkOnlyCreate(Map<String, String> newTestCaseData, String importType) {
        Collection<Map<String, String>> values = headerConfig.values();
        for (Map<String, String> map : values) {
            if (map.get("onlyCreate") != null) {
                boolean onlyCreate = Boolean.valueOf(map.get("onlyCreate"));
                if (onlyCreate) {
                    String field = map.get("field");
                    newTestCaseData.remove(field);
                }
            }
        }
    }

    /**
     * 先处理Test Step信息(更新创建或删除)，
     * 遍历得到OPERATING_ACTION和EXPECTED_RESULTS信息塞入newTestCaseData中, 并将创建和更新的Step
     * ID塞于newRelatedStepIds中
     *
     * @param newTestCaseData   新的Case信息集合
     * @param newRelatedStepIds 创建Test Step的集合
     * @param caseMap           原有Case信息集合
     * @param project           Suite的Project信息
     * @param cmd
     * @param stepCreate
     * @param stepCreateF
     * @param stepUpdate
     * @param stepUpdateF
     */
    @SuppressWarnings("unchecked")
    public void getTestStep(List<String> newRelatedStepIds, List<Map<String, String>> testStepData,
                            Map<String, Object> caseMap, String project, IntegrityUtil cmd, List<String> stepCreate,
                            List<String> stepCreateF, List<String> stepUpdate, List<String> stepUpdateF) {

        int i = 1;
        if (testStepData != null && testStepData.size() > 0) {
            TestInfoImportUI.showLogger(" \t\tHas Test Step size  : " + testStepData.size());
            for (Map<String, String> stepMap : testStepData) {
                String stepId = stepMap.get("ID");
                stepMap.remove("ID");
                // 处理Step Order
                if (stepId == null || stepId.length() == 0) {
                    // 创建Test Step，并关联Test Case
                    try {
                        stepMap.put("Project", project);
                        stepMap.put("State", INIT_STATE);
                        stepId = cmd.createIssue(TEST_STEP, stepMap, null);
                        stepCreate.add(stepId);
                        TestInfoImportUI.showLogger(" \t\tSuccess to create Test Step " + i + ", " + stepId);
                    } catch (APIException e) {
                        stepCreateF.add(stepId);
                        logger.error(APIExceptionUtil.getMsg(e));
                        TestInfoImportUI.showLogger(" \t\tFailed to create Test Step");
                        logger.error("Failed to create test step : " + ExceptionUtil.catchException(e));
                    }
                } else {
                    try {
                        cmd.editissue(stepId, stepMap);
                        stepUpdate.add(stepId);
                        TestInfoImportUI.showLogger(" \t\tSuccess to update Test Step " + i + ", " + stepId);
                    } catch (APIException e) {
                        stepUpdateF.add(stepId);
                        logger.error(APIExceptionUtil.getMsg(e));
                        TestInfoImportUI.showLogger(" \t\tFailed to update Test Step " + i + ", " + stepId);
                        logger.error("Failed to edit test step : " + ExceptionUtil.catchException(e));
                    }
                }
                newRelatedStepIds.add(stepId);
                i++;
            }
        }
    }

    /**
     * 处理导入结果
     *
     * @param caseMap
     * @param cmd
     */
    public void dealTestResults(List<Map<String, String>> sessionList, List<Map<String, String>> resultDatas, IntegrityUtil cmd, String caseID, List<Map<String, String>> testStepDatas) throws APIException {
        if (resultDatas != null && !resultDatas.isEmpty()) {
            for (int i = 0; i < resultDatas.size(); i++) {
                Map<String, String> result = resultDatas.get(i);
                Map<String, String> sessionMap = sessionList.get(i);
                String sessionId = sessionMap.get(ID_FIELD);
                List<Map<String, Object>> result1 = cmd.getResult(sessionId, caseID, "Test Case");
                List<Map<String, String>> stepData = new ArrayList<>();
                if (!testStepDatas.isEmpty()) {
                    stepData = getStepData(i, testStepDatas);
                }
                if (result1.size() > 0) {
                    cmd.editResult(sessionId, caseID, result, stepData);
                } else {
                    cmd.createResult(sessionId, caseID, result, stepData);
                }
            }
        }
    }

    private List<Map<String, String>> getStepData(int index, List<Map<String, String>> testStepDatas) {
        ArrayList<Map<String, String>> objects = new ArrayList<>();
        for (Map<String, String> map : testStepDatas) {
            HashMap<String, String> step = new HashMap<>();
            switch (index) {
                case 0:
                    step.put("ID", map.get("ID"));
                    step.put("Cycle Verdict", map.get("Cycle Verdict"));
                    break;
                case 1:
                    step.put("ID", map.get("ID"));
                    step.put("Cycle Verdict", map.get("Cycle2 Verdict"));
                    break;
                case 2:
                    step.put("ID", map.get("ID"));
                    step.put("Cycle Verdict", map.get("Cycle3 Verdict"));
                    break;
                case 3:
                    step.put("ID", map.get("ID"));
                    step.put("Cycle Verdict", map.get("Cycle4 Verdict"));
                    break;
                case 4:
                    step.put("ID", map.get("ID"));
                    step.put("Cycle Verdict", map.get("Cycle5 Verdict"));
                    break;

            }
            objects.add(step);
        }
        return objects;
    }

    public void dealDefect(List<Map<String, String>> defectList, IntegrityUtil cmd, String caseID, String project) throws APIException {
        for (int i = 0; i < defectList.size(); i++) {
            Map<String, String> defect = defectList.get(i);
            if (defect.get("id") != null) {
                cmd.editIssue(defect.get("id"), defect, null);
                TestInfoImportUI.logger.info("defect更新成功");
            } else {
                defect.put("Project", project);
                defect.put("Summary", String.format("case%s,bug", caseID));
                String defectId = cmd.createIssue("Defect", defect, null);
                cmd.addRelationship(defectId, "Blocks", caseID);
                TestInfoImportUI.logger.info("defect新增成功" + defectId);
            }
        }
    }
}
