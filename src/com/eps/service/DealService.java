package com.eps.service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

import javax.swing.JOptionPane;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import com.eps.util.MKSCommand;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import com.mks.api.response.APIException;
import com.eps.ui.TestResultExportUI;
import com.eps.util.GenerateXmlUtil;

@SuppressWarnings("all")
public class DealService {
    private static final String POST_CONFIG_FILE = "FieldMapping.xml";
    private static final String CATEGORY_CONFIG_FILE = "Category.xml";
    Map<String, List<Map<String, String>>> xmlConfig = new HashMap<>();
    Map<String, List<String>> contentColumns = new HashMap<>();
    Map<String, List<String>> contentSessionColumns = new HashMap<>();
    private List<String> allHeaders = new ArrayList<>();
    private List<String> contentHeaders = new ArrayList<>();
    private List<String> stepHeaders = new ArrayList<>();
    private List<String> realStepFields = new ArrayList<>();
    private List<String> sessionFields = new ArrayList<>();
    private List<String> resultHeaders = new ArrayList<>();
    private Map<String, String> inputCountMap = new HashMap<String, String>();
    public static final String TEST_SUITE = "Test Suite";
    public static final String TEST_SESSION = "Test Session";
    private String SEQUENCE = "Sequence";
    private List<List<Object>> datas = new ArrayList<>();
    private List<List<List<Object>>> allDatas = new ArrayList<>();
    private List<List<String>> listHeaders = new ArrayList<>();
    private List<CellRangeAddress> cellList = new ArrayList<>();
    public static final Map<String, String> HEADER_MAP = new HashMap<String, String>();
    public static final Map<String, String> HEADER_COLOR_RECORD = new HashMap<String, String>();
    private Map<String, String> stepHeaderMap = new HashMap<String, String>();
    private Map<String, String> resultHeaderMap = new HashMap<String, String>();
    private List<String> headers = new ArrayList<>();// 第一行标题
    private List<String> headerTwos = new ArrayList<>();// 第二行标题
    private List<String> sessionHeader = new ArrayList<>();// session行标题
    private Map<String, List<String>> allSheetHeaders = new HashMap<>();
    private List<Object> data = new ArrayList<>();
    private static Integer CycleTest = 1;
    private List<String> testResultHeaders = new ArrayList<>();
    private Map<String, Object> testResultData = new HashMap<>();
    private String suitId;
    public static final List<String> CURRENT_CATEGORIES = new ArrayList<String>();// 记录导入对象的正确Category

    public static final Map<String, List<String>> PICK_FIELD_RECORD = new HashMap<String, List<String>>();

    public static final Map<String, String> FIELD_TYPE_RECORD = new HashMap<String, String>();

    public static final Map<String, String> IMPORT_DOCUMENT_TYPE = new HashMap<String, String>();

    public static final Map<String, String> USER_MAP = new HashMap<String, String>();

    public static final Map<String, String> All_USER_MAP = new HashMap<>();

    public static final Map<String, List<String>> seesionColumns = new HashMap<>();

    private List<String> stepHeader = new ArrayList<>();
    private List<String> testResultFiled = new ArrayList<>();
    private List<String> defectFiled = new ArrayList<>();
    private List<String> sessionId = new ArrayList<>();
    private List<String> headerTwofiled = new ArrayList<>();
    private List<String> testStepfiled = new ArrayList<>();
    private Map<String, String> nameToFiled = new HashMap<>();

    /**
     * 解析XML
     *
     * @param project
     * @throws APIException
     */
    public List<String> parseXML() throws APIException {
        List<String> exportTypes = new ArrayList<String>();// 导出类型list
        try {
            TestResultExportUI.logger.info("start to parse xml : " + POST_CONFIG_FILE);

            Document document = DocumentBuilderFactory.newInstance().newDocumentBuilder()
                    .parse(DealService.class.getClassLoader().getResourceAsStream(POST_CONFIG_FILE));
            Element root = document.getDocumentElement();

            if (root != null) {
                NodeList eleList = root.getElementsByTagName("importType");// 获取导出类型
                if (eleList != null) {
                    List<Map<String, String>> allFields = new ArrayList<>(); // 模板field
                    List<String> ptcFields = new ArrayList<>();// 存放系统自带field
                    List<String> ptcsessionFields = new ArrayList<>();// 存放系统自带sessionfield
                    for (int i = 0; i < eleList.getLength(); i++) {// 获取当前类型节点相关属性值
                        Element item = (Element) eleList.item(i);
                        String typeName = item.getAttribute("name");
                        exportTypes.add(typeName);
                        String documentType = item.getAttribute("type");
                        IMPORT_DOCUMENT_TYPE.put(typeName, documentType);
                        parseData(item, allFields, ptcFields, ptcsessionFields);// 解析数据，往Excel模板汇中存放field
                    }
                    xmlConfig.put(TEST_SUITE, allFields);
                    contentColumns.put(TEST_SUITE, ptcFields);
                    contentSessionColumns.put(TEST_SESSION, ptcsessionFields);
                }
            }
        } catch (ParserConfigurationException e) {
            TestResultExportUI.logger.error("parse config file exception", e);
        } catch (SAXException e) {
            TestResultExportUI.logger.error("get config file exception", e);
        } catch (IOException e) {
            TestResultExportUI.logger.error("io exception", e);
        } finally {
            TestResultExportUI.logger.info(" xmlConfig: " + xmlConfig + " \n, the ptcTestCaseColumns: " + contentColumns);
            return exportTypes;
        }
    }

    /**
     * 解析数据，Excel模板中存放field
     *
     * @param eleList
     * @param list
     * @param ptcFields
     */
    private void parseData(Element exportType, List<Map<String, String>> list, List<String> ptcFields, List<String> ptcsessionFields) {
        NodeList nodeFields = exportType.getElementsByTagName("excelField");
        for (int j = 0; j < nodeFields.getLength(); j++) {
            Map<String, String> map = new HashMap<>();// 存放所有fields Excel模板
            Element fields = (Element) nodeFields.item(j);
            String field = fields.getAttribute("field");
            String type = fields.getAttribute("type");
            String category = fields.getAttribute("category");

            String fieldName = fields.getAttribute("name");
            String titleColor = fields.getAttribute("titleColor");// 获取标题颜色标识
            if (!field.equals("-") && !type.equals("Test Result") && !type.equals("Test Step") && !"Test Session".equals(type)) {
                ptcFields.add(field);// 如果模板中符合以上情况，则直接将field存放到系统自带field的list中
            }
            if (category != null && !"".equals(category)) {
                headerTwofiled.add(field);
                nameToFiled.put(fieldName, field);
            }
            if ("Test Result".equals(category)) {
                testResultFiled.add(field);
            }
            if ("Defect".equals(category)) {
                defectFiled.add(field);
            }
            if ("Test Step".equals(category)) {
                stepHeader.add(fieldName);
                stepHeaderMap.put(fieldName, field);
                testStepfiled.add(field);
            }
            if ("Test Session".equals(type)) {
                sessionHeader.add(fieldName);
            }

            HEADER_MAP.put(fieldName, field);
            if ("Test Step".equals(type)) {
                stepHeaders.add(fieldName);
                realStepFields.add(fieldName);
                HEADER_COLOR_RECORD.put(fieldName, titleColor);
                if (!headers.contains("Test Steps")) {
                    headers.add("Test Steps");
                } else
                    headers.add("-");
                headerTwos.add(fieldName);
            } else if ("Test Session".equals(type)) {
                sessionFields.add(field);
            } else if ("Test Result".equals(type)) {
                if (!resultHeaders.contains(fieldName))
                    resultHeaders.add(fieldName);
                resultHeaderMap.put(fieldName, field);
                if (!headers.contains("1-测试结果")) {
                    headers.add("1-测试结果");
                } else
                    headers.add("-");
                headerTwos.add(fieldName);
            } else {
                if (!contentHeaders.contains(fieldName)) {
                    contentHeaders.add(fieldName);
                    headers.add("-");
                    headerTwos.add(fieldName);
                    HEADER_COLOR_RECORD.put(fieldName, titleColor);
                }
            }
            if ("-".equals(headers.get(0))) {
                HEADER_COLOR_RECORD.put("测试用例", titleColor);
                headers.set(0, "测试用例");
            }
            map.put("name", fields.getAttribute("name"));
            list.add(map);
        }
    }

    /**
     * Description 查询当前要导入类型的 正确Category
     *
     * @param documentType
     * @throws Exception
     */
    public void parseCurrentCategories(String documentType) throws Exception {
        Document doc = DocumentBuilderFactory.newInstance().newDocumentBuilder()
                .parse(DealService.class.getClassLoader().getResourceAsStream(CATEGORY_CONFIG_FILE));
        Element root = doc.getDocumentElement();
        List<String> typeList = new ArrayList<String>();
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
     * 导出TestSuite对象到 Excel模板
     *
     * @param tObjIds
     * @param cmd
     * @param path
     * @throws Exception
     */
    public void exportReport(List<String> tObjIds, MKSCommand cmd, String path) throws Exception {
        this.parseXML(); // 解析xml
        int resultStep = resultHeaders.size();
        GenerateXmlUtil.caseHeaders.addAll(new ArrayList<String>(contentHeaders));
        int dataIndex = 0;
        List<String> firstHeaders = null;
        List<String> secondHeaders = null;
        List<String> needMoreWidthField = new ArrayList<String>();
        needMoreWidthField.add("Summary");
        needMoreWidthField.add("Expected Results");
        needMoreWidthField.add("Test");
        needMoreWidthField.add("Description");

        /** 获取Category信息 */
        String documentType = "Test Suite";
        parseCurrentCategories(documentType);
        /** 获取Category信息 */
        /** 获取 Pick 值信息 */
        List<String> importFields = new ArrayList<String>();
        for (String header : contentHeaders) {
            String field = HEADER_MAP.get(header);
            if (!"-".equals(field)) {
                importFields.add(field);
            }
        }
        FIELD_TYPE_RECORD.putAll(cmd.getAllFieldType(importFields, PICK_FIELD_RECORD));

        String sheetName = "Test Cases";
        List<String> sheetNames = new ArrayList<>();
        /** 获取 Pick 值信息 */
        HSSFWorkbook wookbook = new HSSFWorkbook();
        HSSFSheet sheet = wookbook.createSheet(sheetName);
        for (String suitId : tObjIds) {
            List<String> caseFields = contentColumns.get(TEST_SUITE);
            caseFields.add("Contains");
            caseFields.add("Test Steps");//合并单元格用
            caseFields.add("Blocked By");//合并单元格用
//			List<Map<String, String>> testCaseItem = cmd.allContents(suitId, caseFields);// 测试用例字段
            List<List<Map<String, String>>> allTestCaseItems = cmd.allContentsByHeading(suitId, caseFields);//测试用例字段
            List<String> testStepFields = new ArrayList<>(testStepfiled);// testStep字段集合
            List<Map<String, String>> list = this.xmlConfig.get(TEST_SUITE);
            //寻找Test Session
            int col = 0;

            Map<String, Map<String, String>> sessionDataMap = new HashMap<>();
            for (List<Map<String, String>> testCaseItems : allTestCaseItems) {
                List<Map<String, Object>> result = cmd.getResult(testCaseItems.get(0).get("ID"), testCaseItems.get(0).get("ID"), "Test Case", null);
                for (Map<String, Object> map : result) {
                    if (!sessionId.contains(String.valueOf(map.get("sessionID")))) {
                        sessionId.add(String.valueOf(map.get("sessionID")));
                        List<Map<String, String>> testSession = cmd.getItemByIds(Arrays.asList(String.valueOf(map.get("sessionID"))), sessionFields);
                        testSession.get(0).put("Test", testCaseItems.get(0).get("Category"));
                        sessionDataMap.put(String.valueOf(map.get("sessionID")), testSession.get(0));
                    }
                }
            }
            if (sessionId.size() == 0) {
                //添加一行空数据
                sessionId.add("");
            }
            for (int i = 0; i < sessionId.size(); i++) {
                //查询Test session 数据
                List<List<Object>> datas = new ArrayList<>();
                List<Object> data = new ArrayList<>();
                if (!sessionDataMap.isEmpty()) {
                    for (String filed : sessionFields) {
                        Map<String, String> stringStringMap = sessionDataMap.get(sessionId.get(i));
                        String s = stringStringMap.get(filed);
                        data.add(s);
                    }
                }
                if (data.isEmpty()) {
                    for (int j = 0; j < sessionFields.size(); j++) {
                        data.add("");
                    }
                }
                datas.add(data);
                List<List<String>> listSession = new ArrayList<>();
                listSession.add(sessionHeader);
                if (i == 0) {
                    GenerateXmlUtil.exportComplexExcel(wookbook, sheet, null, null, needMoreWidthField, sheetName,
                            cellList, col, 1);
                    col++;
                }
                GenerateXmlUtil.exportComplexExcel(wookbook, sheet, listSession, datas, needMoreWidthField, sheetName,
                        cellList, col, 1);
                col++;
            }
            GenerateXmlUtil.exportComplexExcel(wookbook, sheet, null, null, needMoreWidthField, sheetName,
                    cellList, sessionId.size() + 2, 1);
            col++;
            //test session标题占一格
            col++;

            //根据TestSessiond的个数增加标题
            if (sessionId.size() > 1) {
                for (int i = 1; i < sessionId.size(); i++) {
                    for (int j = 0; j < resultHeaders.size(); j++) {
                        headerTwos.add(resultHeaders.get(j));
                        headerTwofiled.add(nameToFiled.get(resultHeaders.get(j)));
                        if (j == 0) {
                            headers.add(i + 1 + "-测试结果");
                        } else {
                            headers.add("-");
                        }
                    }
                    //step添加字段
                    testStepFields.add(String.format("Cycle%s Verdict", i + 1));
                }
            }
            for (int i = 0; i < headers.size(); i++) {
                String headerTwo = headerTwos.get(i);
                CellRangeAddress input = null;
                if ("-".equals(headerTwo)) {//上下合并单元格
                    input = new CellRangeAddress(0, 1, i, i);
                } else if (i < headers.size() - 1) {
                    int temp = 1;
                    String header = headers.get(i + temp);
                    while ("-".equals(header)) {
                        temp++;
                        if (i + temp == headers.size()) {
                            break;
                        }
                        header = headers.get(i + temp);
                    }
                    input = new CellRangeAddress(sessionId.size() + 3, sessionId.size() + 3, i, i + temp - 1);//首行合并单元格
                }
                if (input != null)
                    cellList.add(input);
            }
            int indexCol = col + 2;
            for (List<Map<String, String>> testCaseItems : allTestCaseItems) {//根据一级Heading拆分写入不同Sheet
                //cellList = new ArrayList<>();
                if (!testCaseItems.isEmpty()) {
                    Map<String, String> firstCase = testCaseItems.get(0);
                }

                firstHeaders = new ArrayList<>(headers);
                secondHeaders = new ArrayList<>(headerTwos);
                listHeaders = new ArrayList<>();
                for (Map<String, String> testCase : testCaseItems) {
                    // Test Case 中的字段
                    data = new ArrayList<>(headerTwos.size());// 行数据
                    datas.add(data);// 拼接完所有数据 为一行
                    for (int i = 0; i < headerTwos.size(); i++) {
                        String header = headerTwos.get(i);
                        String value = "";
                        if (contentHeaders.contains(header)) {
                            String realField = HEADER_MAP.get(header);
                            value = realField == null ? ""
                                    : testCase.get(realField) == null ? "" : testCase.get(realField);
                        }
                        data.add(i, value);
                        // 根据Test Step 合并单元格
                    }

                    // ---------Test Step
                    List<String> StepsIDsList = new ArrayList<>();// testStepiD集合
                    String steps = testCase.get("Test Steps");
                    if (steps != null && !"".equals(steps)) { // 如果Test Steps字段不为空  再去查里面的字段。
                        String[] StepsID = steps.split(",");

                        for (int m = 0; m < headerTwos.size(); m++) {// 添加用例数据合并单元格
                            if (!stepHeader.contains(headerTwos.get(m))) {
                                CellRangeAddress input = new CellRangeAddress(indexCol, indexCol + StepsID.length - 1, m, m);
                                cellList.add(input);
                            }
                        }
                        for (int i = 0; i < StepsID.length; i++) {
                            StepsIDsList.add(StepsID[i]);
                            //补充 N个 test Step的行出来用来合并单元格*//*
                            List<Object> stepRowData = new ArrayList<>(headers.size());// 行数据
                            if (i > 0) {//第二个Test Step开始添加行
                                for (int m = 0; m < headers.size(); m++) {
                                    stepRowData.add(m, "");
                                }
                                datas.add(stepRowData);
                            }
                            if (i != 0) {
                                indexCol++;
                            }
                        }
                        // st
                        getStepsItem(cmd, StepsIDsList, testStepFields, dataIndex);// 再次查询Test Steps 中字段。

                    }
                    dataIndex = datas.size();
                    //导出 测试结果数据
                    getTestResult(cmd, testCase, data, headerTwofiled);
                    //导出Defect 数据
                    String defectIdString = testCase.get("Blocked By");
                    if (defectIdString != null && !"".equals(defectIdString)) {
                        String[] defectIds = defectIdString.split(",");
                        getDefect(cmd, Arrays.asList(defectIds), data);
                    }
                    indexCol++;
                }
                cmd.getAllUsers();
                replaceLogid();// logid 替换为FullName(工号)
                listHeaders.add(firstHeaders);// 添加完第一行标题
                listHeaders.add(secondHeaders);// 添加完第二行标题
                //TestResultExportUI.logger.info("reverse is data:" + data);
            }
            GenerateXmlUtil.exportComplexExcel(wookbook, sheet, listHeaders, datas, needMoreWidthField, sheetName,
                    cellList, col, sessionId.size() + 3);
        }

        // 拼接判断文件路径名称
        String documentName = MyRunnable.class.newInstance().documentName;
        SimpleDateFormat df = new SimpleDateFormat("yyyy_MM_dd");// 设置日期格式
        String time = df.format(new Date());
        try {
            String actualPath = path.endsWith(".xls") ? path
                    : path + "\\" + documentName + "_" + time + "_(" + 1 + ")-" + tObjIds.get(0).toString()
                    + ".xls";
            File file = new File(actualPath);
            if (!file.exists()) {
                outputFromwrok(actualPath, wookbook);
            } else {
                int showConfirmDialog = JOptionPane.showConfirmDialog(TestResultExportUI.contentPane,
                        "The file already exists, Whether to overwrite this file?");
                String absolutePath = file.getAbsolutePath();
                if (showConfirmDialog == 0) {// 覆盖
                    outputFromwrok(actualPath, wookbook);
                } else if (showConfirmDialog == 1) {// 不覆盖
                    File pathFile = new File(actualPath);
                    String parent = pathFile.getParent();
                    File fileDir = new File(parent);
                    if (fileDir.isDirectory()) {
                        File[] listFiles = fileDir.listFiles();
                        int count = 1;
                        for (File file2 : listFiles) {
                            String filePath = file2.toString();
                            if (filePath.endsWith("xls") || filePath.endsWith("xlsx")) {
                                if ((filePath.endsWith("-" + tObjIds.get(0).toString() + ".xls")
                                        || filePath.endsWith("-" + tObjIds.get(0).toString() + ".xlsx"))
                                        && filePath.contains(documentName + "_" + time)) {
                                    count++;
                                }
                            }
                        }
                        String actualPath2 = path + "\\" + documentName + "_" + time + "_(" + count + ")-"
                                + tObjIds.get(0).toString() + ".xls";
                        outputFromwrok(actualPath2, wookbook);
                    }

                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * @param cmd
     * @throws APIException
     */
    private void replaceLogid() throws APIException {
        // logid 替换为FullName(工号)
        for (List<Object> list : datas) {
            for (int p = 0; p < list.size(); p++) {
                Object obj = list.get(p);
                String header = headers.get(p);
                if (header == null || "-".equals(header))
                    header = headerTwos.get(p);
                String field = HEADER_MAP.get(header);
                if (field != null) {
                    if ("user".equalsIgnoreCase(FIELD_TYPE_RECORD.get(field))) {
                        String loginId = obj.toString();
                        String fullName = USER_MAP.get(loginId);
                        if (fullName == null || "".equals(fullName))
                            list.set(p, loginId);
                        else
                            list.set(p, fullName + "(" + loginId + ")");
                    }
                }
            }
        }
    }

    /**
     * 获取数据中下标
     *
     * @param s
     * @param v
     * @return
     */
    private List<Integer> get(List<String> s, String v) {
        List<Integer> index = new ArrayList<>();
        for (int i = 0; i < s.size(); i++) {
            if (v.equals(s.get(i))) {
                index.add(i);
            }
        }
        return index;
    }


    /**
     * 测试结果处理方法
     *
     * @param cmd
     * @param testCase
     * @throws APIException
     */
    private boolean getTestResult(MKSCommand cmd, Map<String, String> testCase, List<Object> data, List<String> headerTwofiled)
            throws APIException {
        List<Map<String, Object>> result = cmd.getResult(testCase.get("ID"), testCase.get("ID"), "Test Case", testResultFiled);
        if (result != null && result.size() > 0) {
            for (Map<String, Object> map : result) {
                String sessionID = String.valueOf(map.get("sessionID"));
                for (int i = 0; i < sessionId.size(); i++) {
                    if (sessionId.get(i).equals(sessionID)) {
                        for (int j = 0; j < testResultFiled.size(); j++) {
                            List<Integer> v = get(headerTwofiled, testResultFiled.get(j));
                            if (!v.isEmpty()) {
                                data.set(v.get(i), map.get(testResultFiled.get(j)));
                            }
                        }
                    }
                }
            }
        }
        return false;
    }

    private void getDefect(MKSCommand cmd, List<String> ids, List<Object> data) throws APIException {
        List<Map<String, String>> result = cmd.getItemByIds(ids, defectFiled);
        if (result != null && result.size() > 0) {
            for (Map<String, String> map : result) {
                for (int i = 0; i < 1; i++) {
                    for (int j = 0; j < defectFiled.size(); j++) {
                        List<Integer> v = get(headerTwofiled, defectFiled.get(j));
                        if (!v.isEmpty()) {
                            if (v.get(i) == 0) {
                                data.set(v.get(i + 2), map.get(defectFiled.get(j)));
                            } else {
                                data.set(v.get(i), map.get(defectFiled.get(j)));
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * 获取Test Steps中的字段
     *
     * @param cmd
     * @param StepsIDsList
     * @param testSteps
     * @param startIndex   。 当前数据在Test Case 的第几条中
     * @throws APIException
     */
    private void getStepsItem(MKSCommand cmd, List<String> StepsIDsList, List<String> testSteps, Integer startIndex
    ) throws APIException {
        List<Map<String, String>> itemMaps = cmd.getItemByIds(StepsIDsList, testSteps);
        List<String> stepsData = new ArrayList<>();
        for (int count = 0; count < itemMaps.size(); count++) {
            Map<String, String> map = itemMaps.get(count);
            List<Object> rowData = datas.get(startIndex + count);
            for (int i = 0; i < testStepfiled.size(); i++) {
                List<Integer> v = get(headerTwofiled, testStepfiled.get(i));
                if ("ID".equals(testStepfiled.get(i))) {
                    v = Arrays.asList(v.get(1));
                }
                if (!v.isEmpty()) {
                    if (v.size() > 1) {
                        for (int j = 0; j < v.size(); j++) {
                            rowData.set(v.get(j), map.get(testSteps.get(j + 3)));
                        }
                    } else {
                        rowData.set(v.get(0), map.get(testStepfiled.get(i)));
                    }
                }
            }
            /*List<Object> rowData = datas.get(startIndex + count);
            for (String header : stepHeader) {
                String realField = stepHeaderMap.get(header);
                String val = realField == null ? "" : map.get(realField) == null ? "" : map.get(realField);
                if (count == 0) {
                    rowData.set(headerTwos.indexOf(header), val);//封装数据
                } else {
                    rowData.set(headerTwos.indexOf(header), val);//封装数据
                }
            }*/
        }
    }

    public static void outputFromwrok(String filePath, Workbook wookbook) {
        try {
            FileOutputStream output = new FileOutputStream(filePath);
            wookbook.write(output);
            output.flush();
            TestResultExportUI.class.newInstance().isParseSuccess = true;

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    /**
     * 设置Border
     *
     * @param style
     * @param top
     * @param bottom
     * @param left
     * @param right
     * @param border
     */
    public static void setBorder(HSSFCellStyle style, boolean top, boolean bottom, boolean left, boolean right,
                                 short border) {
        if (top)
            style.setBorderTop(border);
        if (bottom)
            style.setBorderBottom(border);
        if (left)
            style.setBorderLeft(border);
        if (right)
            style.setBorderRight(border);
    }

}
