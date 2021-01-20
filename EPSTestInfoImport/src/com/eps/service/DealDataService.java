package com.eps.service;

import com.eps.util.IntegrityUtil;
import com.mks.api.response.APIException;

import java.io.File;
import java.util.List;
import java.util.Map;

/**
 * Description: 处理数据接口
 * @author @ModifyDate
 * Yi Gang  2020/12/17
 */
public interface DealDataService {

	/**
	 * 利用Jsoup解析配置文件，得到相应的参数，为Type选项和创建Document提供信息 (1)
	 * Document:Type,Project,State,Shared Category (2) Content:Type 负责人：汪巍
	 * 
	 * @return
	 * @throws Exception 
	 */
	List<String> parsFieldMapping() throws Exception;

	/**
	 * Description 查询当前要导入类型的 正确Category
	 * @param documentType
	 * @throws Exception
	 */
	void parseCurrentCategories(String documentType) throws Exception;

	/**
	 * Description 处理数据，并校验
	 * @param data
	 * @param importType
	 * @param cmd
	 * @return
	 * @throws Exception 
	 */
	List<List<Map<String, Object>>> checkExcelData(List<List<Map<String, Object>>> datas,
			Map<String, String> errorRecord, IntegrityUtil cmd) throws Exception;

	/**
	 * 处理Excel中的数据，将Test Step信息和Test Case信息拆分开
	 * 
	 * @param data
	 * @return
	 */
	List<List<Map<String, Object>>> dealExcelData(List<List<Map<String, Object>>> datas);

	/**
	 * Description 开始导入数据
	 * @param data
	 * @param cmd
	 * @param importType
	 * @param shortTitle
	 * @param project
	 * @param testSuiteID
	 * @throws Exception
	 */
	void startImport(List<List<Map<String, Object>>> datas, IntegrityUtil cmd, String importType, String shortTitle,
			String project, String testSuiteID) throws APIException;

	/**
	 * 获得Excel中的数据
	 * 
	 * @param filePath
	 * @return
	 * @throws BiffException
	 * @throws IOException
	 */
	List<List<Map<String, Object>>> parseExcel(File file) throws Exception;

}
