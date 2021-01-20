package com.eps.service;

import com.eps.ui.TestInfoImportUI;
import com.eps.util.APIExceptionUtil;
import com.eps.util.IntegrityUtil;
import com.mks.api.response.APIException;

import javax.swing.*;
import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * Description: 线程 异步处理处理
 * @author @ModifyDate
 * Yi Gang  2020/12/18
 */
public class MyRunnable implements Runnable {
	public IntegrityUtil cmd;
	public DealDataService dealDataService;
	public String importType = "Test Case";
	public String testSuiteId;
	public List<List<Map<String,Object>>> datas;
	public String project;
	public String shortTitle;
	public MyRunnable() {
		super();
	}

	@Override
	public void run() {
		try {
			TestInfoImportUI.logger.info("===============Start to import Test Case==============");
			dealDataService.startImport(datas, cmd, importType,shortTitle,project, testSuiteId);
			JOptionPane.showMessageDialog(TestInfoImportUI.contentPane, "Done", "Success", JOptionPane.INFORMATION_MESSAGE);
		} catch (APIException e) {
			TestInfoImportUI.logger.error(APIExceptionUtil.getMsg(e));
			JOptionPane.showMessageDialog(TestInfoImportUI.contentPane, e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
		} catch (Exception e) {
			JOptionPane.showMessageDialog(TestInfoImportUI.contentPane, e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
		} finally {
			try {
				cmd.release();
			} catch (IOException e) {
				
			}
			TestInfoImportUI.logger.info("===============End to import Test Case==============");
		}
	}

	

}
