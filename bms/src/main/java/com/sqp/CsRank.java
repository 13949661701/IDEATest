package com.sqp;

import java.io.Serializable;

/**
 * 催收风险数据统计Bean
 * @author crfchina
 * 上海信而富企业管理有限公司
 */
public class CsRank implements Serializable {

	private static final long serialVersionUID = 8120920089698231866L;
	//  自增ID
	private int id;
	//  日期
	private String observe_day;
	//  模型分组（0：当前逾期天数为0的客户、1:当前逾期天数=1天、2：当前逾期天数2-6天、3：当前逾期天数7-14天、4：当前逾期天数15-89天、5：当前逾期天数>=90天）
	private String group;
	//  风险等级（0:无逾期、1:高风险、2:中风险、3:低风险、-1:当前逾期天数超过90天的）
	private String cs_rank;
	//  每个模型风险分组人数
	private int cs_count;
	//  催收策略编码
	private String csStrategyCode;
	/**
	 * @return the id
	 */
	public int getId() {
		return id;
	}
	/**
	 * @param id the id to set
	 */
	public void setId(int id) {
		this.id = id;
	}
	/**
	 * @return the observe_day
	 */
	public String getObserve_day() {
		return observe_day;
	}
	/**
	 * @param observe_day the observe_day to set
	 */
	public void setObserve_day(String observe_day) {
		this.observe_day = observe_day;
	}
	/**
	 * @return the group
	 */
	public String getGroup() {
		return group;
	}
	/**
	 * @param group the group to set
	 */
	public void setGroup(String group) {
		this.group = group;
	}
	/**
	 * @return the cs_rank
	 */
	public String getCs_rank() {
		return cs_rank;
	}
	/**
	 * @param cs_rank the cs_rank to set
	 */
	public void setCs_rank(String cs_rank) {
		this.cs_rank = cs_rank;
	}
	/**
	 * @return the cs_count
	 */
	public int getCs_count() {
		return cs_count;
	}
	/**
	 * @param cs_count the cs_count to set
	 */
	public void setCs_count(int cs_count) {
		this.cs_count = cs_count;
	}
	/**
	 * @return the csStrategyCode
	 */
	public String getCsStrategyCode() {
		return csStrategyCode;
	}
	/**
	 * @param csStrategyCode the csStrategyCode to set
	 */
	public void setCsStrategyCode(String csStrategyCode) {
		this.csStrategyCode = csStrategyCode;
	}
	
	

}
