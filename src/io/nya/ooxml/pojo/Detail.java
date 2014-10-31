package io.nya.ooxml.pojo;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

@JsonIgnoreProperties(ignoreUnknown=true)
public class Detail {
	@JsonProperty
	public String sn;
	@JsonProperty
	public long match;
	@JsonProperty
	public long upload_time;
	@JsonProperty
	public long date;
	@JsonProperty
	public String isOnline;
	@JsonProperty
	public String device;
	@JsonProperty
	public String orderCompany;
	@JsonProperty
	public String province;
	@JsonProperty
	public String region;
	@JsonProperty
	public String nd_pd_online = "";
	@JsonProperty
	public String nd_pd_orderCompany = "";
	@JsonProperty
	public String nd_pd_province = "";
	@JsonProperty
	public String nd_pd_region = "";
}
