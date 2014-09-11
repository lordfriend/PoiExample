package io.nya.ooxml.pojo;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

@JsonIgnoreProperties(ignoreUnknown=true)
public class Online {
	@JsonProperty
	public long earliest;
	@JsonProperty
	public long latest;
	@JsonProperty
	public String export_id;
	@JsonProperty
	public int total;
	@JsonProperty
	public OnlineData[] content;
	
	@JsonIgnoreProperties(ignoreUnknown=true)
	public static class OnlineData {
		@JsonProperty
		public String _id;
		@JsonProperty
		public int total;
	}
}
