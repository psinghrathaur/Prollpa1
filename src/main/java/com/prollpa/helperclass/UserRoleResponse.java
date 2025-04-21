package com.prollpa.helperclass;

import java.util.List;
import java.util.Map;

public class UserRoleResponse {
	
	private String userName;
	private Map<VSC,List<UserRole>> userVSCList;
	public UserRoleResponse(String userName, Map<VSC, List<UserRole>> userVSCList) {
		super();
		this.userName = userName;
		this.userVSCList = userVSCList;
	}
	public String getUserName() {
		return userName;
	}
	public void setUserName(String userName) {
		this.userName = userName;
	}
	public Map<VSC, List<UserRole>> getUserVSCList() {
		return userVSCList;
	}
	public void setUserVSCList(Map<VSC, List<UserRole>> userVSCList) {
		this.userVSCList = userVSCList;
	}
	public UserRoleResponse() {
		
	}

}
