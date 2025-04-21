package com.prollpa.helperclass;

public class User {
	private long userId;
	private String userName;
	public User() {
		
	}
	public User(long userId, String userName) {
		super();
		this.userId = userId;
		this.userName = userName;
	}
	public long getUserId() {
		return userId;
	}
	public void setUserId(long userId) {
		this.userId = userId;
	}
	public String getUserName() {
		return userName;
	}
	public void setUserName(String userName) {
		this.userName = userName;
	}

}
