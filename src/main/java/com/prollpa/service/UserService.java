package com.prollpa.service;

import com.prollpa.payload.UserDto;

public interface UserService {
	UserDto saveUser(UserDto dto);
	UserDto getUserByUsername(String username);
	UserDto getUserByUserId(long userId);
}

