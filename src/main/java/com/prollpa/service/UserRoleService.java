package com.prollpa.service;

import java.util.List;



import com.prollpa.payload.UserRoleDto;

public interface UserRoleService{
	UserRoleDto saveUserRole(UserRoleDto userDto);
	List<UserRoleDto> getUserRoleByUserId(Long  UserId);
	
	

}
