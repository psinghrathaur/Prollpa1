package com.prollpa.service;

import java.util.List;

import com.prollpa.entity.UserVscRole;
import com.prollpa.payload.UserVscRoleDto;
import com.prollpa.payload.UserVscRoleResponse;

public interface UserVscRoleService {
	List<UserVscRoleDto> saveUserVscRoleList(List<UserVscRoleDto> userVscRole);
	List<UserVscRoleDto> getUserVscRoleByUserId(long userId);
	List<UserVscRoleDto> getUserVscRoleByVscId(long vscId);
	List<UserVscRoleResponse> getUserVscRoleByUserId2(List<UserVscRoleDto> userVscRoleList);
	

}
