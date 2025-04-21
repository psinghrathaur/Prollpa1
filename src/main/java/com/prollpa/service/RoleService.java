package com.prollpa.service;

import java.util.List;

import com.prollpa.entity.Role;
import com.prollpa.payload.RoleDto;

public interface RoleService {
	public Role saveRole(Role role);
	public List<Role> getRoleList();
	public Role getRoleByRoleId(long roleId);
	public  RoleDto mapRoletoDto(Role role);
	public Role getRoleByRoleId(Long roleId);
	

}
