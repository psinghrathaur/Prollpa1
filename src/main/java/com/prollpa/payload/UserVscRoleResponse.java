package com.prollpa.payload;

import java.util.List;

import com.prollpa.entity.Role;
import com.prollpa.entity.UserVscRole;
import com.prollpa.entity.VSC;

import lombok.Data;
@Data
public class UserVscRoleResponse {
	private Long vscId;
    private String vscCenterName;
    private List<RoleDto> role;
}
