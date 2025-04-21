package com.prollpa.payload;

import jakarta.validation.constraints.NotEmpty;
import jakarta.validation.constraints.NotNull;

import lombok.Data;

@Data
public class UserVscRoleDto {
	private Long UserVscRoleId;
	 @NotNull(message = "vsc id null")
	    private Long vscId;  // Use Long instead of long
	     
	    @NotNull(message = "role id null")
	    private Long roleId;  // Use Long instead of long
	     
	    @NotNull(message = "userId id null")
	    private Long userId;
}
