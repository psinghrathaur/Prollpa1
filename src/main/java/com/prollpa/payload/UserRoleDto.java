package com.prollpa.payload;

import jakarta.validation.constraints.NotNull;
import lombok.Data;
@Data
public class UserRoleDto {
	private Long userRoleId;
	@NotNull(message = "role id cannot be null")
    private Long roleId;
	@NotNull(message = "user id cannot be null")
	private Long userId;
	

}
