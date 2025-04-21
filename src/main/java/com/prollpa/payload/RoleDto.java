package com.prollpa.payload;

import jakarta.validation.constraints.NotEmpty;
import lombok.Data;

@Data
public class RoleDto {
	private Long roleId;
	@NotEmpty(message = "roll can not be null")
	private String role;

}
