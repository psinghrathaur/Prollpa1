package com.prollpa.payload;

import jakarta.validation.constraints.NotNull;
import lombok.Data;
import lombok.Getter;
import lombok.Setter;
@Data
public class LoginDto {
	@NotNull(message="username can not be null")
	private String username;
	@NotNull(message="username can not be null")
	private String password;
	

}
