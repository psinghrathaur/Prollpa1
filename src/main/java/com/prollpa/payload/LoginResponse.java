package com.prollpa.payload;

import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class LoginResponse {
	private String jwtToken;
	private String validateTime;
	private String authenticationStatus;
}
