package com.prollpa.exception;

import org.springframework.http.HttpStatus;

import lombok.Builder;
import lombok.Data;
@Data
@Builder
public class ExceptionApiResponse {
	private String message;
	private HttpStatus status;
	private String description;

}
