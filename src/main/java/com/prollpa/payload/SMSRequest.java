package com.prollpa.payload;

import java.util.List;

import lombok.Data;

@Data
public class SMSRequest {
    private String message;
    private String mobileNumber;
}
