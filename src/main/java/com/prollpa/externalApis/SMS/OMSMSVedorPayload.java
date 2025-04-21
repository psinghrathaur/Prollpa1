package com.prollpa.externalApis.SMS;

import lombok.Data;

@Data
public class OMSMSVedorPayload {
	
	private Destination[] destination;
    private String message;
    private String sender;
    private String referenceID;
    private String callBack;
    
    public static class Destination {
        private String mobileNo;

        public Destination(String mobileNo) {
            this.mobileNo = mobileNo;
        }

        public String getMobileNo() {
            return mobileNo;
        }

        public void setMobileNo(String mobileNo) {
            this.mobileNo = mobileNo;
        }
    }

}
