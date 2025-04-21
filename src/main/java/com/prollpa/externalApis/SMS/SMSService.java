package com.prollpa.externalApis.SMS;

import java.util.List;
import org.springframework.http.MediaType;

import javax.xml.transform.TransformerConfigurationException;

import org.apache.coyote.BadRequestException;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.springframework.web.client.RestTemplate;
import org.springframework.web.util.UriComponentsBuilder;

import com.prollpa.exception.ResourceNotFoundException;
import com.prollpa.payload.SMSResponse;


@Service
public class SMSService {
    private String apiUrl="https://api.optimummeasures.com/api/SMS";

    private String userName="sv_sudan";

    
    private String password="evugrj7y";

    private String senderId="VFS";
    private RestTemplate restTemplate;
    public SMSService(RestTemplate restTemplate) {
    	this.restTemplate=restTemplate;
    }
//india
//    public String sendSms(String phoneNumber, String message) {
//    	try {
//            // Construct URL
//            String url = UriComponentsBuilder.fromHttpUrl(apiUrl)
//                    .queryParam("to", phoneNumber)
//                    .queryParam("from", senderId)
//                    .queryParam("message", message)
//                    .toUriString();
//
//            // Add Basic Authentication Header
//            HttpHeaders headers = new HttpHeaders();
//            headers.setBasicAuth(username, password);
//
//            HttpEntity<Void> entity = new HttpEntity<>(headers);
//
//            // Send Request
//            ResponseEntity<String> response = restTemplate.getForEntity(url, String.class, entity);
//
//            return response.getBody();
//
//        } catch (Exception e) {
//            e.printStackTrace();
//            return "Failed to send SMS: " + e.getMessage();
//        } 
//    }
    

    public String sendSms(String mobileNumber,String message) {
    	 String payload = "{"
    	            + "\"Destination\": ["
    	            + "{\"MobileNo\": \"" +"249"+ mobileNumber + "\"},"
    	            + "],"
    	            + "\"Message\": \"" + message + "\","
    	            + "\"Sender\": \""+senderId+"\","
    	            + "\"ReferenceID\": \"ID09890\","
    	            + "\"CallBack\": \"Yes\""
    	            + "}";
    	 

    	        // Set the parameters for authentication in the URL
    	        String url = apiUrl + "?UserName=" + userName + "&Password=" + password;

    	        // Set headers
    	        HttpHeaders headers = new HttpHeaders();
    	        headers.setContentType(MediaType.APPLICATION_JSON);

    	        // Create the request entity with the payload
    	        HttpEntity<String> entity = new HttpEntity<>(payload, headers);

    	        // Create a RestTemplate instance to make the API call
    	        RestTemplate restTemplate = new RestTemplate();

    	        // Send the POST request to the API with authentication parameters in the URL
    	        ResponseEntity<String> response = restTemplate.exchange(
    	                url,
    	                HttpMethod.POST,
    	                entity,
    	                String.class
    	        );

    	        // Print the response from the API
    	        System.out.println("Response: " + response.getBody());
    	        return response.toString();
    	    }
    
    public ResponseEntity<String> getSmsStatus(String id) {
        // Construct the URL with ID, UserName, and Password
        String url = apiUrl + "/GetStatus?ID=" + id + "&UserName=" + userName + "&Password=" + password;

        // Set headers
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_JSON);

        // Create the request entity
        HttpEntity<Void> entity = new HttpEntity<>(headers);

        // Use RestTemplate to make the GET request
        restTemplate = new RestTemplate();
        ResponseEntity<String> response ;
        // Send the GET request
        try {
        	response= restTemplate.exchange(
                url,
                HttpMethod.GET,
                entity,
                String.class
        );
        }catch(Exception e) {
        	throw new ResourceNotFoundException(e.getMessage());
        }

        // Print the response
        
        System.out.println("Response: " + response.getClass());

        // Return the response
        return response;
    }
}

