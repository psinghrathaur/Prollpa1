package com.prollpa.config;

import org.modelmapper.ModelMapper;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.stereotype.Component;
import org.springframework.web.client.RestTemplate;

@Configuration
public class MyConfig {
    @Bean
	public ModelMapper getModelMapper() {
		return new ModelMapper();
	}
    @Bean
    public RestTemplate restTemplate() {
        return new RestTemplate();
    }
}
